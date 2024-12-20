VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmConfiguration 
   Caption         =   "Configuration"
   ClientHeight    =   10815
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
   ScaleHeight     =   10815
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Activate 
      Caption         =   "Activate"
      Height          =   735
      Left            =   13680
      TabIndex        =   610
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txtAO_Type 
      Height          =   390
      Left            =   14160
      TabIndex        =   560
      Text            =   "0"
      Top             =   5640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer tmrTest 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   14520
      Top             =   7680
   End
   Begin VB.Timer tmrTimeoutCT 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   14520
      Top             =   8160
   End
   Begin VB.TextBox txtErrorV 
      Height          =   390
      Left            =   13440
      TabIndex        =   475
      Text            =   "0"
      Top             =   5640
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtAO 
      Height          =   390
      Left            =   13440
      TabIndex        =   474
      Text            =   "0"
      Top             =   6120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer tmrSendCT 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   14040
      Top             =   8160
   End
   Begin VB.TextBox txtAlarm 
      Height          =   390
      Left            =   13440
      TabIndex        =   432
      Text            =   "0"
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer tmrFinishedLight 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   13440
      Top             =   8640
   End
   Begin VB.Timer tmrFinishedBeep 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   13440
      Top             =   8160
   End
   Begin VB.TextBox txtErrorMap 
      Height          =   390
      Left            =   13440
      TabIndex        =   260
      Text            =   "0"
      Top             =   6600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox cmbIOList 
      Height          =   390
      Index           =   4
      Left            =   13560
      TabIndex        =   237
      Text            =   "NA"
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Timer tmrDIO 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   13440
      Top             =   7680
   End
   Begin VB.ComboBox cmbIOList 
      Height          =   390
      Index           =   3
      Left            =   13560
      TabIndex        =   106
      Text            =   "NA"
      Top             =   4440
      Width           =   2295
   End
   Begin VB.ComboBox cmbIOList 
      Height          =   390
      Index           =   2
      Left            =   13560
      TabIndex        =   105
      Text            =   "NA"
      Top             =   4080
      Width           =   2295
   End
   Begin VB.ComboBox cmbIOList 
      Height          =   390
      Index           =   1
      Left            =   13560
      TabIndex        =   104
      Text            =   "NA"
      Top             =   3720
      Width           =   2295
   End
   Begin VB.ComboBox cmbIOList 
      Height          =   390
      Index           =   0
      ItemData        =   "frmConfiguration.frx":0000
      Left            =   13560
      List            =   "frmConfiguration.frx":0002
      TabIndex        =   103
      Text            =   "NA"
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Timer tmrWatchDog 
      Interval        =   1000
      Left            =   14040
      Top             =   8640
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   735
      Left            =   13680
      TabIndex        =   20
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   735
      Left            =   13680
      TabIndex        =   19
      Top             =   720
      Width           =   1455
   End
   Begin TabDlg.SSTab tabConfiguration 
      Height          =   10575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   18653
      _Version        =   393216
      Tabs            =   7
      Tab             =   2
      TabsPerRow      =   7
      TabHeight       =   1058
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "機台設置"
      TabPicture(0)   =   "frmConfiguration.frx":0004
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tabMain"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "DIO"
      TabPicture(1)   =   "frmConfiguration.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDO"
      Tab(1).Control(1)=   "fraDI"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "AIO"
      TabPicture(2)   =   "frmConfiguration.frx":003C
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fraAI"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraAO"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Parameter Advance I"
      TabPicture(3)   =   "frmConfiguration.frx":0058
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame7"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame5"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Frame2"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "fraSmoothCurve"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Parameter Advance II"
      TabPicture(4)   =   "frmConfiguration.frx":0074
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame9"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame10"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Frame11"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Parameter Advance III"
      TabPicture(5)   =   "frmConfiguration.frx":0090
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame8"
      Tab(5).Control(1)=   "fraUniformity"
      Tab(5).Control(2)=   "Frame14"
      Tab(5).Control(3)=   "Frame21"
      Tab(5).Control(4)=   "Frame23"
      Tab(5).Control(5)=   "Frame24"
      Tab(5).Control(6)=   "Frame25"
      Tab(5).ControlCount=   7
      TabCaption(6)   =   "TC Module"
      TabPicture(6)   =   "frmConfiguration.frx":00AC
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "fraAdvDaqAI"
      Tab(6).Control(1)=   "Frame13"
      Tab(6).ControlCount=   2
      Begin VB.Frame Frame25 
         Caption         =   "燈管設定"
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
         Left            =   -71520
         TabIndex        =   601
         Top             =   5640
         Width           =   4695
         Begin VB.ComboBox cmbOrder3 
            Height          =   390
            Index           =   0
            ItemData        =   "frmConfiguration.frx":00C8
            Left            =   120
            List            =   "frmConfiguration.frx":00D8
            TabIndex        =   623
            Top             =   2160
            Width           =   630
         End
         Begin VB.ComboBox cmbOrder1 
            Height          =   390
            Index           =   0
            ItemData        =   "frmConfiguration.frx":00E8
            Left            =   120
            List            =   "frmConfiguration.frx":00F8
            TabIndex        =   622
            Top             =   1560
            Width           =   630
         End
         Begin VB.ComboBox cmbOrder4 
            Height          =   390
            Index           =   0
            ItemData        =   "frmConfiguration.frx":0108
            Left            =   2400
            List            =   "frmConfiguration.frx":0118
            TabIndex        =   621
            Top             =   2160
            Width           =   615
         End
         Begin VB.ComboBox cmbOrder2 
            Height          =   390
            Index           =   0
            ItemData        =   "frmConfiguration.frx":0128
            Left            =   2400
            List            =   "frmConfiguration.frx":0138
            TabIndex        =   620
            Top             =   1560
            Width           =   630
         End
         Begin VB.ComboBox cmbCTName4 
            Height          =   390
            Index           =   0
            ItemData        =   "frmConfiguration.frx":0148
            Left            =   3000
            List            =   "frmConfiguration.frx":0158
            TabIndex        =   614
            Top             =   2160
            Width           =   735
         End
         Begin VB.ComboBox cmbCTName2 
            Height          =   390
            Index           =   0
            ItemData        =   "frmConfiguration.frx":016A
            Left            =   3000
            List            =   "frmConfiguration.frx":017A
            TabIndex        =   613
            Top             =   1560
            Width           =   735
         End
         Begin VB.ComboBox cmbCTName3 
            Height          =   390
            Index           =   0
            ItemData        =   "frmConfiguration.frx":018C
            Left            =   720
            List            =   "frmConfiguration.frx":019C
            TabIndex        =   612
            Top             =   2160
            Width           =   735
         End
         Begin VB.ComboBox cmbCTName1 
            Height          =   390
            Index           =   0
            ItemData        =   "frmConfiguration.frx":01AE
            Left            =   720
            List            =   "frmConfiguration.frx":01BE
            TabIndex        =   611
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txtNumber4 
            Height          =   390
            Left            =   3720
            TabIndex        =   609
            Top             =   2160
            Width           =   855
         End
         Begin VB.TextBox txtNumber2 
            Height          =   390
            Left            =   3720
            TabIndex        =   608
            Top             =   1560
            Width           =   855
         End
         Begin VB.TextBox txtNumber3 
            Height          =   390
            Left            =   1440
            TabIndex        =   607
            Top             =   2160
            Width           =   855
         End
         Begin VB.TextBox txtNumber1 
            Height          =   390
            Left            =   1440
            TabIndex        =   606
            Top             =   1560
            Width           =   855
         End
         Begin VB.TextBox txtCTNumber 
            Height          =   390
            Left            =   2880
            TabIndex        =   605
            Top             =   480
            Width           =   855
         End
         Begin VB.CheckBox ckeCTDisplay 
            Caption         =   "燈管表格顯示"
            Height          =   270
            Left            =   240
            TabIndex        =   603
            Top             =   960
            Width           =   1815
         End
         Begin VB.CheckBox ckeForcePreheat 
            Caption         =   "強制預熱"
            Height          =   270
            Left            =   240
            TabIndex        =   602
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "燈管數量："
            Height          =   375
            Index           =   0
            Left            =   1680
            TabIndex        =   604
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Frame Frame24 
         Caption         =   "Robot (設定完要重開):"
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
         Height          =   1215
         Left            =   -71520
         TabIndex        =   549
         Top             =   4320
         Width           =   4695
         Begin VB.TextBox txtRobotIP 
            Height          =   390
            Left            =   720
            TabIndex        =   556
            Text            =   "192.168.0.11"
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "IP:"
            Height          =   270
            Index           =   205
            Left            =   120
            TabIndex        =   557
            Top             =   480
            Width           =   270
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "TCM (設定完要重開):"
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
         Left            =   -71520
         TabIndex        =   548
         Top             =   2280
         Width           =   4695
         Begin VB.CommandButton cmdAz1 
            Caption         =   "Az1"
            Height          =   615
            Left            =   3240
            TabIndex        =   553
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtAzIP1 
            Height          =   390
            Left            =   720
            TabIndex        =   552
            Text            =   "192.168.0.11"
            Top             =   480
            Width           =   2295
         End
         Begin VB.TextBox txtAzIP2 
            Height          =   390
            Left            =   720
            TabIndex        =   551
            Text            =   "192.168.0.12"
            Top             =   1080
            Width           =   2295
         End
         Begin VB.CommandButton cmdAz2 
            Caption         =   "Az2"
            Height          =   615
            Left            =   3240
            TabIndex        =   550
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "IP1:"
            Height          =   270
            Index           =   194
            Left            =   120
            TabIndex        =   555
            Top             =   480
            Width           =   405
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "IP2:"
            Height          =   270
            Index           =   195
            Left            =   120
            TabIndex        =   554
            Top             =   1200
            Width           =   405
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "PM (設定完要重開):"
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
         Height          =   1215
         Left            =   -71520
         TabIndex        =   528
         Top             =   840
         Width           =   4695
         Begin VB.TextBox txtPMbig 
            Height          =   390
            Left            =   840
            TabIndex        =   530
            Text            =   "1"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtPMsmall 
            Height          =   390
            Left            =   2880
            TabIndex        =   529
            Text            =   "1"
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Big:"
            Height          =   270
            Index           =   198
            Left            =   240
            TabIndex        =   532
            Top             =   480
            Width           =   420
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Small:"
            Height          =   270
            Index           =   197
            Left            =   1920
            TabIndex        =   531
            Top             =   480
            Width           =   645
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Type List"
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
         Height          =   6615
         Left            =   -74760
         TabIndex        =   485
         Top             =   2160
         Width           =   3135
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "1=舊溫卡+舊DIO卡"
            Height          =   270
            Index           =   186
            Left            =   240
            TabIndex        =   492
            Top             =   480
            Width           =   2010
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "2=舊溫卡(RTA50)"
            Height          =   270
            Index           =   185
            Left            =   240
            TabIndex        =   491
            Top             =   840
            Width           =   1875
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "3=新溫卡(USB)+舊DIO卡"
            Height          =   270
            Index           =   184
            Left            =   240
            TabIndex        =   490
            Top             =   1200
            Width           =   2655
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "5=新溫卡(USB)+新DIO卡"
            Height          =   270
            Index           =   180
            Left            =   240
            TabIndex        =   489
            Top             =   1560
            Width           =   2655
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "6=舊溫卡+新DIO卡"
            Height          =   270
            Index           =   181
            Left            =   240
            TabIndex        =   488
            Top             =   1920
            Width           =   2010
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "7=新溫卡(TCP)+舊DIO卡"
            Height          =   270
            Index           =   182
            Left            =   240
            TabIndex        =   487
            Top             =   2280
            Width           =   2640
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "8=新溫卡(TCP)+新DIO卡"
            Height          =   270
            Index           =   183
            Left            =   240
            TabIndex        =   486
            Top             =   2640
            Width           =   2640
         End
      End
      Begin TabDlg.SSTab tabMain 
         Height          =   8295
         Left            =   -74880
         TabIndex        =   301
         Top             =   720
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   14631
         _Version        =   393216
         Tab             =   2
         TabHeight       =   882
         TabCaption(0)   =   "一般設置"
         TabPicture(0)   =   "frmConfiguration.frx":01D0
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame1"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Robot"
         TabPicture(1)   =   "frmConfiguration.frx":01EC
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Command15"
         Tab(1).Control(1)=   "Command14"
         Tab(1).Control(2)=   "Command13"
         Tab(1).Control(3)=   "txtPlaceH"
         Tab(1).Control(4)=   "txtPickH"
         Tab(1).Control(5)=   "Frame20"
         Tab(1).Control(6)=   "txtCurrPos(3)"
         Tab(1).Control(7)=   "txtCurrPos(2)"
         Tab(1).Control(8)=   "txtCurrPos(1)"
         Tab(1).Control(9)=   "txtWriteTeach(3)"
         Tab(1).Control(10)=   "txtWriteTeach(2)"
         Tab(1).Control(11)=   "txtWriteTeach(1)"
         Tab(1).Control(12)=   "Command12"
         Tab(1).Control(13)=   "Command11"
         Tab(1).Control(14)=   "txtCurrPos(0)"
         Tab(1).Control(15)=   "Text8"
         Tab(1).Control(16)=   "Command10"
         Tab(1).Control(17)=   "Command5"
         Tab(1).Control(18)=   "txtWriteTeach(0)"
         Tab(1).Control(19)=   "txtTeachIndex"
         Tab(1).Control(20)=   "Text2"
         Tab(1).Control(21)=   "Text1"
         Tab(1).Control(22)=   "Frame19"
         Tab(1).Control(23)=   "Frame18"
         Tab(1).Control(24)=   "txtRobotPort"
         Tab(1).Control(25)=   "fgTeach"
         Tab(1).Control(26)=   "lbName(168)"
         Tab(1).Control(27)=   "lbName(167)"
         Tab(1).Control(28)=   "lbName(166)"
         Tab(1).Control(29)=   "lbName(165)"
         Tab(1).Control(30)=   "lbName(164)"
         Tab(1).ControlCount=   31
         TabCaption(2)   =   "警報設置"
         TabPicture(2)   =   "frmConfiguration.frx":0208
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Label2"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "fgAlarm"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).ControlCount=   2
         Begin VB.CommandButton Command15 
            Caption         =   "Write All"
            Height          =   495
            Left            =   -67080
            TabIndex        =   406
            Top             =   1800
            Width           =   1095
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Read All"
            Height          =   495
            Left            =   -68280
            TabIndex        =   401
            Top             =   1800
            Width           =   1095
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Move"
            Height          =   495
            Left            =   -69480
            TabIndex        =   400
            Top             =   1800
            Width           =   1095
         End
         Begin VB.TextBox txtPlaceH 
            Height          =   390
            Left            =   -73320
            TabIndex        =   395
            Text            =   "1"
            Top             =   1800
            Width           =   855
         End
         Begin VB.TextBox txtPickH 
            Height          =   390
            Left            =   -73320
            TabIndex        =   393
            Text            =   "1"
            Top             =   1320
            Width           =   855
         End
         Begin VB.Frame Frame20 
            Caption         =   "Speed"
            Height          =   2775
            Left            =   -74760
            TabIndex        =   391
            Top             =   2280
            Width           =   1575
            Begin VB.OptionButton optRobotSpeed 
               Caption         =   "4"
               Height          =   495
               Index           =   3
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   399
               Top             =   2160
               Width           =   1095
            End
            Begin VB.OptionButton optRobotSpeed 
               Caption         =   "3"
               Height          =   495
               Index           =   2
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   398
               Top             =   1560
               Width           =   1095
            End
            Begin VB.OptionButton optRobotSpeed 
               Caption         =   "2"
               Height          =   495
               Index           =   1
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   397
               Top             =   960
               Width           =   1095
            End
            Begin VB.OptionButton optRobotSpeed 
               Caption         =   "1"
               Height          =   495
               Index           =   0
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   396
               Top             =   360
               Width           =   1095
            End
         End
         Begin VB.TextBox txtCurrPos 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Height          =   390
            Index           =   3
            Left            =   -65640
            TabIndex        =   390
            Text            =   "0"
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtCurrPos 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Height          =   390
            Index           =   2
            Left            =   -66600
            TabIndex        =   389
            Text            =   "0"
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtCurrPos 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Height          =   390
            Index           =   1
            Left            =   -67560
            TabIndex        =   388
            Text            =   "0"
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtWriteTeach 
            Alignment       =   1  'Right Justify
            Height          =   390
            Index           =   3
            Left            =   -65640
            TabIndex        =   387
            Text            =   "0"
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox txtWriteTeach 
            Alignment       =   1  'Right Justify
            Height          =   390
            Index           =   2
            Left            =   -66600
            TabIndex        =   386
            Text            =   "0"
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox txtWriteTeach 
            Alignment       =   1  'Right Justify
            Height          =   390
            Index           =   1
            Left            =   -67560
            TabIndex        =   385
            Text            =   "0"
            Top             =   1320
            Width           =   855
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Read"
            Height          =   495
            Left            =   -64680
            TabIndex        =   383
            Top             =   720
            Width           =   1095
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Write"
            Height          =   495
            Left            =   -63480
            TabIndex        =   382
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtCurrPos 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Height          =   390
            Index           =   0
            Left            =   -68520
            TabIndex        =   381
            Text            =   "0"
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox Text8 
            Alignment       =   1  'Right Justify
            Height          =   390
            Left            =   -69480
            TabIndex        =   380
            Text            =   "0"
            Top             =   840
            Width           =   855
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Read"
            Height          =   495
            Left            =   -64680
            TabIndex        =   379
            Top             =   1320
            Width           =   1095
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Write"
            Height          =   495
            Left            =   -63480
            TabIndex        =   378
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txtWriteTeach 
            Alignment       =   1  'Right Justify
            Height          =   390
            Index           =   0
            Left            =   -68520
            TabIndex        =   377
            Text            =   "0"
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox txtTeachIndex 
            Alignment       =   1  'Right Justify
            Height          =   390
            Left            =   -69480
            TabIndex        =   376
            Text            =   "0"
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox Text2 
            Height          =   2175
            Left            =   -74760
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   374
            Text            =   "frmConfiguration.frx":0224
            Top             =   5760
            Width           =   5175
         End
         Begin VB.TextBox Text1 
            Height          =   495
            Left            =   -74760
            TabIndex        =   373
            Text            =   "Text1"
            Top             =   5160
            Width           =   5175
         End
         Begin VB.Frame Frame19 
            Caption         =   "Jog"
            Height          =   2775
            Left            =   -73320
            TabIndex        =   367
            Top             =   2280
            Width           =   1575
            Begin VB.CommandButton Command1 
               Caption         =   "A4"
               Height          =   495
               Left            =   240
               TabIndex        =   371
               Top             =   2160
               Width           =   1095
            End
            Begin VB.CommandButton Command2 
               Caption         =   "A3"
               Height          =   495
               Left            =   240
               TabIndex        =   370
               Top             =   1560
               Width           =   1095
            End
            Begin VB.CommandButton Command3 
               Caption         =   "A2"
               Height          =   495
               Left            =   240
               TabIndex        =   369
               Top             =   960
               Width           =   1095
            End
            Begin VB.CommandButton Command4 
               Caption         =   "A1"
               Height          =   495
               Left            =   240
               TabIndex        =   368
               Top             =   360
               Width           =   1095
            End
         End
         Begin VB.Frame Frame18 
            Caption         =   "Move"
            Height          =   2775
            Left            =   -71760
            TabIndex        =   362
            Top             =   2280
            Width           =   2295
            Begin VB.TextBox Text7 
               Alignment       =   1  'Right Justify
               Height          =   390
               Left            =   1320
               TabIndex        =   405
               Text            =   "0"
               Top             =   2160
               Width           =   855
            End
            Begin VB.TextBox Text6 
               Alignment       =   1  'Right Justify
               Height          =   390
               Left            =   1320
               TabIndex        =   404
               Text            =   "0"
               Top             =   1560
               Width           =   855
            End
            Begin VB.TextBox Text5 
               Alignment       =   1  'Right Justify
               Height          =   390
               Left            =   1320
               TabIndex        =   403
               Text            =   "0"
               Top             =   960
               Width           =   855
            End
            Begin VB.TextBox Text4 
               Alignment       =   1  'Right Justify
               Height          =   390
               Left            =   1320
               TabIndex        =   402
               Text            =   "0"
               Top             =   360
               Width           =   855
            End
            Begin VB.CommandButton Command6 
               Caption         =   "A1"
               Height          =   495
               Left            =   120
               TabIndex        =   366
               Top             =   360
               Width           =   1095
            End
            Begin VB.CommandButton Command7 
               Caption         =   "A2"
               Height          =   495
               Left            =   120
               TabIndex        =   365
               Top             =   960
               Width           =   1095
            End
            Begin VB.CommandButton Command8 
               Caption         =   "A3"
               Height          =   495
               Left            =   120
               TabIndex        =   364
               Top             =   1560
               Width           =   1095
            End
            Begin VB.CommandButton Command9 
               Caption         =   "A4"
               Height          =   495
               Left            =   120
               TabIndex        =   363
               Top             =   2160
               Width           =   1095
            End
         End
         Begin VB.TextBox txtRobotPort 
            Height          =   390
            Left            =   -73320
            TabIndex        =   360
            Text            =   "1"
            Top             =   840
            Width           =   855
         End
         Begin VB.Frame Frame1 
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
            Height          =   7695
            Left            =   -74880
            TabIndex        =   302
            Top             =   480
            Width           =   12735
            Begin VB.CommandButton Cmd_TcOffset 
               Caption         =   "溫度偏置"
               Height          =   495
               Left            =   11160
               TabIndex        =   615
               Top             =   6600
               Width           =   1455
            End
            Begin VB.Frame Frame17 
               Caption         =   "降溫參數"
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
               Height          =   1335
               Index           =   3
               Left            =   5760
               TabIndex        =   567
               Top             =   1200
               Width           =   2775
               Begin VB.TextBox txtTempDownTimeout 
                  Height          =   390
                  Left            =   1140
                  TabIndex        =   571
                  Text            =   "0"
                  Top             =   840
                  Width           =   975
               End
               Begin VB.TextBox txtParaNormal 
                  Height          =   390
                  Index           =   28
                  Left            =   1155
                  TabIndex        =   568
                  Text            =   "0"
                  Top             =   360
                  Width           =   975
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "Sec"
                  Height          =   270
                  Index           =   209
                  Left            =   2280
                  TabIndex        =   573
                  Top             =   960
                  Width           =   420
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
                  Left            =   240
                  TabIndex        =   572
                  Top             =   840
                  Width           =   750
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "℃"
                  Height          =   270
                  Index           =   10
                  Left            =   2235
                  TabIndex        =   570
                  Top             =   480
                  Width           =   240
               End
               Begin VB.Label lbName 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Open <"
                  Height          =   270
                  Index           =   206
                  Left            =   240
                  TabIndex        =   569
                  Top             =   360
                  Width           =   765
               End
            End
            Begin VB.TextBox txtGatePS1 
               Height          =   390
               Left            =   2640
               TabIndex        =   564
               Text            =   "0"
               Top             =   1800
               Width           =   1335
            End
            Begin VB.TextBox txtGatePS2 
               Height          =   390
               Left            =   4320
               TabIndex        =   563
               Text            =   "0"
               Top             =   1800
               Width           =   975
            End
            Begin VB.TextBox txtParaNormal 
               Height          =   390
               Index           =   27
               Left            =   7560
               TabIndex        =   533
               Text            =   "0"
               Top             =   3360
               Width           =   975
            End
            Begin VB.CheckBox chkCalibration 
               Caption         =   "Calibration"
               Height          =   270
               Left            =   7440
               TabIndex        =   523
               Top             =   7200
               Width           =   2175
            End
            Begin VB.CheckBox chkUseTempMeter 
               Caption         =   "溫度表連線"
               Height          =   495
               Left            =   11160
               Style           =   1  'Graphical
               TabIndex        =   522
               Top             =   6000
               Width           =   1455
            End
            Begin VB.Frame fraDatabase 
               Caption         =   "連線模式"
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
               Height          =   975
               Index           =   1
               Left            =   240
               TabIndex        =   506
               Top             =   240
               Width           =   3255
               Begin VB.ComboBox cmbCIMPort 
                  Height          =   390
                  ItemData        =   "frmConfiguration.frx":022A
                  Left            =   1560
                  List            =   "frmConfiguration.frx":0237
                  TabIndex        =   509
                  Text            =   "0=不連線"
                  Top             =   360
                  Width           =   1695
               End
               Begin VB.CheckBox chkCIM 
                  Caption         =   "CIM連線"
                  Height          =   495
                  Left            =   120
                  Style           =   1  'Graphical
                  TabIndex        =   507
                  Top             =   360
                  Width           =   1335
               End
            End
            Begin VB.Frame fraAutoMode 
               Caption         =   "自動機模式"
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
               Height          =   975
               Index           =   4
               Left            =   3600
               TabIndex        =   498
               Top             =   240
               Width           =   4935
               Begin VB.CheckBox chkAutoDoor 
                  Caption         =   "自動門開啟"
                  Height          =   495
                  Left            =   3240
                  Style           =   1  'Graphical
                  TabIndex        =   501
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.TextBox txtParaNormal 
                  Height          =   390
                  Index           =   26
                  Left            =   1800
                  TabIndex        =   500
                  Text            =   "0"
                  Top             =   480
                  Width           =   1095
               End
               Begin VB.CheckBox chkAutoMode 
                  Caption         =   "自動機連線"
                  Height          =   495
                  Left            =   240
                  Style           =   1  'Graphical
                  TabIndex        =   499
                  Top             =   360
                  Width           =   1455
               End
            End
            Begin VB.Frame fraDatabase 
               Caption         =   "連線模式"
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
               Height          =   975
               Index           =   0
               Left            =   240
               TabIndex        =   495
               Top             =   240
               Width           =   3255
               Begin VB.CheckBox chkBarcodeServer 
                  Caption         =   "資料庫連線"
                  Height          =   495
                  Left            =   120
                  Style           =   1  'Graphical
                  TabIndex        =   497
                  Top             =   360
                  Width           =   1335
               End
               Begin VB.TextBox txtServerPath 
                  Height          =   390
                  Left            =   1560
                  TabIndex        =   496
                  Text            =   "Y:"
                  Top             =   480
                  Width           =   1335
               End
            End
            Begin VB.TextBox txtParaNormal 
               Height          =   390
               Index           =   24
               Left            =   6600
               TabIndex        =   482
               Text            =   "0"
               Top             =   4800
               Width           =   615
            End
            Begin VB.TextBox txtParaNormal 
               Height          =   390
               Index           =   23
               Left            =   7560
               TabIndex        =   480
               Text            =   "500"
               Top             =   2880
               Width           =   975
            End
            Begin VB.TextBox txtParaNormal 
               Height          =   390
               Index           =   22
               Left            =   7320
               TabIndex        =   478
               Text            =   "0"
               Top             =   5640
               Width           =   975
            End
            Begin VB.TextBox txtParaNormal 
               Height          =   390
               Index           =   21
               Left            =   3960
               TabIndex        =   469
               Text            =   "0"
               Top             =   3240
               Width           =   975
            End
            Begin VB.CheckBox chkHoldSafety 
               Caption         =   "Hold Safety"
               Height          =   270
               Left            =   5400
               TabIndex        =   445
               Top             =   7200
               Width           =   2175
            End
            Begin VB.Frame Frame17 
               Caption         =   "監控提示設定"
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
               Height          =   1575
               Index           =   2
               Left            =   10200
               TabIndex        =   440
               Top             =   7680
               Width           =   2415
               Begin VB.TextBox txtParaNormal 
                  Height          =   390
                  Index           =   25
                  Left            =   1320
                  TabIndex        =   494
                  Text            =   "0"
                  Top             =   1080
                  Width           =   975
               End
               Begin VB.TextBox txtMonitorRuns 
                  Height          =   390
                  Left            =   1320
                  TabIndex        =   442
                  Text            =   "10"
                  Top             =   360
                  Width           =   975
               End
               Begin VB.TextBox txtTestRunKey 
                  Height          =   390
                  Left            =   1320
                  TabIndex        =   441
                  Text            =   "PR"
                  Top             =   720
                  Width           =   975
               End
               Begin VB.Label lbName 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "開門超過(s)"
                  Height          =   270
                  Index           =   187
                  Left            =   15
                  TabIndex        =   493
                  Top             =   1080
                  Width           =   1230
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "空跑關鍵字"
                  Height          =   270
                  Index           =   21
                  Left            =   120
                  TabIndex        =   444
                  Top             =   720
                  Width           =   1200
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "每跑幾次:"
                  Height          =   270
                  Index           =   20
                  Left            =   240
                  TabIndex        =   443
                  Top             =   360
                  Width           =   1020
               End
            End
            Begin VB.TextBox txtParaNormal 
               Height          =   390
               Index           =   2
               Left            =   2640
               TabIndex        =   435
               Text            =   "0"
               Top             =   2280
               Width           =   1335
            End
            Begin VB.TextBox txtParaNormal 
               Height          =   390
               Index           =   3
               Left            =   2640
               TabIndex        =   434
               Text            =   "500"
               Top             =   2760
               Width           =   1335
            End
            Begin VB.CheckBox chkAutoDeleteRecipe 
               Caption         =   "Auto Delete"
               Height          =   270
               Left            =   8880
               TabIndex        =   430
               Top             =   6720
               Width           =   2175
            End
            Begin VB.TextBox txtParaNormal 
               Height          =   390
               Index           =   20
               Left            =   2640
               TabIndex        =   428
               Text            =   "C:\Program Files\eRTP-100\Recipe\op"
               Top             =   6600
               Width           =   6135
            End
            Begin VB.Frame Frame17 
               Caption         =   "使用者密碼"
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
               Height          =   2295
               Index           =   0
               Left            =   8640
               TabIndex        =   416
               Top             =   240
               Width           =   3975
               Begin VB.TextBox txtActivePage 
                  Height          =   390
                  Index           =   2
                  Left            =   3240
                  TabIndex        =   422
                  Text            =   "0"
                  Top             =   1800
                  Width           =   495
               End
               Begin VB.TextBox txtActivePage 
                  Height          =   390
                  Index           =   1
                  Left            =   3240
                  TabIndex        =   421
                  Text            =   "0"
                  Top             =   1320
                  Width           =   495
               End
               Begin VB.TextBox txtActivePage 
                  Height          =   390
                  Index           =   0
                  Left            =   3240
                  TabIndex        =   420
                  Text            =   "0"
                  Top             =   840
                  Width           =   495
               End
               Begin VB.TextBox txtAdmin 
                  Height          =   390
                  IMEMode         =   3  'DISABLE
                  Left            =   1920
                  PasswordChar    =   "*"
                  TabIndex        =   419
                  Text            =   "0"
                  Top             =   840
                  Width           =   1215
               End
               Begin VB.TextBox txtOperator 
                  Height          =   390
                  IMEMode         =   3  'DISABLE
                  Left            =   1920
                  PasswordChar    =   "*"
                  TabIndex        =   418
                  Text            =   "0"
                  Top             =   1800
                  Width           =   1215
               End
               Begin VB.TextBox txtEngineer 
                  Height          =   390
                  IMEMode         =   3  'DISABLE
                  Left            =   1920
                  PasswordChar    =   "*"
                  TabIndex        =   417
                  Text            =   "0"
                  Top             =   1320
                  Width           =   1215
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "Page"
                  Height          =   270
                  Index           =   109
                  Left            =   3240
                  TabIndex        =   427
                  Top             =   360
                  Width           =   570
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "Admin(ad)"
                  Height          =   270
                  Index           =   145
                  Left            =   240
                  TabIndex        =   426
                  Top             =   840
                  Width           =   1095
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "Operator(op)"
                  Height          =   270
                  Index           =   144
                  Left            =   240
                  TabIndex        =   425
                  Top             =   1800
                  Width           =   1350
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "Engineer(eg)"
                  Height          =   270
                  Index           =   143
                  Left            =   240
                  TabIndex        =   424
                  Top             =   1320
                  Width           =   1365
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "Password"
                  Height          =   270
                  Index           =   142
                  Left            =   2160
                  TabIndex        =   423
                  Top             =   360
                  Width           =   1050
               End
            End
            Begin VB.Frame Frame17 
               Caption         =   "Offset"
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
               Height          =   2895
               Index           =   1
               Left            =   8640
               TabIndex        =   411
               Top             =   2640
               Width           =   2415
               Begin VB.CommandButton cmdCustom 
                  Caption         =   "Custom"
                  Height          =   375
                  Left            =   240
                  TabIndex        =   521
                  Top             =   2520
                  Width           =   975
               End
               Begin VB.TextBox txtRatioEX 
                  Height          =   390
                  Index           =   5
                  Left            =   960
                  TabIndex        =   516
                  Text            =   "1"
                  Top             =   1800
                  Width           =   1335
               End
               Begin VB.TextBox txtRatioEX 
                  Height          =   390
                  Index           =   4
                  Left            =   960
                  TabIndex        =   514
                  Text            =   "1"
                  Top             =   1440
                  Width           =   1335
               End
               Begin VB.TextBox txtRatioEX 
                  Height          =   390
                  Index           =   3
                  Left            =   960
                  TabIndex        =   512
                  Text            =   "1"
                  Top             =   1080
                  Width           =   1335
               End
               Begin VB.TextBox txtRatioEX 
                  Height          =   390
                  Index           =   2
                  Left            =   960
                  TabIndex        =   510
                  Text            =   "1"
                  Top             =   720
                  Width           =   1335
               End
               Begin VB.TextBox txtRatioEX 
                  Height          =   390
                  Index           =   1
                  Left            =   960
                  TabIndex        =   415
                  Text            =   "1"
                  Top             =   360
                  Width           =   1335
               End
               Begin VB.TextBox txtRatioEX 
                  Height          =   390
                  Index           =   0
                  Left            =   960
                  TabIndex        =   414
                  Text            =   "1"
                  Top             =   2160
                  Visible         =   0   'False
                  Width           =   1335
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "TC5"
                  Height          =   270
                  Index           =   191
                  Left            =   360
                  TabIndex        =   517
                  Top             =   1800
                  Width           =   450
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "TC4"
                  Height          =   270
                  Index           =   190
                  Left            =   360
                  TabIndex        =   515
                  Top             =   1440
                  Width           =   450
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "TC3"
                  Height          =   270
                  Index           =   189
                  Left            =   360
                  TabIndex        =   513
                  Top             =   1080
                  Width           =   450
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "TC2"
                  Height          =   270
                  Index           =   188
                  Left            =   360
                  TabIndex        =   511
                  Top             =   720
                  Width           =   450
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "TC"
                  Height          =   270
                  Index           =   172
                  Left            =   480
                  TabIndex        =   413
                  Top             =   2160
                  Visible         =   0   'False
                  Width           =   315
               End
               Begin VB.Label lbName 
                  AutoSize        =   -1  'True
                  Caption         =   "TC1"
                  Height          =   270
                  Index           =   170
                  Left            =   360
                  TabIndex        =   412
                  Top             =   360
                  Width           =   450
               End
            End
            Begin VB.CommandButton cmdLogPath 
               Caption         =   "Open"
               Height          =   390
               Left            =   8880
               TabIndex        =   410
               Top             =   6120
               Width           =   735
            End
            Begin VB.TextBox txtParaNormal 
               Height          =   390
               Index           =   19
               Left            =   10650
               TabIndex        =   407
               Text            =   "0"
               Top             =   7200
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.TextBox txtParaNormal 
               Height          =   390
               Index           =   0
               Left            =   4320
               TabIndex        =   330
               Text            =   "1200"
               Top             =   1320
               Width           =   975
            End
            Begin VB.TextBox txtParaNormal 
               Height          =   390
               Index           =   1
               Left            =   480
               TabIndex        =   329
               Text            =   "30"
               Top             =   7320
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.TextBox txtParaNormal 
               Height          =   390
               Index           =   5
               Left            =   120
               TabIndex        =   328
               Text            =   "5"
               Top             =   7320
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.TextBox txtParaNormal 
               Height          =   390
               Index           =   6
               Left            =   120
               TabIndex        =   327
               Text            =   "10"
               Top             =   7320
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.TextBox txtParaNormal 
               Height          =   390
               Index           =   4
               Left            =   480
               TabIndex        =   326
               Text            =   "1"
               Top             =   7320
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.CheckBox chkUniformity 
               Caption         =   "Disable"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   325
               Top             =   7320
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.TextBox txtParaNormal 
               Height          =   390
               Index           =   7
               Left            =   240
               TabIndex        =   324
               Text            =   "500"
               Top             =   7320
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Frame Frame15 
               Caption         =   "Mode"
               Height          =   1695
               Left            =   0
               TabIndex        =   321
               Top             =   7440
               Visible         =   0   'False
               Width           =   3855
               Begin VB.OptionButton optManual 
                  Caption         =   "Manual"
                  Height          =   375
                  Left            =   480
                  TabIndex        =   323
                  Top             =   480
                  Value           =   -1  'True
                  Width           =   2415
               End
               Begin VB.OptionButton optSemiAuto 
                  Caption         =   "Semi-Auto"
                  Height          =   375
                  Left            =   480
                  TabIndex        =   322
                  Top             =   1080
                  Width           =   2415
               End
            End
            Begin VB.TextBox txtParaNormal 
               Height          =   390
               Index           =   9
               Left            =   2640
               TabIndex        =   320
               Text            =   "0"
               Top             =   3720
               Width           =   1335
            End
            Begin VB.TextBox txtParaNormal 
               Height          =   390
               Index           =   10
               Left            =   2640
               TabIndex        =   319
               Text            =   "0"
               Top             =   1320
               Width           =   1335
            End
            Begin VB.TextBox txtParaNormal 
               Height          =   390
               Index           =   11
               Left            =   2640
               TabIndex        =   318
               Text            =   "0"
               Top             =   4680
               Width           =   1335
            End
            Begin VB.TextBox txtParaNormal 
               Height          =   390
               Index           =   12
               Left            =   2655
               TabIndex        =   317
               Text            =   "0"
               Top             =   5160
               Width           =   1335
            End
            Begin VB.TextBox txtParaNormal 
               Height          =   390
               Index           =   13
               Left            =   2640
               TabIndex        =   316
               Text            =   "C:\Program Files\eRTP-100\Log"
               Top             =   6120
               Width           =   6135
            End
            Begin VB.TextBox txtParaNormal 
               Height          =   390
               Index           =   14
               Left            =   2640
               TabIndex        =   315
               Text            =   "0"
               Top             =   5640
               Width           =   1335
            End
            Begin VB.TextBox txtParaNormal 
               Height          =   390
               Index           =   15
               Left            =   5640
               TabIndex        =   314
               Text            =   "0"
               Top             =   5640
               Width           =   975
            End
            Begin VB.TextBox txtParaNormal 
               Height          =   390
               Index           =   16
               Left            =   2640
               TabIndex        =   313
               Text            =   "0"
               Top             =   4200
               Width           =   1335
            End
            Begin VB.Frame Frame12 
               Caption         =   "Monitor TC"
               Height          =   3375
               Left            =   11160
               TabIndex        =   305
               Top             =   2640
               Width           =   1455
               Begin VB.CheckBox chkMonitorTCActive 
                  Caption         =   "MTC8"
                  Height          =   270
                  Index           =   7
                  Left            =   240
                  TabIndex        =   574
                  Top             =   2880
                  Width           =   1095
               End
               Begin VB.CheckBox chkMonitorTCActive 
                  Caption         =   "MTC3"
                  Height          =   270
                  Index           =   2
                  Left            =   240
                  TabIndex        =   312
                  Top             =   1080
                  Width           =   1095
               End
               Begin VB.CheckBox chkMonitorTCActive 
                  Caption         =   "MTC2"
                  Height          =   270
                  Index           =   1
                  Left            =   240
                  TabIndex        =   311
                  Top             =   720
                  Width           =   1095
               End
               Begin VB.CheckBox chkMonitorTCActive 
                  Caption         =   "MTC1"
                  Height          =   270
                  Index           =   0
                  Left            =   240
                  TabIndex        =   310
                  Top             =   360
                  Width           =   1095
               End
               Begin VB.CheckBox chkMonitorTCActive 
                  Caption         =   "MTC4"
                  Height          =   270
                  Index           =   3
                  Left            =   240
                  TabIndex        =   309
                  Top             =   1440
                  Width           =   1095
               End
               Begin VB.CheckBox chkMonitorTCActive 
                  Caption         =   "MTC5"
                  Height          =   270
                  Index           =   4
                  Left            =   240
                  TabIndex        =   308
                  Top             =   1800
                  Width           =   1095
               End
               Begin VB.CheckBox chkMonitorTCActive 
                  Caption         =   "MTC6"
                  Height          =   270
                  Index           =   5
                  Left            =   240
                  TabIndex        =   307
                  Top             =   2160
                  Width           =   1095
               End
               Begin VB.CheckBox chkMonitorTCActive 
                  Caption         =   "MTC7"
                  Height          =   270
                  Index           =   6
                  Left            =   240
                  TabIndex        =   306
                  Top             =   2520
                  Width           =   1095
               End
            End
            Begin VB.TextBox txtParaNormal 
               Height          =   390
               Index           =   17
               Left            =   2640
               TabIndex        =   304
               Text            =   "0"
               Top             =   7080
               Width           =   1335
            End
            Begin VB.TextBox txtParaNormal 
               Height          =   390
               Index           =   18
               Left            =   2640
               TabIndex        =   303
               Text            =   "0"
               Top             =   3240
               Width           =   1335
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "~"
               Height          =   270
               Index           =   208
               Left            =   4080
               TabIndex        =   566
               Top             =   1920
               Width           =   135
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Pressure Limit"
               Height          =   270
               Index           =   257
               Left            =   855
               TabIndex        =   565
               Top             =   1800
               Width           =   1515
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Pump Delay"
               Height          =   270
               Index           =   199
               Left            =   6225
               TabIndex        =   534
               Top             =   3360
               Width           =   1275
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "sec"
               Height          =   270
               Index           =   179
               Left            =   7320
               TabIndex        =   484
               Top             =   4920
               Width           =   375
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Alarm of lamp"
               Height          =   270
               Index           =   178
               Left            =   5130
               TabIndex        =   483
               Top             =   4800
               Width           =   1440
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "O2 Down"
               Height          =   270
               Index           =   177
               Left            =   6480
               TabIndex        =   481
               Top             =   2880
               Width           =   975
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Step"
               Height          =   270
               Index           =   176
               Left            =   8400
               TabIndex        =   479
               Top             =   5760
               Width           =   495
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "mTorr (0=off)"
               Height          =   270
               Index           =   15
               Left            =   4080
               TabIndex        =   439
               Top             =   2880
               Width           =   1350
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Vacuum Timout"
               Height          =   270
               Index           =   3
               Left            =   480
               TabIndex        =   438
               Top             =   2280
               Width           =   1860
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Vacuum Down"
               Height          =   270
               Index           =   7
               Left            =   360
               TabIndex        =   437
               Top             =   2760
               Width           =   1995
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Sec"
               Height          =   270
               Index           =   14
               Left            =   4200
               TabIndex        =   436
               Top             =   2400
               Width           =   420
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Recipe File Path"
               Height          =   270
               Index           =   136
               Left            =   690
               TabIndex        =   429
               Top             =   6600
               Width           =   1740
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Idle Remind"
               Height          =   270
               Index           =   169
               Left            =   9240
               TabIndex        =   409
               Top             =   7200
               Visible         =   0   'False
               Width           =   1230
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "min(0=off)"
               Height          =   270
               Index           =   156
               Left            =   12120
               TabIndex        =   408
               Top             =   7200
               Visible         =   0   'False
               Width           =   1050
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Tempurature Limit"
               Height          =   270
               Index           =   0
               Left            =   465
               TabIndex        =   359
               Top             =   1320
               Width           =   1890
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Preheat Intensity"
               Height          =   270
               Index           =   2
               Left            =   240
               TabIndex        =   358
               Top             =   7440
               Visible         =   0   'False
               Width           =   2535
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "%"
               Height          =   270
               Index           =   13
               Left            =   600
               TabIndex        =   357
               Top             =   7440
               Visible         =   0   'False
               Width           =   210
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Sec"
               Height          =   270
               Index           =   24
               Left            =   960
               TabIndex        =   356
               Top             =   7440
               Visible         =   0   'False
               Width           =   420
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Pumping Flow Change"
               Height          =   270
               Index           =   25
               Left            =   120
               TabIndex        =   355
               Top             =   7440
               Visible         =   0   'False
               Width           =   2370
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "TC Differential Range"
               Height          =   270
               Index           =   26
               Left            =   120
               TabIndex        =   354
               Top             =   7440
               Visible         =   0   'False
               Width           =   2250
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "℃"
               Height          =   270
               Index           =   27
               Left            =   1080
               TabIndex        =   353
               Top             =   7320
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Pressure Control"
               Height          =   270
               Index           =   80
               Left            =   360
               TabIndex        =   352
               Top             =   7440
               Visible         =   0   'False
               Width           =   1755
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Torr (Max=2 / Min=0.15)"
               Height          =   270
               Index           =   81
               Left            =   0
               TabIndex        =   351
               Top             =   7320
               Visible         =   0   'False
               Width           =   2490
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Uniformity Test"
               Height          =   270
               Index           =   127
               Left            =   240
               TabIndex        =   350
               Top             =   7320
               Visible         =   0   'False
               Width           =   1545
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Vent Gate"
               Height          =   270
               Index           =   134
               Left            =   240
               TabIndex        =   349
               Top             =   7440
               Visible         =   0   'False
               Width           =   2610
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Torr (100 ~ 760)"
               Height          =   270
               Index           =   135
               Left            =   480
               TabIndex        =   348
               Top             =   7440
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Finished Beep"
               Height          =   270
               Index           =   84
               Left            =   480
               TabIndex        =   347
               Top             =   3720
               Width           =   1890
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "sec (0=off)"
               Height          =   270
               Index           =   86
               Left            =   4200
               TabIndex        =   346
               Top             =   3840
               Width           =   1110
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "~"
               Height          =   270
               Index           =   103
               Left            =   4080
               TabIndex        =   345
               Top             =   1440
               Width           =   135
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "hrs"
               Height          =   270
               Index           =   105
               Left            =   4200
               TabIndex        =   344
               Top             =   4800
               Width           =   315
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Lifetime of Lamp"
               Height          =   270
               Index           =   117
               Left            =   615
               TabIndex        =   343
               Top             =   4680
               Width           =   1755
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "hrs"
               Height          =   270
               Index           =   121
               Left            =   4215
               TabIndex        =   342
               Top             =   5280
               Width           =   315
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Used of Lamp"
               Height          =   270
               Index           =   122
               Left            =   915
               TabIndex        =   341
               Top             =   5160
               Width           =   1470
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Log File Path"
               Height          =   270
               Index           =   141
               Left            =   1035
               TabIndex        =   340
               Top             =   6120
               Width           =   1395
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "℃ (0=off)"
               Height          =   270
               Index           =   146
               Left            =   4200
               TabIndex        =   339
               Top             =   5760
               Width           =   975
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Max Monitor error"
               Height          =   270
               Index           =   147
               Left            =   540
               TabIndex        =   338
               Top             =   5640
               Width           =   1830
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "Sec"
               Height          =   270
               Index           =   148
               Left            =   6720
               TabIndex        =   337
               Top             =   5760
               Width           =   420
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "℃(0=off)"
               Height          =   270
               Index           =   150
               Left            =   4200
               TabIndex        =   336
               Top             =   4320
               Width           =   915
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Finished Light"
               Height          =   270
               Index           =   151
               Left            =   900
               TabIndex        =   335
               Top             =   4200
               Width           =   1470
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Cycle Run"
               Height          =   270
               Index           =   159
               Left            =   1320
               TabIndex        =   334
               Top             =   7080
               Width           =   1050
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "(0=off)"
               Height          =   270
               Index           =   160
               Left            =   4080
               TabIndex        =   333
               Top             =   7200
               Width           =   675
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "mTorr"
               Height          =   270
               Index           =   162
               Left            =   5040
               TabIndex        =   332
               Top             =   3360
               Width           =   615
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Gauge Value <"
               Height          =   270
               Index           =   163
               Left            =   810
               TabIndex        =   331
               Top             =   3240
               Width           =   1560
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgTeach 
            Height          =   5535
            Left            =   -69480
            TabIndex        =   375
            Top             =   2400
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   9763
            _Version        =   393216
            Rows            =   17
            AllowBigSelection=   0   'False
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgAlarm 
            Height          =   6495
            Left            =   240
            TabIndex        =   431
            Top             =   720
            Width           =   12495
            _ExtentX        =   22040
            _ExtentY        =   11456
            _Version        =   393216
            Rows            =   17
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label2 
            Caption         =   "0=不顯示警報,1=顯示(不鳴),2=顯示(鳴),4=顯示(鳴,立即停止)"
            Height          =   375
            Left            =   5760
            TabIndex        =   433
            Top             =   7440
            Width           =   6975
         End
         Begin VB.Label lbName 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PlaceH:"
            Height          =   270
            Index           =   168
            Left            =   -74280
            TabIndex        =   394
            Top             =   1800
            Width           =   825
         End
         Begin VB.Label lbName 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "PickH:"
            Height          =   270
            Index           =   167
            Left            =   -74160
            TabIndex        =   392
            Top             =   1320
            Width           =   690
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Curr Pos:"
            Height          =   270
            Index           =   166
            Left            =   -70680
            TabIndex        =   384
            Top             =   840
            Width           =   990
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Teach Point:"
            Height          =   270
            Index           =   165
            Left            =   -70920
            TabIndex        =   372
            Top             =   1320
            Width           =   1305
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Robot Port:"
            Height          =   270
            Index           =   164
            Left            =   -74640
            TabIndex        =   361
            Top             =   840
            Width           =   1185
         End
      End
      Begin VB.Frame Frame13 
         Height          =   7335
         Left            =   -65040
         TabIndex        =   238
         Top             =   720
         Width           =   3015
         Begin VB.CheckBox chkOnlyRecipe 
            Caption         =   "Only used in Process"
            Height          =   375
            Left            =   240
            TabIndex        =   520
            Top             =   2400
            Width           =   2655
         End
         Begin VB.Frame Frame16 
            Caption         =   "Use Error Map"
            Height          =   3015
            Left            =   1920
            TabIndex        =   252
            Top             =   6600
            Visible         =   0   'False
            Width           =   2295
            Begin VB.CheckBox chkErrorTCActive 
               Caption         =   "TC7"
               Height          =   375
               Index           =   6
               Left            =   240
               TabIndex        =   259
               Top             =   2520
               Width           =   1695
            End
            Begin VB.CheckBox chkErrorTCActive 
               Caption         =   "TC6"
               Height          =   375
               Index           =   5
               Left            =   240
               TabIndex        =   258
               Top             =   2160
               Width           =   1695
            End
            Begin VB.CheckBox chkErrorTCActive 
               Caption         =   "TC5"
               Height          =   375
               Index           =   4
               Left            =   240
               TabIndex        =   257
               Top             =   1800
               Width           =   1695
            End
            Begin VB.CheckBox chkErrorTCActive 
               Caption         =   "TC4"
               Height          =   375
               Index           =   3
               Left            =   240
               TabIndex        =   256
               Top             =   1440
               Width           =   1695
            End
            Begin VB.CheckBox chkErrorTCActive 
               Caption         =   "TC3"
               Height          =   375
               Index           =   2
               Left            =   240
               TabIndex        =   255
               Top             =   1080
               Width           =   1695
            End
            Begin VB.CheckBox chkErrorTCActive 
               Caption         =   "TC2"
               Height          =   375
               Index           =   1
               Left            =   240
               TabIndex        =   254
               Top             =   720
               Width           =   1695
            End
            Begin VB.CheckBox chkErrorTCActive 
               Caption         =   "TC1"
               Height          =   375
               Index           =   0
               Left            =   240
               TabIndex        =   253
               Top             =   360
               Width           =   1695
            End
         End
         Begin VB.ComboBox cmbTCVoltageRange 
            Height          =   390
            Left            =   600
            TabIndex        =   241
            Text            =   "0"
            Top             =   1800
            Width           =   2175
         End
         Begin VB.ComboBox cmbTCType 
            Height          =   390
            Left            =   600
            TabIndex        =   240
            Text            =   "0"
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label1 
            Caption         =   "TC Voltage Range(重開):"
            Height          =   375
            Left            =   240
            TabIndex        =   242
            Top             =   1320
            Width           =   2655
         End
         Begin VB.Label lbTCType 
            Caption         =   "TC Type(設定完要重開):"
            Height          =   375
            Left            =   240
            TabIndex        =   239
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame fraAdvDaqAI 
         Appearance      =   0  'Flat
         Caption         =   "PCI-1710HGU : AI"
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
         Height          =   7335
         Left            =   -74880
         TabIndex        =   235
         Top             =   720
         Width           =   9615
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgAdvDaqAI 
            Height          =   6495
            Left            =   240
            TabIndex        =   236
            Top             =   720
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   11456
            _Version        =   393216
            Rows            =   17
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame fraUniformity 
         Caption         =   "Uniformity"
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
         Height          =   855
         Left            =   -63120
         TabIndex        =   221
         Top             =   7920
         Visible         =   0   'False
         Width           =   735
         Begin VB.CheckBox chkAlarmBuzzer 
            Caption         =   "Close"
            Enabled         =   0   'False
            Height          =   495
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   299
            Top             =   5880
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox chkSimulation 
            Caption         =   "Simulation"
            Height          =   375
            Left            =   240
            TabIndex        =   298
            Top             =   5400
            Width           =   3375
         End
         Begin VB.CheckBox chkCTCheck 
            Caption         =   "CT Check Status"
            Height          =   375
            Left            =   240
            TabIndex        =   297
            Top             =   4920
            Width           =   3375
         End
         Begin VB.CheckBox chkMonitorTC 
            Caption         =   "On"
            Height          =   495
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   295
            Top             =   4200
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtUniformityRampStartPoint 
            Height          =   390
            Left            =   2040
            TabIndex        =   234
            Text            =   "0"
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox txtSubWeightD2 
            Height          =   390
            Left            =   2040
            TabIndex        =   232
            Text            =   "0"
            Top             =   2160
            Width           =   1575
         End
         Begin VB.TextBox txtSubWeightD1 
            Height          =   390
            Left            =   2040
            TabIndex        =   231
            Text            =   "0"
            Top             =   1680
            Width           =   1575
         End
         Begin VB.TextBox txtUniformityHoldStartPoint 
            Height          =   390
            Left            =   2040
            TabIndex        =   228
            Text            =   "0"
            Top             =   2640
            Width           =   1575
         End
         Begin VB.CheckBox chkUniformityRampActive 
            Caption         =   "Ramp Active"
            Height          =   495
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   227
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txtSubWeight2 
            Height          =   390
            Left            =   2040
            TabIndex        =   225
            Text            =   "0"
            Top             =   3600
            Width           =   1575
         End
         Begin VB.TextBox txtSubWeight1 
            Height          =   390
            Left            =   2040
            TabIndex        =   224
            Text            =   "0"
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Alarm Buzzer"
            Height          =   270
            Index           =   6
            Left            =   240
            TabIndex        =   300
            Top             =   6000
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Monitor"
            Height          =   270
            Index           =   22
            Left            =   240
            TabIndex        =   296
            Top             =   4320
            Width           =   780
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Ramp Start Point"
            Height          =   270
            Index           =   133
            Left            =   240
            TabIndex        =   233
            Top             =   1200
            Width           =   1785
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "SubWeightD2"
            Height          =   270
            Index           =   132
            Left            =   360
            TabIndex        =   230
            Top             =   2160
            Width           =   1470
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "SubWeightD1"
            Height          =   270
            Index           =   131
            Left            =   360
            TabIndex        =   229
            Top             =   1680
            Width           =   1470
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Hold Start Point"
            Height          =   270
            Index           =   130
            Left            =   360
            TabIndex        =   226
            Top             =   2640
            Width           =   1635
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "SubWeight2"
            Height          =   270
            Index           =   129
            Left            =   360
            TabIndex        =   223
            Top             =   3600
            Width           =   1290
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "SubWeight1"
            Height          =   270
            Index           =   128
            Left            =   360
            TabIndex        =   222
            Top             =   3120
            Width           =   1290
         End
      End
      Begin VB.Frame fraSmoothCurve 
         Caption         =   "Ramp Smooth Curve"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1095
         Left            =   -66960
         TabIndex        =   192
         Top             =   7800
         Width           =   4935
         Begin VB.CheckBox chkSmoothDisplay 
            Caption         =   "Display"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   196
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtSmoothTime 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2040
            TabIndex        =   195
            Text            =   "5000"
            Top             =   600
            Width           =   1335
         End
         Begin VB.CheckBox chkSmoothRamp 
            Caption         =   "Ramp Smooth"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   193
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Delay Time(ms)"
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
            Left            =   2040
            TabIndex        =   194
            Top             =   240
            Width           =   1365
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Moudle Activity"
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
         Height          =   7095
         Left            =   -67680
         TabIndex        =   177
         Top             =   720
         Width           =   5655
         Begin VB.CheckBox ChkTestMode 
            Caption         =   "Enable"
            Height          =   270
            Left            =   1560
            TabIndex        =   625
            Top             =   6600
            Width           =   1095
         End
         Begin VB.CheckBox ChkShowChamberNo 
            Caption         =   "Enable"
            Height          =   270
            Left            =   4440
            TabIndex        =   619
            Top             =   5760
            Width           =   1095
         End
         Begin VB.CheckBox ChkTcOffset 
            Caption         =   "Enable"
            Height          =   270
            Left            =   2040
            TabIndex        =   617
            Top             =   5760
            Width           =   1095
         End
         Begin VB.TextBox txtCommSCR 
            Height          =   390
            Left            =   4800
            TabIndex        =   600
            Top             =   5280
            Width           =   375
         End
         Begin VB.CheckBox ckSCREable 
            Caption         =   "Enable"
            Height          =   270
            Left            =   3720
            TabIndex        =   599
            Top             =   5280
            Width           =   1095
         End
         Begin VB.CheckBox ckeStopTCM 
            Caption         =   "Enable"
            Height          =   270
            Left            =   1560
            TabIndex        =   597
            Top             =   5280
            Width           =   1095
         End
         Begin VB.CheckBox ChkWriteOffsetToTCM 
            Caption         =   "Offset寫入TCM"
            Height          =   270
            Left            =   2760
            TabIndex        =   581
            Top             =   6240
            Width           =   2295
         End
         Begin VB.CheckBox ChkDefineProcStep 
            Caption         =   "ProcessStep自定義"
            Height          =   375
            Left            =   240
            TabIndex        =   576
            Top             =   6120
            Width           =   2415
         End
         Begin VB.CheckBox chkModuleEnable 
            Caption         =   "Enable"
            Height          =   375
            Index           =   19
            Left            =   3720
            TabIndex        =   561
            Top             =   1920
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chkModuleEnable 
            Caption         =   "Enable"
            Height          =   375
            Index           =   18
            Left            =   3720
            TabIndex        =   546
            Top             =   4800
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chkModuleEnable 
            Caption         =   "Enable"
            Height          =   375
            Index           =   17
            Left            =   3720
            TabIndex        =   535
            Top             =   4320
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chkModuleEnable 
            Caption         =   "Enable"
            Height          =   375
            Index           =   16
            Left            =   3720
            TabIndex        =   527
            Top             =   3840
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chkModuleEnable 
            Caption         =   "Enable"
            Height          =   375
            Index           =   15
            Left            =   3720
            TabIndex        =   524
            Top             =   3360
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chkModuleEnable 
            Caption         =   "Enable"
            Height          =   375
            Index           =   14
            Left            =   3720
            TabIndex        =   518
            Top             =   2880
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chkModuleEnable 
            Caption         =   "Enable"
            Height          =   375
            Index           =   13
            Left            =   3720
            TabIndex        =   504
            Top             =   2400
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chkModuleEnable 
            Caption         =   "Enable"
            Height          =   375
            Index           =   12
            Left            =   3720
            TabIndex        =   467
            Top             =   1440
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.TextBox txtComCT 
            Height          =   390
            Left            =   4800
            TabIndex        =   466
            Text            =   "3"
            Top             =   960
            Width           =   375
         End
         Begin VB.CheckBox chkModuleEnable 
            Caption         =   "Enable"
            Height          =   375
            Index           =   11
            Left            =   3720
            TabIndex        =   464
            Top             =   960
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chkModuleEnable 
            Caption         =   "Enable"
            Height          =   375
            Index           =   10
            Left            =   3720
            TabIndex        =   446
            Top             =   480
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chkModuleEnable 
            Caption         =   "Enable"
            Height          =   375
            Index           =   9
            Left            =   1560
            TabIndex        =   279
            Top             =   4800
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chkModuleEnable 
            Caption         =   "Enable"
            Height          =   375
            Index           =   8
            Left            =   1560
            TabIndex        =   267
            Top             =   3840
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chkModuleEnable 
            Caption         =   "Enable"
            Height          =   375
            Index           =   7
            Left            =   1560
            TabIndex        =   265
            Top             =   4320
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkModuleEnable 
            Caption         =   "Enable"
            Height          =   375
            Index           =   6
            Left            =   1560
            TabIndex        =   263
            Top             =   3360
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkModuleEnable 
            Caption         =   "Enable"
            Height          =   375
            Index           =   5
            Left            =   1560
            TabIndex        =   248
            Top             =   2880
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkModuleEnable 
            Caption         =   "Enable"
            Height          =   375
            Index           =   4
            Left            =   1560
            TabIndex        =   187
            Top             =   2400
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chkModuleEnable 
            Caption         =   "Enable"
            Height          =   375
            Index           =   3
            Left            =   1560
            TabIndex        =   186
            Top             =   1920
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chkModuleEnable 
            Caption         =   "Enable"
            Height          =   375
            Index           =   2
            Left            =   1560
            TabIndex        =   185
            Top             =   1440
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chkModuleEnable 
            Caption         =   "Enable"
            Height          =   375
            Index           =   1
            Left            =   1560
            TabIndex        =   184
            Top             =   960
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chkModuleEnable 
            Caption         =   "Enable"
            Height          =   375
            Index           =   0
            Left            =   1560
            TabIndex        =   183
            Top             =   480
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.Label LbTestMode 
            Caption         =   "TestMode:"
            Height          =   255
            Left            =   240
            TabIndex        =   624
            Top             =   6600
            Width           =   1095
         End
         Begin VB.Label Lb_ShowChamberNo 
            Caption         =   "ChamberNo"
            Height          =   255
            Left            =   3120
            TabIndex        =   618
            Top             =   5760
            Width           =   1695
         End
         Begin VB.Label LbTcOffset 
            Caption         =   "HoldTempOffset"
            Height          =   255
            Left            =   240
            TabIndex        =   616
            Top             =   5760
            Width           =   1815
         End
         Begin VB.Label Label6 
            Caption         =   "SCR"
            Height          =   255
            Left            =   3000
            TabIndex        =   598
            Top             =   5280
            Width           =   495
         End
         Begin VB.Label lbStoptcm 
            Caption         =   "StopTCM"
            Height          =   255
            Left            =   240
            TabIndex        =   596
            Top             =   5280
            Width           =   1095
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "MTC2"
            Height          =   270
            Index           =   207
            Left            =   3000
            TabIndex        =   562
            Top             =   1920
            Width           =   645
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Cover"
            Height          =   270
            Index           =   204
            Left            =   2760
            TabIndex        =   547
            Top             =   4800
            Width           =   735
         End
         Begin VB.Label lbName 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "T-Pump"
            Height          =   270
            Index           =   200
            Left            =   2700
            TabIndex        =   536
            Top             =   4320
            Width           =   825
         End
         Begin VB.Label lbName 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Az2"
            Height          =   270
            Index           =   196
            Left            =   3120
            TabIndex        =   526
            Top             =   3840
            Width           =   405
         End
         Begin VB.Label lbName 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Az1"
            Height          =   270
            Index           =   193
            Left            =   3105
            TabIndex        =   525
            Top             =   3360
            Width           =   405
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "M-Loop"
            Height          =   270
            Index           =   192
            Left            =   2760
            TabIndex        =   519
            Top             =   2880
            Width           =   810
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "CIM"
            Height          =   270
            Index           =   171
            Left            =   3000
            TabIndex        =   505
            Top             =   2400
            Width           =   420
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "MTC1"
            Height          =   270
            Index           =   175
            Left            =   3000
            TabIndex        =   468
            Top             =   1440
            Width           =   645
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "CT"
            Height          =   270
            Index           =   174
            Left            =   3000
            TabIndex        =   465
            Top             =   960
            Width           =   315
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Auto"
            Height          =   270
            Index           =   89
            Left            =   3000
            TabIndex        =   447
            Top             =   480
            Width           =   480
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "APC"
            Height          =   270
            Index           =   158
            Left            =   240
            TabIndex        =   280
            Top             =   4800
            Width           =   510
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "PN-Recipe"
            Height          =   270
            Index           =   149
            Left            =   240
            TabIndex        =   268
            Top             =   3840
            Width           =   1155
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Auto Door"
            Height          =   270
            Index           =   140
            Left            =   240
            TabIndex        =   266
            Top             =   4320
            Width           =   1065
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Barcode"
            Height          =   270
            Index           =   139
            Left            =   240
            TabIndex        =   264
            Top             =   3360
            Width           =   900
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Oxygen"
            Height          =   270
            Index           =   16
            Left            =   240
            TabIndex        =   249
            Top             =   2880
            Width           =   780
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Database"
            Height          =   270
            Index           =   116
            Left            =   240
            TabIndex        =   182
            Top             =   2400
            Width           =   1035
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Vacuum"
            Height          =   270
            Index           =   115
            Left            =   240
            TabIndex        =   181
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Gas"
            Height          =   270
            Index           =   114
            Left            =   240
            TabIndex        =   180
            Top             =   1440
            Width           =   435
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Cooling"
            Height          =   270
            Index           =   113
            Left            =   240
            TabIndex        =   179
            Top             =   960
            Width           =   810
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Heating"
            Height          =   270
            Index           =   112
            Left            =   240
            TabIndex        =   178
            Top             =   480
            Width           =   810
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Alarm Activity"
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
         Height          =   9615
         Left            =   -74760
         TabIndex        =   133
         Top             =   720
         Width           =   6975
         Begin VB.CheckBox chkAlarmEnable 
            Caption         =   "Enable"
            Height          =   270
            Index           =   15
            Left            =   3360
            TabIndex        =   595
            Top             =   8160
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkAlarmEnable 
            Caption         =   "Enable"
            Height          =   270
            Index           =   14
            Left            =   3360
            TabIndex        =   591
            Top             =   7680
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CommandButton cmdModbusRTU 
            Caption         =   "RTU"
            Height          =   495
            Left            =   5280
            TabIndex        =   575
            Top             =   2640
            Width           =   1455
         End
         Begin VB.CommandButton cmdCIMUDP 
            Caption         =   "CIM UDP"
            Height          =   495
            Left            =   5280
            TabIndex        =   508
            Top             =   2040
            Width           =   1455
         End
         Begin VB.CommandButton cmdDCR 
            Caption         =   "DCR"
            Height          =   495
            Left            =   5280
            TabIndex        =   503
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CommandButton cmdAutoUDP 
            Caption         =   "Auto UDP"
            Height          =   495
            Left            =   5280
            TabIndex        =   502
            Top             =   840
            Width           =   1455
         End
         Begin VB.CheckBox chkAlarmEnable 
            Caption         =   "Enable"
            Height          =   270
            Index           =   13
            Left            =   3360
            TabIndex        =   250
            Top             =   7200
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkAlarmEnable 
            Caption         =   "Enable"
            Height          =   270
            Index           =   12
            Left            =   3360
            TabIndex        =   190
            Top             =   6720
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkAlarmEnable 
            Caption         =   "Enable"
            Height          =   270
            Index           =   11
            Left            =   3360
            TabIndex        =   188
            Top             =   6240
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkAlarmEnable 
            Caption         =   "Enable"
            Height          =   270
            Index           =   10
            Left            =   3360
            TabIndex        =   176
            Top             =   5760
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkAlarmEnable 
            Caption         =   "Enable"
            Height          =   270
            Index           =   9
            Left            =   3360
            TabIndex        =   154
            Top             =   5280
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkAlarmEnable 
            Caption         =   "Enable"
            Height          =   270
            Index           =   8
            Left            =   3360
            TabIndex        =   153
            Top             =   4800
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkAlarmEnable 
            Caption         =   "Enable"
            Height          =   270
            Index           =   7
            Left            =   3360
            TabIndex        =   152
            Top             =   4320
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkAlarmEnable 
            Caption         =   "Enable"
            Height          =   270
            Index           =   6
            Left            =   3360
            TabIndex        =   151
            Top             =   3840
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkAlarmEnable 
            Caption         =   "Enable"
            Height          =   270
            Index           =   5
            Left            =   3360
            TabIndex        =   150
            Top             =   3360
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkAlarmEnable 
            Caption         =   "Enable"
            Height          =   270
            Index           =   4
            Left            =   3360
            TabIndex        =   149
            Top             =   2880
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkAlarmEnable 
            Caption         =   "Enable"
            Height          =   270
            Index           =   3
            Left            =   3360
            TabIndex        =   148
            Top             =   2400
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkAlarmEnable 
            Caption         =   "Enable"
            Height          =   270
            Index           =   2
            Left            =   3360
            TabIndex        =   147
            Top             =   1920
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkAlarmEnable 
            Caption         =   "Enable"
            Height          =   270
            Index           =   1
            Left            =   3360
            TabIndex        =   146
            Top             =   1440
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox chkAlarmEnable 
            Caption         =   "Enable"
            Height          =   270
            Index           =   0
            Left            =   3360
            TabIndex        =   145
            Top             =   960
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "RValRange Alarm"
            Height          =   270
            Index           =   212
            Left            =   360
            TabIndex        =   594
            Top             =   8160
            Width           =   1875
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "TcWafer Alarm"
            Height          =   270
            Index           =   211
            Left            =   360
            TabIndex        =   590
            Top             =   7680
            Width           =   1560
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Chamber Pressure"
            Height          =   270
            Index           =   19
            Left            =   360
            TabIndex        =   251
            Top             =   7200
            Width           =   1980
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "APC Communication"
            Height          =   270
            Index           =   119
            Left            =   360
            TabIndex        =   191
            Top             =   6720
            Width           =   2205
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "TC Alarm"
            Height          =   270
            Index           =   118
            Left            =   360
            TabIndex        =   189
            Top             =   3840
            Width           =   990
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Alarm Buzzer"
            Height          =   270
            Index           =   111
            Left            =   360
            TabIndex        =   175
            Top             =   2400
            Width           =   1380
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "CT Alarm"
            Height          =   270
            Index           =   102
            Left            =   360
            TabIndex        =   144
            Top             =   3360
            Width           =   990
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Enable / Disable"
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
            Index           =   101
            Left            =   3360
            TabIndex        =   143
            Top             =   480
            Width           =   1830
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Vacuum Gauge"
            Height          =   270
            Index           =   100
            Left            =   360
            TabIndex        =   142
            Top             =   5280
            Width           =   1620
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Vacuum Switch"
            Height          =   270
            Index           =   99
            Left            =   360
            TabIndex        =   141
            Top             =   5760
            Width           =   1605
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Door Interlock"
            Height          =   270
            Index           =   98
            Left            =   360
            TabIndex        =   140
            Top             =   6240
            Width           =   1440
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Water Alarm"
            Height          =   270
            Index           =   97
            Left            =   360
            TabIndex        =   139
            Top             =   4320
            Width           =   1305
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Air Pressure"
            Height          =   270
            Index           =   96
            Left            =   360
            TabIndex        =   138
            Top             =   4800
            Width           =   1305
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "EMS Alarm"
            Height          =   270
            Index           =   95
            Left            =   360
            TabIndex        =   137
            Top             =   1920
            Width           =   1200
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Chamber Overheat"
            Height          =   270
            Index           =   94
            Left            =   360
            TabIndex        =   136
            Top             =   2880
            Width           =   1980
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "System Ready"
            Height          =   270
            Index           =   93
            Left            =   360
            TabIndex        =   135
            Top             =   1440
            Width           =   1515
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "System Alarm"
            Height          =   270
            Index           =   92
            Left            =   360
            TabIndex        =   134
            Top             =   960
            Width           =   1455
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Specical"
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
         Height          =   2415
         Left            =   -67680
         TabIndex        =   118
         Top             =   7920
         Width           =   5655
         Begin VB.TextBox TxthdTimes 
            Height          =   400
            Left            =   4200
            TabIndex        =   579
            Text            =   "0"
            Top             =   1920
            Width           =   855
         End
         Begin VB.TextBox TxthdOffset 
            Height          =   400
            Left            =   2640
            TabIndex        =   578
            Text            =   "0"
            Top             =   1920
            Width           =   1290
         End
         Begin VB.ComboBox cmbControlMode 
            Height          =   390
            ItemData        =   "frmConfiguration.frx":0259
            Left            =   1800
            List            =   "frmConfiguration.frx":025B
            TabIndex        =   219
            Text            =   "Combo1"
            Top             =   1440
            Width           =   1695
         End
         Begin VB.CheckBox chkMultiLoop 
            Caption         =   "Multi-Loop"
            Height          =   615
            Left            =   3480
            Style           =   1  'Graphical
            TabIndex        =   218
            Top             =   480
            Width           =   1695
         End
         Begin VB.CheckBox chkLampMonitor 
            Caption         =   "Lamp Monitor"
            Height          =   615
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   126
            Top             =   480
            Width           =   1695
         End
         Begin VB.CheckBox chkResetIntegral 
            Caption         =   "Integral Reset"
            Height          =   615
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   119
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "次"
            Height          =   375
            Left            =   5040
            TabIndex        =   580
            Top             =   1920
            Width           =   495
         End
         Begin VB.Line Line2 
            X1              =   3840
            X2              =   4200
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Label lbHdOffset 
            Caption         =   "HoldTime Offset(Sec):"
            Height          =   255
            Left            =   240
            TabIndex        =   577
            Top             =   2000
            Width           =   3255
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Control Loop"
            Height          =   270
            Index           =   126
            Left            =   240
            TabIndex        =   220
            Top             =   1440
            Width           =   1350
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Misc (設定完要重開):"
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
         Height          =   1215
         Left            =   -74760
         TabIndex        =   117
         Top             =   840
         Width           =   3135
         Begin VB.TextBox txtRtaType 
            Height          =   390
            Left            =   1560
            TabIndex        =   293
            Text            =   "0"
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "RTA Type:"
            Height          =   270
            Index           =   110
            Left            =   240
            TabIndex        =   294
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Vacuum"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   7095
         Left            =   -66960
         TabIndex        =   107
         Top             =   720
         Width           =   4935
         Begin VB.TextBox txtParaVacuum 
            Alignment       =   1  'Right Justify
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
            Left            =   3480
            TabIndex        =   542
            Text            =   "0"
            Top             =   7560
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtParaVacuum 
            Alignment       =   1  'Right Justify
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
            Left            =   3480
            TabIndex        =   541
            Text            =   "0"
            Top             =   6960
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Frame Frame22 
            Caption         =   "Gauge Para"
            Height          =   1455
            Left            =   120
            TabIndex        =   537
            Top             =   5520
            Width           =   4095
            Begin VB.TextBox txtGaugeVN 
               Height          =   390
               Left            =   2640
               TabIndex        =   545
               Text            =   "0"
               Top             =   960
               Width           =   1335
            End
            Begin VB.TextBox txtGaugeVP 
               Height          =   390
               Left            =   2640
               TabIndex        =   544
               Text            =   "0"
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox txtGaugeD 
               Height          =   390
               Left            =   600
               TabIndex        =   543
               Text            =   "0"
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "D"
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
               Index           =   203
               Left            =   360
               TabIndex        =   540
               Top             =   360
               Width           =   135
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "V-"
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
               Index           =   202
               Left            =   2220
               TabIndex        =   539
               Top             =   960
               Width           =   195
            End
            Begin VB.Label lbName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "V+"
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
               Index           =   201
               Left            =   2160
               TabIndex        =   538
               Top             =   360
               Width           =   255
            End
         End
         Begin VB.TextBox txtParaVacuum 
            Alignment       =   1  'Right Justify
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
            Left            =   2640
            TabIndex        =   476
            Text            =   "1"
            Top             =   5040
            Width           =   1335
         End
         Begin VB.TextBox txtParaVacuum 
            Alignment       =   1  'Right Justify
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
            Left            =   2640
            TabIndex        =   277
            Text            =   "1"
            Top             =   3600
            Width           =   1335
         End
         Begin VB.TextBox txtParaVacuum 
            Alignment       =   1  'Right Justify
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
            Left            =   2640
            TabIndex        =   276
            Text            =   "0"
            Top             =   4560
            Width           =   1335
         End
         Begin VB.TextBox txtParaVacuum 
            Alignment       =   1  'Right Justify
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
            Left            =   2640
            TabIndex        =   272
            Text            =   "1"
            Top             =   4080
            Width           =   1335
         End
         Begin VB.TextBox txtParaVacuum 
            Alignment       =   1  'Right Justify
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
            Left            =   2640
            TabIndex        =   269
            Text            =   "0.2"
            Top             =   3120
            Width           =   1335
         End
         Begin VB.TextBox txtParaVacuum 
            Alignment       =   1  'Right Justify
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
            Left            =   2640
            TabIndex        =   159
            Text            =   "0"
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtParaVacuum 
            Alignment       =   1  'Right Justify
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
            Left            =   2640
            TabIndex        =   158
            Text            =   "0"
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txtParaVacuum 
            Alignment       =   1  'Right Justify
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
            Left            =   2640
            TabIndex        =   157
            Text            =   "0"
            Top             =   1200
            Width           =   1335
         End
         Begin VB.TextBox txtParaVacuum 
            Alignment       =   1  'Right Justify
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
            Left            =   2640
            TabIndex        =   156
            Text            =   "0"
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtParaVacuum 
            Alignment       =   1  'Right Justify
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
            Left            =   2640
            TabIndex        =   155
            Text            =   "0"
            Top             =   2160
            Width           =   1335
         End
         Begin VB.TextBox txtParaVacuum 
            Alignment       =   1  'Right Justify
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
            Left            =   2640
            TabIndex        =   113
            Text            =   "2"
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label lbName 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "MFC Ratio"
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
            Index           =   104
            Left            =   1425
            TabIndex        =   477
            Top             =   5160
            Width           =   930
         End
         Begin VB.Label lbName 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "APC MFC Port(1~6)"
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
            Index           =   161
            Left            =   600
            TabIndex        =   278
            Top             =   3720
            Width           =   1755
         End
         Begin VB.Label lbName 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Keep Purge(0=max)"
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
            Index           =   157
            Left            =   600
            TabIndex        =   275
            Top             =   4680
            Width           =   1740
         End
         Begin VB.Label lbName 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "APC Interval"
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
            Index           =   155
            Left            =   240
            TabIndex        =   274
            Top             =   4200
            Width           =   2115
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "mSec"
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
            Index           =   154
            Left            =   4080
            TabIndex        =   273
            Top             =   4200
            Width           =   615
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Torr"
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
            Index           =   153
            Left            =   4080
            TabIndex        =   271
            Top             =   3120
            Width           =   330
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Gauge Zoom in"
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
            Index           =   152
            Left            =   240
            TabIndex        =   270
            Top             =   3240
            Width           =   1320
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Vacuum Gauge Offset"
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
            Index           =   28
            Left            =   240
            TabIndex        =   169
            Top             =   360
            Width           =   1920
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Torr"
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
            Index           =   29
            Left            =   4080
            TabIndex        =   168
            Top             =   360
            Width           =   330
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Release Open Delay"
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
            Index           =   37
            Left            =   240
            TabIndex        =   167
            Top             =   1320
            Width           =   1785
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "mSec"
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
            Index           =   45
            Left            =   4080
            TabIndex        =   166
            Top             =   840
            Width           =   510
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Throttle Valve Initial Pos."
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
            Index           =   70
            Left            =   240
            TabIndex        =   165
            Top             =   2280
            Width           =   2145
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "%"
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
            Index           =   71
            Left            =   4080
            TabIndex        =   164
            Top             =   2280
            Width           =   180
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Throttle Valve Full Delay"
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
            Index           =   72
            Left            =   240
            TabIndex        =   163
            Top             =   1800
            Width           =   2100
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "mSec"
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
            Index           =   73
            Left            =   4080
            TabIndex        =   162
            Top             =   1800
            Width           =   510
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Angle Valve Open Delay"
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
            Index           =   74
            Left            =   240
            TabIndex        =   161
            Top             =   840
            Width           =   2100
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "mSec"
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
            Index           =   75
            Left            =   4080
            TabIndex        =   160
            Top             =   1320
            Width           =   510
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Torr"
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
            Index           =   79
            Left            =   4080
            TabIndex        =   112
            Top             =   2640
            Width           =   330
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "APC Gauge Valve Limit"
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
            Index           =   78
            Left            =   240
            TabIndex        =   111
            Top             =   2760
            Width           =   2040
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "CT Alert Gate Value"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   2775
         Left            =   -66960
         TabIndex        =   43
         Top             =   9000
         Visible         =   0   'False
         Width           =   4695
         Begin VB.TextBox txtParaCTGate 
            Alignment       =   1  'Right Justify
            Height          =   390
            Index           =   13
            Left            =   3360
            TabIndex        =   77
            Text            =   "3"
            Top             =   3840
            Width           =   855
         End
         Begin VB.TextBox txtParaCTGate 
            Alignment       =   1  'Right Justify
            Height          =   390
            Index           =   12
            Left            =   3360
            TabIndex        =   76
            Text            =   "3"
            Top             =   3360
            Width           =   855
         End
         Begin VB.TextBox txtParaCTGate 
            Alignment       =   1  'Right Justify
            Height          =   390
            Index           =   11
            Left            =   3360
            TabIndex        =   75
            Text            =   "3"
            Top             =   2880
            Width           =   855
         End
         Begin VB.TextBox txtParaCTGate 
            Alignment       =   1  'Right Justify
            Height          =   390
            Index           =   10
            Left            =   3360
            TabIndex        =   74
            Text            =   "3"
            Top             =   2400
            Width           =   855
         End
         Begin VB.TextBox txtParaCTGate 
            Alignment       =   1  'Right Justify
            Height          =   390
            Index           =   9
            Left            =   3360
            TabIndex        =   73
            Text            =   "3"
            Top             =   1920
            Width           =   855
         End
         Begin VB.TextBox txtParaCTGate 
            Alignment       =   1  'Right Justify
            Height          =   390
            Index           =   8
            Left            =   3360
            TabIndex        =   72
            Text            =   "3"
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtParaCTGate 
            Alignment       =   1  'Right Justify
            Height          =   390
            Index           =   7
            Left            =   3360
            TabIndex        =   71
            Text            =   "3"
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtIntensityRef 
            Alignment       =   1  'Right Justify
            Height          =   390
            Index           =   1
            Left            =   3360
            TabIndex        =   69
            Text            =   "50"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtIntensityRef 
            Alignment       =   1  'Right Justify
            Height          =   390
            Index           =   0
            Left            =   1080
            TabIndex        =   66
            Text            =   "10"
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtParaCTGate 
            Alignment       =   1  'Right Justify
            Height          =   390
            Index           =   0
            Left            =   1080
            TabIndex        =   50
            Text            =   "1"
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtParaCTGate 
            Alignment       =   1  'Right Justify
            Height          =   390
            Index           =   1
            Left            =   1080
            TabIndex        =   49
            Text            =   "1"
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtParaCTGate 
            Alignment       =   1  'Right Justify
            Height          =   390
            Index           =   2
            Left            =   1080
            TabIndex        =   48
            Text            =   "1"
            Top             =   1920
            Width           =   855
         End
         Begin VB.TextBox txtParaCTGate 
            Alignment       =   1  'Right Justify
            Height          =   390
            Index           =   3
            Left            =   1080
            TabIndex        =   47
            Text            =   "1"
            Top             =   2400
            Width           =   855
         End
         Begin VB.TextBox txtParaCTGate 
            Alignment       =   1  'Right Justify
            Height          =   390
            Index           =   4
            Left            =   1080
            TabIndex        =   46
            Text            =   "1"
            Top             =   2880
            Width           =   855
         End
         Begin VB.TextBox txtParaCTGate 
            Alignment       =   1  'Right Justify
            Height          =   390
            Index           =   5
            Left            =   1080
            TabIndex        =   45
            Text            =   "1"
            Top             =   3360
            Width           =   855
         End
         Begin VB.TextBox txtParaCTGate 
            Alignment       =   1  'Right Justify
            Height          =   390
            Index           =   6
            Left            =   1080
            TabIndex        =   44
            Text            =   "1"
            Top             =   3840
            Width           =   855
         End
         Begin VB.TextBox txtCTAlertGateWeight 
            Height          =   390
            Left            =   2400
            TabIndex        =   128
            Text            =   "80"
            Top             =   4560
            Width           =   975
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   270
            Index           =   88
            Left            =   3480
            TabIndex        =   129
            Top             =   4680
            Width           =   210
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Alert Gate Weight"
            Height          =   270
            Index           =   87
            Left            =   360
            TabIndex        =   127
            Top             =   4560
            Width           =   1845
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00000080&
            X1              =   360
            X2              =   4560
            Y1              =   4320
            Y2              =   4320
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "CT 7b"
            Height          =   270
            Index           =   63
            Left            =   2640
            TabIndex        =   91
            Top             =   3840
            Width           =   645
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "CT 6b"
            Height          =   270
            Index           =   62
            Left            =   2640
            TabIndex        =   90
            Top             =   3360
            Width           =   645
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "CT 5b"
            Height          =   270
            Index           =   61
            Left            =   2640
            TabIndex        =   89
            Top             =   2880
            Width           =   645
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "CT 4b"
            Height          =   270
            Index           =   60
            Left            =   2640
            TabIndex        =   88
            Top             =   2400
            Width           =   645
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "CT 3b"
            Height          =   270
            Index           =   59
            Left            =   2640
            TabIndex        =   87
            Top             =   1920
            Width           =   645
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "CT 2b"
            Height          =   270
            Index           =   58
            Left            =   2640
            TabIndex        =   86
            Top             =   1440
            Width           =   645
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "CT 1b"
            Height          =   270
            Index           =   57
            Left            =   2640
            TabIndex        =   85
            Top             =   960
            Width           =   645
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "V"
            Height          =   270
            Index           =   56
            Left            =   4320
            TabIndex        =   84
            Top             =   3840
            Width           =   195
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "V"
            Height          =   270
            Index           =   55
            Left            =   4320
            TabIndex        =   83
            Top             =   3360
            Width           =   195
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "V"
            Height          =   270
            Index           =   54
            Left            =   4320
            TabIndex        =   82
            Top             =   2880
            Width           =   195
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "V"
            Height          =   270
            Index           =   53
            Left            =   4320
            TabIndex        =   81
            Top             =   2400
            Width           =   195
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "V"
            Height          =   270
            Index           =   52
            Left            =   4320
            TabIndex        =   80
            Top             =   1920
            Width           =   195
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "V"
            Height          =   270
            Index           =   51
            Left            =   4320
            TabIndex        =   79
            Top             =   1440
            Width           =   195
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "V"
            Height          =   270
            Index           =   50
            Left            =   4320
            TabIndex        =   78
            Top             =   960
            Width           =   195
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   270
            Index           =   49
            Left            =   4320
            TabIndex        =   70
            Top             =   600
            Width           =   210
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Power II"
            Height          =   270
            Index           =   48
            Left            =   2520
            TabIndex        =   68
            Top             =   480
            Width           =   825
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   270
            Index           =   47
            Left            =   2040
            TabIndex        =   67
            Top             =   600
            Width           =   210
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Power I"
            Height          =   270
            Index           =   46
            Left            =   240
            TabIndex        =   65
            Top             =   480
            Width           =   780
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "CT 1a"
            Height          =   270
            Index           =   30
            Left            =   360
            TabIndex        =   64
            Top             =   960
            Width           =   645
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "CT 2a"
            Height          =   270
            Index           =   31
            Left            =   360
            TabIndex        =   63
            Top             =   1440
            Width           =   645
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "CT 3a"
            Height          =   270
            Index           =   32
            Left            =   360
            TabIndex        =   62
            Top             =   1920
            Width           =   645
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "CT 4a"
            Height          =   270
            Index           =   33
            Left            =   360
            TabIndex        =   61
            Top             =   2400
            Width           =   645
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "CT 5a"
            Height          =   270
            Index           =   34
            Left            =   360
            TabIndex        =   60
            Top             =   2880
            Width           =   645
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "CT 6a"
            Height          =   270
            Index           =   35
            Left            =   360
            TabIndex        =   59
            Top             =   3360
            Width           =   645
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "CT 7a"
            Height          =   270
            Index           =   36
            Left            =   360
            TabIndex        =   58
            Top             =   3840
            Width           =   645
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "V"
            Height          =   270
            Index           =   38
            Left            =   2040
            TabIndex        =   57
            Top             =   960
            Width           =   195
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "V"
            Height          =   270
            Index           =   39
            Left            =   2040
            TabIndex        =   56
            Top             =   1440
            Width           =   195
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "V"
            Height          =   270
            Index           =   40
            Left            =   2040
            TabIndex        =   55
            Top             =   1920
            Width           =   195
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "V"
            Height          =   270
            Index           =   41
            Left            =   2040
            TabIndex        =   54
            Top             =   2400
            Width           =   195
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "V"
            Height          =   270
            Index           =   42
            Left            =   2040
            TabIndex        =   53
            Top             =   2880
            Width           =   195
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "V"
            Height          =   270
            Index           =   43
            Left            =   2040
            TabIndex        =   52
            Top             =   3360
            Width           =   195
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "V"
            Height          =   270
            Index           =   44
            Left            =   2040
            TabIndex        =   51
            Top             =   3840
            Width           =   195
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Gas "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   4095
         Left            =   -74880
         TabIndex        =   31
         Top             =   5040
         Width           =   7815
         Begin VB.TextBox txtParaGasErrorN 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   6
            Left            =   6360
            TabIndex        =   589
            Text            =   "0"
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox txtParaGasError 
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
            Left            =   5520
            TabIndex        =   588
            Text            =   "0"
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox txtParaGasBias 
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
            Left            =   4680
            TabIndex        =   587
            Text            =   "0"
            Top             =   3600
            Width           =   800
         End
         Begin VB.TextBox txtParaGasUnit 
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
            Left            =   3840
            TabIndex        =   586
            Text            =   "SLPM"
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox txtParaGasValue 
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
            Left            =   3000
            TabIndex        =   585
            Text            =   "10"
            Top             =   3600
            Width           =   735
         End
         Begin VB.TextBox txtParaGasAlias 
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
            Left            =   1560
            TabIndex        =   584
            Text            =   "NA"
            Top             =   3600
            Width           =   1335
         End
         Begin VB.CheckBox chkGasEnable 
            Caption         =   "Check1"
            Height          =   270
            Index           =   6
            Left            =   120
            TabIndex        =   582
            Top             =   3720
            Width           =   255
         End
         Begin VB.TextBox txtParaGasValue 
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
            Left            =   3000
            TabIndex        =   462
            Text            =   "10"
            Top             =   3120
            Width           =   735
         End
         Begin VB.TextBox txtParaGasAlias 
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
            Left            =   1560
            TabIndex        =   461
            Text            =   "NA"
            Top             =   3120
            Width           =   1335
         End
         Begin VB.TextBox txtParaGasBias 
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
            Left            =   4680
            TabIndex        =   460
            Text            =   "0"
            Top             =   3120
            Width           =   800
         End
         Begin VB.CheckBox chkGasEnable 
            Caption         =   "Check1"
            Height          =   270
            Index           =   5
            Left            =   120
            TabIndex        =   459
            Top             =   3240
            Width           =   255
         End
         Begin VB.TextBox txtParaGasUnit 
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
            Left            =   3840
            TabIndex        =   458
            Text            =   "SLPM"
            Top             =   3120
            Width           =   735
         End
         Begin VB.TextBox txtParaGasError 
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
            Left            =   5520
            TabIndex        =   457
            Text            =   "0"
            Top             =   3120
            Width           =   735
         End
         Begin VB.TextBox txtParaGasErrorN 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   5
            Left            =   6360
            TabIndex        =   456
            Text            =   "0"
            Top             =   3120
            Width           =   735
         End
         Begin VB.TextBox txtParaGasValue 
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
            Left            =   3000
            TabIndex        =   454
            Text            =   "10"
            Top             =   2640
            Width           =   735
         End
         Begin VB.TextBox txtParaGasAlias 
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
            Left            =   1560
            TabIndex        =   453
            Text            =   "NA"
            Top             =   2640
            Width           =   1335
         End
         Begin VB.TextBox txtParaGasBias 
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
            Left            =   4680
            TabIndex        =   452
            Text            =   "0"
            Top             =   2640
            Width           =   800
         End
         Begin VB.CheckBox chkGasEnable 
            Caption         =   "Check1"
            Height          =   270
            Index           =   4
            Left            =   120
            TabIndex        =   451
            Top             =   2760
            Width           =   255
         End
         Begin VB.TextBox txtParaGasUnit 
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
            Left            =   3840
            TabIndex        =   450
            Text            =   "SLPM"
            Top             =   2640
            Width           =   735
         End
         Begin VB.TextBox txtParaGasError 
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
            Left            =   5520
            TabIndex        =   449
            Text            =   "0"
            Top             =   2640
            Width           =   735
         End
         Begin VB.TextBox txtParaGasErrorN 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   4
            Left            =   6360
            TabIndex        =   448
            Text            =   "0"
            Top             =   2640
            Width           =   735
         End
         Begin VB.TextBox txtParaGasErrorN 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   3
            Left            =   6360
            TabIndex        =   292
            Text            =   "0"
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox txtParaGasErrorN 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   6360
            TabIndex        =   291
            Text            =   "0"
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox txtParaGasErrorN 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   6360
            TabIndex        =   290
            Text            =   "0"
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox txtParaGasErrorN 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   6360
            TabIndex        =   288
            Text            =   "0"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtParaGasError 
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
            Left            =   5520
            TabIndex        =   287
            Text            =   "0"
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox txtParaGasError 
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
            Left            =   5520
            TabIndex        =   286
            Text            =   "0"
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox txtParaGasError 
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
            Left            =   5520
            TabIndex        =   285
            Text            =   "0"
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox txtParaGasError 
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
            Left            =   5520
            TabIndex        =   284
            Text            =   "0"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtParaGasUnit 
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
            Left            =   3840
            TabIndex        =   247
            Text            =   "SLPM"
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox txtParaGasUnit 
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
            Left            =   3840
            TabIndex        =   246
            Text            =   "%"
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox txtParaGasUnit 
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
            Left            =   3840
            TabIndex        =   245
            Text            =   "SLPM"
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox txtParaGasUnit 
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
            Left            =   3840
            TabIndex        =   244
            Text            =   "SLPM"
            Top             =   720
            Width           =   735
         End
         Begin VB.CheckBox chkGasEnable 
            Caption         =   "Check1"
            Height          =   270
            Index           =   3
            Left            =   120
            TabIndex        =   203
            Top             =   2280
            Width           =   255
         End
         Begin VB.CheckBox chkGasEnable 
            Caption         =   "Check1"
            Height          =   270
            Index           =   2
            Left            =   120
            TabIndex        =   202
            Top             =   1800
            Width           =   255
         End
         Begin VB.CheckBox chkGasEnable 
            Caption         =   "Check1"
            Height          =   270
            Index           =   1
            Left            =   120
            TabIndex        =   201
            Top             =   1320
            Width           =   255
         End
         Begin VB.CheckBox chkGasEnable 
            Caption         =   "Check1"
            Height          =   270
            Index           =   0
            Left            =   120
            TabIndex        =   200
            Top             =   840
            Width           =   255
         End
         Begin VB.TextBox txtParaGasBias 
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
            Left            =   4680
            TabIndex        =   174
            Text            =   "0"
            Top             =   2160
            Width           =   800
         End
         Begin VB.TextBox txtParaGasBias 
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
            Left            =   4680
            TabIndex        =   173
            Text            =   "0"
            Top             =   1680
            Width           =   800
         End
         Begin VB.TextBox txtParaGasBias 
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
            Left            =   4680
            TabIndex        =   172
            Text            =   "0"
            Top             =   1200
            Width           =   800
         End
         Begin VB.TextBox txtParaGasBias 
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
            Left            =   4680
            TabIndex        =   171
            Text            =   "0"
            Top             =   720
            Width           =   800
         End
         Begin VB.TextBox txtParaGasAlias 
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
            Left            =   1560
            TabIndex        =   131
            Text            =   "NA"
            Top             =   2160
            Width           =   1335
         End
         Begin VB.TextBox txtParaGasValue 
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
            Left            =   3000
            TabIndex        =   130
            Text            =   "10"
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox txtParaGasValue 
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
            Left            =   3000
            TabIndex        =   37
            Text            =   "100"
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox txtParaGasValue 
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
            Left            =   3000
            TabIndex        =   36
            Text            =   "10"
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox txtParaGasValue 
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
            Left            =   3000
            TabIndex        =   35
            Text            =   "30"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtParaGasAlias 
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
            Left            =   1560
            TabIndex        =   34
            Text            =   "Pump"
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtParaGasAlias 
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
            Left            =   1560
            TabIndex        =   33
            Text            =   "Ar"
            Top             =   1200
            Width           =   1335
         End
         Begin VB.TextBox txtParaGasAlias 
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
            Left            =   1560
            TabIndex        =   32
            Text            =   "N2"
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Gas7 Alias"
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
            Left            =   480
            TabIndex        =   583
            Top             =   3720
            Width           =   960
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Gas6 Alias"
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
            Index           =   173
            Left            =   480
            TabIndex        =   463
            Top             =   3240
            Width           =   960
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Gas5 Alias"
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
            Index           =   90
            Left            =   480
            TabIndex        =   455
            Top             =   2760
            Width           =   960
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Error(N)"
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
            Index           =   108
            Left            =   6480
            TabIndex        =   289
            Top             =   360
            Width           =   675
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Error(±%)"
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
            Index           =   107
            Left            =   5520
            TabIndex        =   283
            Top             =   360
            Width           =   825
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Alias"
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
            Index           =   138
            Left            =   1560
            TabIndex        =   243
            Top             =   360
            Width           =   435
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Bias(V)"
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
            Index           =   106
            Left            =   4680
            TabIndex        =   170
            Top             =   360
            Width           =   645
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Gas4 Alias"
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
            Index           =   91
            Left            =   480
            TabIndex        =   132
            Top             =   2280
            Width           =   960
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Unit"
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
            Index           =   18
            Left            =   3840
            TabIndex        =   42
            Top             =   360
            Width           =   345
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Max"
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
            Index           =   17
            Left            =   3000
            TabIndex        =   41
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Gas3 Alias"
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
            Index           =   9
            Left            =   480
            TabIndex        =   40
            Top             =   1800
            Width           =   960
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Gas2 Alias"
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
            Index           =   5
            Left            =   480
            TabIndex        =   39
            Top             =   1320
            Width           =   960
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Gas1 Alias"
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
            Index           =   4
            Left            =   480
            TabIndex        =   38
            Top             =   840
            Width           =   960
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Heating"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   4335
         Left            =   -74880
         TabIndex        =   21
         Top             =   720
         Width           =   7815
         Begin VB.TextBox Txt_RValueRange 
            Height          =   390
            Left            =   6000
            TabIndex        =   593
            Text            =   "0.1"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtParaNormal 
            Height          =   390
            Index           =   8
            Left            =   2400
            TabIndex        =   281
            Text            =   "100"
            Top             =   3600
            Width           =   1335
         End
         Begin VB.TextBox txtParaHeat 
            Alignment       =   1  'Right Justify
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
            Left            =   2400
            TabIndex        =   261
            Text            =   "5"
            Top             =   3120
            Width           =   1335
         End
         Begin VB.TextBox txtPropertyCoeff 
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
            Left            =   6600
            TabIndex        =   217
            Text            =   "1"
            Top             =   4320
            Width           =   615
         End
         Begin VB.TextBox txtPropertyCoeff 
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
            Left            =   5760
            TabIndex        =   216
            Text            =   "1"
            Top             =   4320
            Width           =   615
         End
         Begin VB.TextBox txtPropertyCoeff 
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
            Left            =   4920
            TabIndex        =   215
            Text            =   "1"
            Top             =   4320
            Width           =   615
         End
         Begin VB.TextBox txtPropertyCoeff 
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
            Left            =   4080
            TabIndex        =   214
            Text            =   "1"
            Top             =   4320
            Width           =   615
         End
         Begin VB.TextBox txtPropertyCoeff 
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
            Left            =   3240
            TabIndex        =   213
            Text            =   "1"
            Top             =   4320
            Width           =   615
         End
         Begin VB.TextBox txtParaHeat 
            Alignment       =   1  'Right Justify
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
            Left            =   3840
            TabIndex        =   198
            Text            =   "-18"
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtParaHeat 
            Alignment       =   1  'Right Justify
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
            Left            =   2400
            TabIndex        =   197
            Text            =   "1"
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtIntensityWeightS 
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
            Left            =   3240
            TabIndex        =   124
            Text            =   "100"
            Top             =   4560
            Width           =   615
         End
         Begin VB.TextBox txtIntensityWeightS 
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
            Left            =   4080
            TabIndex        =   123
            Text            =   "100"
            Top             =   4560
            Width           =   615
         End
         Begin VB.TextBox txtIntensityWeightS 
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
            Left            =   4920
            TabIndex        =   122
            Text            =   "100"
            Top             =   4560
            Width           =   615
         End
         Begin VB.TextBox txtIntensityWeightS 
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
            Left            =   5760
            TabIndex        =   121
            Text            =   "100"
            Top             =   4560
            Width           =   615
         End
         Begin VB.TextBox txtIntensityWeightS 
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
            Left            =   6600
            TabIndex        =   120
            Text            =   "100"
            Top             =   4560
            Width           =   615
         End
         Begin VB.TextBox txtParaHeat 
            Alignment       =   1  'Right Justify
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
            Left            =   2400
            TabIndex        =   115
            Text            =   "0"
            Top             =   2640
            Width           =   1335
         End
         Begin VB.TextBox txtParaHeat 
            Alignment       =   1  'Right Justify
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
            Left            =   2400
            TabIndex        =   109
            Text            =   "1"
            Top             =   2160
            Width           =   1335
         End
         Begin VB.TextBox txtIntensityWeight 
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
            Left            =   6600
            TabIndex        =   102
            Text            =   "100"
            Top             =   5040
            Width           =   615
         End
         Begin VB.TextBox txtIntensityWeight 
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
            Left            =   5760
            TabIndex        =   101
            Text            =   "100"
            Top             =   5040
            Width           =   615
         End
         Begin VB.TextBox txtIntensityWeight 
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
            Left            =   4920
            TabIndex        =   100
            Text            =   "100"
            Top             =   5040
            Width           =   615
         End
         Begin VB.TextBox txtIntensityWeight 
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
            Left            =   4080
            TabIndex        =   99
            Text            =   "100"
            Top             =   5040
            Width           =   615
         End
         Begin VB.TextBox txtIntensityWeight 
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
            Left            =   3240
            TabIndex        =   98
            Text            =   "100"
            Top             =   5040
            Width           =   615
         End
         Begin VB.TextBox txtParaHeat 
            Alignment       =   1  'Right Justify
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
            Left            =   2400
            TabIndex        =   29
            Text            =   "1"
            Top             =   1200
            Width           =   1335
         End
         Begin VB.TextBox txtParaHeat 
            Alignment       =   1  'Right Justify
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
            Left            =   3840
            TabIndex        =   28
            Text            =   "-18"
            Top             =   1200
            Width           =   1335
         End
         Begin VB.TextBox txtParaHeat 
            Alignment       =   1  'Right Justify
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
            Left            =   2400
            TabIndex        =   23
            Text            =   "600"
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txtParaHeat 
            Alignment       =   1  'Right Justify
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
            Left            =   2400
            TabIndex        =   22
            Text            =   "90"
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lbl_RValueRange 
            Caption         =   "RValRange:"
            Height          =   375
            Left            =   4680
            TabIndex        =   592
            Top             =   255
            Width           =   1455
         End
         Begin VB.Label lbName 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Intensity Limit"
            Height          =   270
            Index           =   137
            Left            =   600
            TabIndex        =   282
            Top             =   3600
            Width           =   1515
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Number of Banks"
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
            Index           =   125
            Left            =   240
            TabIndex        =   262
            Top             =   3240
            Width           =   1500
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Property Coefficient "
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
            Left            =   240
            TabIndex        =   212
            Top             =   4320
            Width           =   1755
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "TC Temp. Cvt."
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
            Index           =   123
            Left            =   240
            TabIndex        =   199
            Top             =   1800
            Width           =   1230
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Intensity Weight in Steady (%)"
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
            Left            =   240
            TabIndex        =   125
            Top             =   4560
            Width           =   2655
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "%"
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
            Index           =   83
            Left            =   3960
            TabIndex        =   116
            Top             =   2760
            Width           =   180
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Intensity Keep"
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
            Index           =   82
            Left            =   240
            TabIndex        =   114
            Top             =   2760
            Width           =   1245
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "mSec"
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
            Index           =   77
            Left            =   3960
            TabIndex        =   110
            Top             =   2280
            Width           =   510
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Preheat Timeout"
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
            Index           =   76
            Left            =   240
            TabIndex        =   108
            Top             =   2280
            Width           =   1425
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
            Left            =   6600
            TabIndex        =   97
            Top             =   4320
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
            Left            =   5760
            TabIndex        =   96
            Top             =   4320
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
            Left            =   4920
            TabIndex        =   95
            Top             =   4320
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
            Left            =   4080
            TabIndex        =   94
            Top             =   4320
            Width           =   555
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
            Left            =   3240
            TabIndex        =   93
            Top             =   4320
            Width           =   555
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Intensity Weight in Dynamic (%)"
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
            Left            =   240
            TabIndex        =   92
            Top             =   5040
            Width           =   2805
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Pyrometer Temp. Cvt."
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
            Left            =   240
            TabIndex        =   30
            Top             =   1320
            Width           =   1890
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "℃"
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
            Left            =   3960
            TabIndex        =   27
            Top             =   840
            Width           =   240
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "℃"
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
            Index           =   11
            Left            =   3960
            TabIndex        =   26
            Top             =   360
            Width           =   240
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Min Pyrometer Temp."
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
            Index           =   8
            Left            =   240
            TabIndex        =   25
            Top             =   840
            Width           =   1875
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Overheat Temp."
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
            Index           =   1
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Width           =   1365
         End
      End
      Begin VB.Frame fraAO 
         Caption         =   "AO"
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
         Height          =   7335
         Left            =   5760
         TabIndex        =   4
         Top             =   780
         Width           =   7335
         Begin MSComctlLib.Slider sldAO 
            Height          =   255
            Index           =   0
            Left            =   4920
            TabIndex        =   11
            Top             =   240
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Max             =   5
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgAO 
            Height          =   6735
            Left            =   240
            TabIndex        =   7
            Top             =   480
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   11880
            _Version        =   393216
            Rows            =   17
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSComctlLib.Slider sldAO 
            Height          =   255
            Index           =   1
            Left            =   4920
            TabIndex        =   12
            Top             =   600
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Max             =   5
         End
         Begin MSComctlLib.Slider sldAO 
            Height          =   255
            Index           =   2
            Left            =   4920
            TabIndex        =   13
            Top             =   960
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Max             =   5
         End
         Begin MSComctlLib.Slider sldAO 
            Height          =   255
            Index           =   3
            Left            =   4920
            TabIndex        =   14
            Top             =   1320
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Max             =   5
         End
         Begin MSComctlLib.Slider sldAO 
            Height          =   255
            Index           =   4
            Left            =   4920
            TabIndex        =   15
            Top             =   1680
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Max             =   5
         End
         Begin MSComctlLib.Slider sldAO 
            Height          =   255
            Index           =   5
            Left            =   4920
            TabIndex        =   16
            Top             =   2040
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Max             =   5
         End
         Begin MSComctlLib.Slider sldAO 
            Height          =   255
            Index           =   6
            Left            =   4920
            TabIndex        =   17
            Top             =   2400
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Max             =   5
         End
         Begin MSComctlLib.Slider sldAO 
            Height          =   255
            Index           =   7
            Left            =   4920
            TabIndex        =   18
            Top             =   2760
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Max             =   5
         End
         Begin MSComctlLib.Slider sldAO 
            Height          =   255
            Index           =   8
            Left            =   4920
            TabIndex        =   204
            Top             =   3120
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Max             =   5
         End
         Begin MSComctlLib.Slider sldAO 
            Height          =   255
            Index           =   9
            Left            =   4920
            TabIndex        =   205
            Top             =   3480
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Max             =   5
         End
         Begin MSComctlLib.Slider sldAO 
            Height          =   255
            Index           =   10
            Left            =   4920
            TabIndex        =   206
            Top             =   3840
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Max             =   5
         End
         Begin MSComctlLib.Slider sldAO 
            Height          =   255
            Index           =   11
            Left            =   4920
            TabIndex        =   207
            Top             =   4200
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Max             =   5
         End
         Begin MSComctlLib.Slider sldAO 
            Height          =   255
            Index           =   12
            Left            =   4920
            TabIndex        =   208
            Top             =   4560
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Max             =   5
         End
         Begin MSComctlLib.Slider sldAO 
            Height          =   255
            Index           =   13
            Left            =   4920
            TabIndex        =   209
            Top             =   4920
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Max             =   5
         End
         Begin MSComctlLib.Slider sldAO 
            Height          =   255
            Index           =   14
            Left            =   4920
            TabIndex        =   210
            Top             =   5280
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Max             =   5
         End
         Begin MSComctlLib.Slider sldAO 
            Height          =   255
            Index           =   15
            Left            =   4920
            TabIndex        =   211
            Top             =   5640
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Max             =   5
         End
         Begin MSComctlLib.Slider sldAO 
            Height          =   255
            Index           =   16
            Left            =   4920
            TabIndex        =   470
            Top             =   6000
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Max             =   5
         End
         Begin MSComctlLib.Slider sldAO 
            Height          =   255
            Index           =   17
            Left            =   4920
            TabIndex        =   471
            Top             =   6360
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Max             =   5
         End
         Begin MSComctlLib.Slider sldAO 
            Height          =   255
            Index           =   18
            Left            =   4920
            TabIndex        =   472
            Top             =   6720
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Max             =   5
         End
         Begin MSComctlLib.Slider sldAO 
            Height          =   255
            Index           =   19
            Left            =   4920
            TabIndex        =   473
            Top             =   7080
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   1
            Max             =   5
         End
         Begin VB.Label Label4 
            Caption         =   "1=0~10V"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3000
            TabIndex        =   559
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "0=0~5V"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3000
            TabIndex        =   558
            Top             =   120
            Width           =   735
         End
         Begin VB.Shape shpAO 
            FillColor       =   &H00004040&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   19
            Left            =   4680
            Top             =   7080
            Width           =   255
         End
         Begin VB.Shape shpAO 
            FillColor       =   &H00004040&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   18
            Left            =   4680
            Top             =   6720
            Width           =   255
         End
         Begin VB.Shape shpAO 
            FillColor       =   &H00004040&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   17
            Left            =   4680
            Top             =   6360
            Width           =   255
         End
         Begin VB.Shape shpAO 
            FillColor       =   &H00004040&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   16
            Left            =   4680
            Top             =   6000
            Width           =   255
         End
         Begin VB.Shape shpAO 
            FillColor       =   &H00004040&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   15
            Left            =   4680
            Top             =   5640
            Width           =   255
         End
         Begin VB.Shape shpAO 
            FillColor       =   &H00004040&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   14
            Left            =   4680
            Top             =   5280
            Width           =   255
         End
         Begin VB.Shape shpAO 
            FillColor       =   &H00004040&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   13
            Left            =   4680
            Top             =   4920
            Width           =   255
         End
         Begin VB.Shape shpAO 
            FillColor       =   &H00004040&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   12
            Left            =   4680
            Top             =   4560
            Width           =   255
         End
         Begin VB.Shape shpAO 
            FillColor       =   &H00004040&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   11
            Left            =   4680
            Top             =   4200
            Width           =   255
         End
         Begin VB.Shape shpAO 
            FillColor       =   &H00004040&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   10
            Left            =   4680
            Top             =   3840
            Width           =   255
         End
         Begin VB.Shape shpAO 
            FillColor       =   &H00004040&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   9
            Left            =   4680
            Top             =   3480
            Width           =   255
         End
         Begin VB.Shape shpAO 
            FillColor       =   &H00004040&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   8
            Left            =   4680
            Top             =   3120
            Width           =   255
         End
         Begin VB.Shape shpAO 
            FillColor       =   &H00004040&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   7
            Left            =   4680
            Top             =   2760
            Width           =   255
         End
         Begin VB.Shape shpAO 
            FillColor       =   &H00004040&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   6
            Left            =   4680
            Top             =   2400
            Width           =   255
         End
         Begin VB.Shape shpAO 
            FillColor       =   &H00004040&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   5
            Left            =   4680
            Top             =   2040
            Width           =   255
         End
         Begin VB.Shape shpAO 
            FillColor       =   &H00004040&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   4
            Left            =   4680
            Top             =   1680
            Width           =   255
         End
         Begin VB.Shape shpAO 
            FillColor       =   &H00004040&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   3
            Left            =   4680
            Top             =   1320
            Width           =   255
         End
         Begin VB.Shape shpAO 
            FillColor       =   &H00004040&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   2
            Left            =   4680
            Top             =   960
            Width           =   255
         End
         Begin VB.Shape shpAO 
            FillColor       =   &H00004040&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   1
            Left            =   4680
            Top             =   600
            Width           =   255
         End
         Begin VB.Shape shpAO 
            FillColor       =   &H00004040&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   0
            Left            =   4680
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame fraAI 
         Caption         =   "AI"
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
         Height          =   7335
         Left            =   120
         TabIndex        =   3
         Top             =   780
         Width           =   5535
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgAI 
            Height          =   6735
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   11880
            _Version        =   393216
            Rows            =   17
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame fraDO 
         Caption         =   "DO"
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
         Height          =   7335
         Left            =   -68400
         TabIndex        =   2
         Top             =   780
         Width           =   5655
         Begin VB.Frame Frame4 
            Height          =   6855
            Left            =   4200
            TabIndex        =   10
            Top             =   240
            Width           =   1335
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   31
               Left            =   840
               Picture         =   "frmConfiguration.frx":025D
               Top             =   5880
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   30
               Left            =   840
               Picture         =   "frmConfiguration.frx":03CF
               Top             =   5520
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   29
               Left            =   840
               Picture         =   "frmConfiguration.frx":0541
               Top             =   5160
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   28
               Left            =   840
               Picture         =   "frmConfiguration.frx":06B3
               Top             =   4800
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   27
               Left            =   840
               Picture         =   "frmConfiguration.frx":0825
               Top             =   4440
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   26
               Left            =   840
               Picture         =   "frmConfiguration.frx":0997
               Top             =   4080
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   25
               Left            =   840
               Picture         =   "frmConfiguration.frx":0B09
               Top             =   3720
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   24
               Left            =   840
               Picture         =   "frmConfiguration.frx":0C7B
               Top             =   3360
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   23
               Left            =   840
               Picture         =   "frmConfiguration.frx":0DED
               Top             =   3000
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   22
               Left            =   840
               Picture         =   "frmConfiguration.frx":0F5F
               Top             =   2640
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   21
               Left            =   840
               Picture         =   "frmConfiguration.frx":10D1
               Top             =   2280
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   20
               Left            =   840
               Picture         =   "frmConfiguration.frx":1243
               Top             =   1920
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   19
               Left            =   840
               Picture         =   "frmConfiguration.frx":13B5
               Top             =   1560
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   18
               Left            =   840
               Picture         =   "frmConfiguration.frx":1527
               Top             =   1200
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   17
               Left            =   840
               Picture         =   "frmConfiguration.frx":1699
               Top             =   840
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   16
               Left            =   840
               Picture         =   "frmConfiguration.frx":180B
               Top             =   480
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   15
               Left            =   240
               Picture         =   "frmConfiguration.frx":197D
               Top             =   5880
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   14
               Left            =   240
               Picture         =   "frmConfiguration.frx":1AEF
               Top             =   5520
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   13
               Left            =   240
               Picture         =   "frmConfiguration.frx":1C61
               Top             =   5160
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   12
               Left            =   240
               Picture         =   "frmConfiguration.frx":1DD3
               Top             =   4800
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   11
               Left            =   240
               Picture         =   "frmConfiguration.frx":1F45
               Top             =   4440
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   10
               Left            =   240
               Picture         =   "frmConfiguration.frx":20B7
               Top             =   4080
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   9
               Left            =   240
               Picture         =   "frmConfiguration.frx":2229
               Top             =   3720
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   8
               Left            =   240
               Picture         =   "frmConfiguration.frx":239B
               Top             =   3360
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   7
               Left            =   240
               Picture         =   "frmConfiguration.frx":250D
               Top             =   3000
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   6
               Left            =   240
               Picture         =   "frmConfiguration.frx":267F
               Top             =   2640
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   5
               Left            =   240
               Picture         =   "frmConfiguration.frx":27F1
               Top             =   2280
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   4
               Left            =   240
               Picture         =   "frmConfiguration.frx":2963
               Top             =   1920
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   3
               Left            =   240
               Picture         =   "frmConfiguration.frx":2AD5
               Top             =   1560
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   2
               Left            =   240
               Picture         =   "frmConfiguration.frx":2C47
               Top             =   1200
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   1
               Left            =   240
               Picture         =   "frmConfiguration.frx":2DB9
               Top             =   840
               Width           =   300
            End
            Begin VB.Image imgDO 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   0
               Left            =   240
               Picture         =   "frmConfiguration.frx":2F2B
               Top             =   480
               Width           =   300
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgDO 
            Height          =   6735
            Left            =   360
            TabIndex        =   8
            Top             =   360
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   11880
            _Version        =   393216
            Rows            =   17
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame fraDI 
         Caption         =   "DI"
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
         Height          =   7335
         Left            =   -74520
         TabIndex        =   1
         Top             =   780
         Width           =   6015
         Begin VB.Frame Frame3 
            Height          =   7000
            Left            =   4200
            TabIndex        =   9
            Top             =   240
            Width           =   1695
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   31
               Left            =   1080
               Picture         =   "frmConfiguration.frx":309D
               Top             =   5880
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   30
               Left            =   1080
               Picture         =   "frmConfiguration.frx":320F
               Top             =   5520
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   29
               Left            =   1080
               Picture         =   "frmConfiguration.frx":3381
               Top             =   5160
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   28
               Left            =   1080
               Picture         =   "frmConfiguration.frx":34F3
               Top             =   4800
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   27
               Left            =   1080
               Picture         =   "frmConfiguration.frx":3665
               Top             =   4440
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   26
               Left            =   1080
               Picture         =   "frmConfiguration.frx":37D7
               Top             =   4080
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   25
               Left            =   1080
               Picture         =   "frmConfiguration.frx":3949
               Top             =   3720
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   24
               Left            =   1080
               Picture         =   "frmConfiguration.frx":3ABB
               Top             =   3360
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   23
               Left            =   1080
               Picture         =   "frmConfiguration.frx":3C2D
               Top             =   3000
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   22
               Left            =   1080
               Picture         =   "frmConfiguration.frx":3D9F
               Top             =   2640
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   21
               Left            =   1080
               Picture         =   "frmConfiguration.frx":3F11
               Top             =   2280
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   20
               Left            =   1080
               Picture         =   "frmConfiguration.frx":4083
               Top             =   1920
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   19
               Left            =   1080
               Picture         =   "frmConfiguration.frx":41F5
               Top             =   1560
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   18
               Left            =   1080
               Picture         =   "frmConfiguration.frx":4367
               Top             =   1200
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   17
               Left            =   1080
               Picture         =   "frmConfiguration.frx":44D9
               Top             =   840
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   16
               Left            =   1080
               Picture         =   "frmConfiguration.frx":464B
               Top             =   480
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   15
               Left            =   480
               Picture         =   "frmConfiguration.frx":47BD
               Top             =   5880
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   14
               Left            =   480
               Picture         =   "frmConfiguration.frx":492F
               Top             =   5520
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   13
               Left            =   480
               Picture         =   "frmConfiguration.frx":4AA1
               Top             =   5160
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   12
               Left            =   480
               Picture         =   "frmConfiguration.frx":4C13
               Top             =   4800
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   11
               Left            =   480
               Picture         =   "frmConfiguration.frx":4D85
               Top             =   4440
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   10
               Left            =   480
               Picture         =   "frmConfiguration.frx":4EF7
               Top             =   4080
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   9
               Left            =   480
               Picture         =   "frmConfiguration.frx":5069
               Top             =   3720
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   8
               Left            =   480
               Picture         =   "frmConfiguration.frx":51DB
               Top             =   3360
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   7
               Left            =   480
               Picture         =   "frmConfiguration.frx":534D
               Top             =   3000
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   6
               Left            =   480
               Picture         =   "frmConfiguration.frx":54BF
               Top             =   2640
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   5
               Left            =   480
               Picture         =   "frmConfiguration.frx":5631
               Top             =   2280
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   4
               Left            =   480
               Picture         =   "frmConfiguration.frx":57A3
               Top             =   1920
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   3
               Left            =   480
               Picture         =   "frmConfiguration.frx":5915
               Top             =   1560
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   2
               Left            =   480
               Picture         =   "frmConfiguration.frx":5A87
               Top             =   1200
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   1
               Left            =   480
               Picture         =   "frmConfiguration.frx":5BF9
               Top             =   840
               Width           =   300
            End
            Begin VB.Image imgDI 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   0
               Left            =   480
               Picture         =   "frmConfiguration.frx":5D6B
               Top             =   480
               Width           =   300
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgDI 
            Height          =   6735
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   11880
            _Version        =   393216
            Rows            =   17
            AllowBigSelection=   0   'False
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
   End
   Begin VB.Timer tmrAIO 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   14040
      Top             =   7680
   End
   Begin MSComDlg.CommonDialog cdFile 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   14520
      Top             =   8520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
      BaudRate        =   19200
   End
   Begin VB.Image imgOn 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   720
      Picture         =   "frmConfiguration.frx":5EDD
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgOff 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   360
      Picture         =   "frmConfiguration.frx":604F
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
End
Attribute VB_Name = "frmConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'<<<<<<<<<<<< DI / 9114 >>>>>>>>>>>>
'Name define of DI
Const NAME_SYSTEM_ALARM = "System Alarm"
Const NAME_SYSTEM_READY = "System Ready"
Const NAME_POWER_INPUT220 = "Power Input 220V"
Const NAME_MC_STATUS = "MC ON/OFF"
Const NAME_EMS_ALARM = "EMS Alarm"
Const NAME_CHAMBER_OVERHEAT = "Chamber OverHeat"
Const NAME_AIR_ALARM = "Air Pressure"
Const NAME_WATER_ALARM = "Water Flow"
Const NAME_DOOR_OPEN = "Door Unclamp"
Const NAME_DOOR_CLOSE = "Door Close"
Const NAME_DOOR_OPEN_SENSOR = "Door Open Sensor"
Const NAME_DOOR_CLOSE_SENSOR = "Door Close Sensor"
Const NAME_DI_DOOR_CLAMP = "Door Clamp"
Const NAME_VAC_SENSOR_R = "Vacuum Sensor H"
Const NAME_VAC_SENSOR_L = "Vacuum Sensor L"
Const NAME_CHAMBER_SENSOR_R = "Chamber Safe H"
Const NAME_CHAMBER_SENSOR_L = "Chamber Safe L"
Const NAME_CT1 = "CT1"
Const NAME_CT2 = "CT2"
Const NAME_CT3 = "CT3"
Const NAME_CT4 = "CT4"
Const NAME_CT5 = "CT5"
Const NAME_CT6 = "CT6"
Const NAME_ARM_INFRONT = "Arm Front Sensor"
Const NAME_ARM_INREAR = "Arm Rear Sensor"
Const NAME_COVER_ALARM1 = "Cover Alarm1"
Const NAME_COVER_SERVO_RDY1 = "Cover Servo Ready1"
Const NAME_COVER_ORIG_RDY1 = "Cover Origin Ready1"
Const NAME_COVER_ISMOVING1 = "Cover Is Moving1"
Const NAME_COVER_UP_INP1 = "Cover Up Inpos1"
Const NAME_COVER_DOWN_INP1 = "Cover Down Inpos1"
Const NAME_COVER_ALARM2 = "Cover Alarm2"
Const NAME_COVER_SERVO_RDY2 = "Cover Servo Ready2"
Const NAME_COVER_ORIG_RDY2 = "Cover Origin Ready2"
Const NAME_COVER_ISMOVING2 = "Cover Is Moving2"
Const NAME_COVER_UP_INP2 = "Cover Up Inpos2"
Const NAME_COVER_DOWN_INP2 = "Cover Down Inpos2"
Const NAME_PUMP_ALARM = "Pump Alarm"
'========================================================================================================

'<<<<<<<<<<<< DO / 9114 >>>>>>>>>>>>
'Name define of DO
'Const NAME_SYSTEM_ALARM = "System Alarm"
Const NAME_BUZZER = "Buzzer Stop"
'Const NAME_WATER_AIR = "Water/Air stop"
Const NAME_PC_CHECK1 = "PC check1(WDT)"
Const NAME_PC_CHECK2 = "PC check2(WDT)"
Const NAME_MFC_SERVO_1 = "MFC1 Valve Servo"
Const NAME_MFC_SERVO_2 = "MFC2 Valve Servo"
Const NAME_MFC_SERVO_3 = "MFC3 Valve Servo"
Const NAME_MFC_SERVO_4 = "MFC4 Valve Servo"
Const NAME_MFC_SERVO_5 = "MFC5 Valve Servo"
Const NAME_MFC_SERVO_6 = "MFC6 Valve Servo"
Const NAME_GAS_VALVE_1 = "Gas1 Valve ON"
Const NAME_GAS_VALVE_2 = "Gas2 Valve ON"
Const NAME_GAS_VALVE_3 = "Gas3 Valve ON"
Const NAME_GAS_VALVE_4 = "Gas4 Valve ON"
Const NAME_GAS_VALVE_5 = "Gas5 Valve ON"
Const NAME_GAS_VALVE_6 = "Gas6 Valve ON"
Const NAME_CDA_VALVE = "Lamp Cooling"
Const NAME_PUMP_POWER = "Pump Power"
Const NAME_ANGLE_VALVE = "Angle Value"
Const NAME_RELEASE_VALVE = "Release Value"
Const NAME_APC_GAUGE_VALVE = "Gauge Value"  'For APC Gauge
Const NAME_APC_GAUGE_ANGLE = "Gauge Angle"
Const NAME_DOOR_OPEN_VALVE = "Door Open Valve"
Const NAME_DOOR_CLOSE_VALVE = "Door Close Valve"
Const NAME_DOOR_CLAMP = "Door Clamp"
Const NAME_ALARM_LIGHT_RED = "R Light"
Const NAME_ALARM_LIGHT_YELLOW = "Y Light"
Const NAME_ALARM_LIGHT_BLUE = "B Light"
Const NAME_ALARM_LIGHT_GREEN = "G Light"
Const NAME_EXHAUST = "Exhaust"
Const NAME_ARM_FRONT = "Hold Safety"
Const NAME_ARM_REAR = "Arm Move Rear"
Const NAME_COVER_ARESET = "Cover Alarm Reset"
Const NAME_COVER_SERVO = "Cover Servo ON"
Const NAME_COVER_ORGIN = "Cover Orgin ON"
Const NAME_COVER_MOVE = "Cover Move"
Const NAME_COVER_POS_01 = "Cover POS_01"

'========================================================================================================

'<<<<<<<<<<<< AI / 9114 >>>>>>>>>>>>
'Name define of AI
Const NAME_CT_01 = "CT1"
Const NAME_CT_02 = "CT2"
Const NAME_CT_03 = "CT3"
Const NAME_CT_04 = "CT4"
Const NAME_CT_05 = "CT5"
Const NAME_CT_06 = "CT6"
Const NAME_CT_07 = "CT7"
Const NAME_MFC_READ_1 = "MFC1 Read"
Const NAME_MFC_READ_2 = "MFC2 Read"
Const NAME_MFC_READ_3 = "MFC3 Read"
Const NAME_MFC_READ_4 = "MFC4 Read"
Const NAME_MFC_READ_5 = "MFC5 Read"
Const NAME_MFC_READ_6 = "MFC6 Read"
Const NAME_VACUUM_GAUGE = "Vacuum Gauge"
Const NAME_VACUUM_GAUGE2 = "Vacuum gauge 2"
Const NAME_OXYGEN_GAUGE = "Oxygen Gauge"
Const NAME_TC_CVT_1 = "TC Converter1"
Const NAME_PYROMETER = "Pyrometer" 'Reserved
'Rev4.1.4
Const NAME_TC_WAF_1 = "TC Wafer1"
Const NAME_TC_WAF_2 = "TC Wafer2"
Const NAME_TC_WAF_3 = "TC Wafer3"
Const NAME_TC_WAF_4 = "TC Wafer4"
Const NAME_TC_WAF_5 = "TC Wafer5"
'========================================================================================================

'<<<<<<<<<<<< AO / 6208 >>>>>>>>>>>>
'Name define of AO
'Heating
Const NAME_SET_SCR_TBC = "SCR-TBC Set"
Const NAME_SET_SCR_TR = "SCR-TR Set"
Const NAME_SET_SCR_TL = "SCR-TL Set"
Const NAME_SET_SCR_BF = "SCR-BF Set"
Const NAME_SET_SCR_BR = "SCR-BR Set"
Const NAME_SET_SCR_6 = "SCR-6 Set"
Const NAME_SET_SCR_7 = "SCR-7 Set"
Const NAME_SET_SCR_8 = "SCR-8 Set"
Const NAME_SET_SCR_9 = "SCR-9 Set"
Const NAME_SET_SCR_10 = "SCR-10 Set"
Const NAME_SET_SCR_11 = "SCR-11 Set"
Const NAME_SET_SCR_12 = "SCR-12 Set"
Const NAME_SET_SCR_13 = "SCR-13 Set"
Const NAME_SET_SCR_14 = "SCR-14 Set"
Const NAME_SET_SCR_15 = "SCR-15 Set"
Const NAME_SET_SCR_16 = "SCR-16 Set"
Const NAME_SET_SCR_17 = "SCR-17 Set"
'Gas
Const NAME_MFC_SET_1 = "MFC1 Set"
Const NAME_MFC_SET_2 = "MFC2 Set"
Const NAME_MFC_SET_3 = "MFC3 Set"
Const NAME_MFC_SET_4 = "MFC4 Set"
Const NAME_MFC_SET_5 = "MFC5 Set"
Const NAME_MFC_SET_6 = "MFC6 Set"
'========================================================================================================

'<<<<<<<<<<<< AI / PCI-1719HGU >>>>>>>>>>>>
'Name define of AI
Const NAME_TC_01 = "TC1"
Const NAME_TC_02 = "TC2"
Const NAME_TC_03 = "TC3"
Const NAME_TC_04 = "TC4"
Const NAME_TC_05 = "TC5"
Const NAME_TC_06 = "TC6"
Const NAME_TC_07 = "TC7"
Const NAME_TC_08 = "TC8"
Const NAME_TC_09 = "TC9"
Const NAME_TC_10 = "TC10"
Const NAME_TC_11 = "TC11"
Const NAME_TC_12 = "TC12"
Const NAME_TC_13 = "TC13"
Const NAME_TC_14 = "TC14"
Const NAME_TC_15 = "TC15"
Const NAME_TC_16 = "TC16"
Const NAME_TC_17 = "TC17"
Const NAME_TC_18 = "TC18"
Const NAME_TC_19 = "TC19"
Const NAME_TC_20 = "TC20"
Const NAME_TC_21 = "TC21"
Const NAME_TC_22 = "TC22"
Const NAME_TC_23 = "TC23"
Const NAME_TC_24 = "TC24"

Const NAME_O2_01 = "O2 S"
Const NAME_PS_01 = "PS"
Const NAME_P = "P"
'Name define of AO Type
Const NAME_AOT_01 = "0=5V"
Const NAME_AOT_02 = "1=10V"

'WatchDog

Dim blnSave  As Boolean
Dim blnReceivedPM  As Boolean
Dim intSendPMCount As Integer
Dim NewGasCount As Integer

Private Sub Activate_Click()
frmActivate.Show


End Sub

Private Sub chkAutoMode_Click()
'    If chkAutoMode.Value = 1 Then
'        frmUDP.OpenUDP
'    End If
End Sub

Private Sub chkAutoDoor_Click()
    If chkAutoDoor.value = 1 Then
        If SysDI.IsChamberGaugeL = 0 Then
            chkAutoDoor.value = 0
            ShowMessageOK "腔體處於真空狀態,請先釋放壓力!"
            Exit Sub
        End If
        frmUDP.wsServer.sendData "$DR=1,"
    Else
        frmUDP.wsServer.sendData "$DR=0,"
    End If
End Sub


Private Sub chkCIM_Click()
    If chkCIM.value = 1 Then
        frmCIM.Send "$CMD=" & CStr(cmbCIMPort.ListIndex)
    Else
        frmCIM.Send "$CMD=0"
    End If
End Sub

Private Sub chkUseTempMeter_Click()
    If chkUseTempMeter.value = 1 Then
        frmUDP.OpenUDP
        gbblnSendRecipe = True
        
    End If
    Para.UseTempMeter = chkUseTempMeter.value
End Sub

Private Sub ChkWriteOffsetToTCM_Click()
    If ChkWriteOffsetToTCM.value = 1 Then
        OpenOffsetValue
    End If
End Sub

Private Sub Cmd_TcOffset_Click()
frmTcOffset.Show
End Sub

Private Sub cmdAutoUDP_Click()
    frmUDP.Show
End Sub


Private Sub cmdAz1_Click()
    frmAz1.Show
End Sub

Private Sub cmdAz2_Click()
    frmAz2.Show
End Sub

Private Sub cmdCIMUDP_Click()
    frmCIM.Show
End Sub

Private Sub cmdCustom_Click()
    frmLoginCustom.Show
    
End Sub

Private Sub cmdDCR_Click()
    frmDCR.Show
End Sub

Private Sub cmdModbusRTU_Click()
     frmModBusRtu.Show
End Sub

Private Sub Form_Activate()
    
    'fraAutoMode(4).Visible = IIf(gbintActiveModule_Auto = 1, True, False)
End Sub

Private Sub Form_Load()
    InitialForm
    InitialIO
    ParameterOpen
    OpenOffsetValue
    OpenFunctionIni
    tabMain.TabVisible(1) = False
    If gbintLoginRight = 5 Then
    lbHdOffset.Visible = True
    TxthdOffset.Visible = True
    TxthdOffset.text = CommnonReadini("Special_Setting", "Hold_Offset", App.Path + ProcDict_Path)
    TxthdTimes.text = CommnonReadini("Special_Setting", "Hold_Times", App.Path + ProcDict_Path)
    End If
    Activate.Visible = False
   If GbTcoffset_Switch = 1 Then
   Cmd_TcOffset.Visible = True
   Else
   Cmd_TcOffset.Visible = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
'    ResetDO
'    ResetAO
'    SetDO gblngDO_PC_Check1, True
'    SetDO gblngDO_PC_Check2, True
'    Ixud_DriverClose
   
End Sub

Public Sub InitialIO()
    Dim i As Integer
    
    'Initial DO
    'lngDO_WaterAirStop = 0
    gblngDO_PC_Check1 = 0
    gblngDO_PC_Check2 = 1
    gblngDO_BuzzerStop = -3
    gblngDO_MFC_ValveServo1 = -4
    gblngDO_MFC_ValveServo2 = -5
    gblngDO_MFC_ValveServo3 = -6
    gblngDO_MFC_ValveServo4 = -7
    gblngDO_MFC_ValveServo5 = -7
    gblngDO_MFC_ValveServo6 = -7
    gblngDO_GasValve1 = -8
    gblngDO_GasValve2 = -9
    gblngDO_GasValve3 = -10
    gblngDO_GasValve4 = -11
    gblngDO_GasValve5 = -11
    gblngDO_GasValve6 = -11
    gblngDO_AngleValve = -12
    gblngDO_ReleaseValve = -13
    gblngDO_PumpPower = -14
    gblngDO_SystemAlarm = -15
    gblngDO_APCGaugeValve = -11
    gblngDO_APCGaugeAngle = -11
    gblngDO_AlarmRed = -13
    gblngDO_AlarmYellow = -14
    gblngDO_AlarmBlue = -15
    gblngDO_Exhaust = -15
    gblngDO_ARM_FRONT = -15
    
    'Initial DI
    gblngDI_SystemAlarm = -1
    gblngDI_SystemReady = -1
    'lngDI_PowerInput220 = 7
    gblngDI_EMS_Alarm = -2
    gblngDI_MC_Input = -3
    gblngDI_ChamberOverheat = -4
    gblngDI_AirAlarm = -5
    gblngDI_WaterAlarm = -6
    gblngDI_DoorOpen = -7
    gblngDI_DoorClose = -8
    gblngDI_DoorOpenSensor = -9
    gblngDI_DoorCloseSensor = -10
    gblngDI_DoorClamp = -8
    gblngDI_VacuumGateR = -11
    gblngDI_VacuumGateL = -12
    gblngDI_ChamberGateR = -15
    gblngDI_ChamberGateL = -15
    gblngDI_CT1 = -10
    gblngDI_CT2 = -11
    gblngDI_CT3 = -12
    gblngDI_CT4 = -13
    gblngDI_CT5 = -14
    gblngDI_CT6 = -15
    gblngDI_ARM_FRONT = -15
    gblngDI_ARM_REAR = -15
    
    
    'Initial AI
    gblngAI_CT_01 = -1
    gblngAI_CT_02 = -2
    gblngAI_CT_03 = -3
    gblngAI_CT_04 = -4
    gblngAI_CT_05 = -5
    gblngAI_CT_06 = -6
    gblngAI_CT_07 = -7
    For i = 0 To gbintMaxGasEnable
        gblngAI_MFC_Read(i) = -1 '8 + i
        
    Next i
'    gblngAI_MFC_Read(1) = 8
'    gblngAI_MFC_Read(2) = 9
'    gblngAI_MFC_Read(3) = 10
'    gblngAI_MFC_Read4 = 11
    gblngAI_Vacuum_Gauge = -12
    gblngAI_Vacuum_Gauge2 = -12
    gblngAI_Oxygen_Gauge = -12
    gblngAI_TC_Cvt1 = 7
    gblngAI_TC_Cvt2 = 8
    gblngAI_Pyrometer = 9
    'Rev4.1.4
    gblngAI_TCWafer1 = -11
    gblngAI_TCWafer2 = -12
    gblngAI_TCWafer3 = -13
    gblngAI_TCWafer4 = -14
    gblngAI_TCWafer5 = -15

    'Initial AO
    gblngAO_SCR_TBC = -1
    gblngAO_SCR_TR = -2
    gblngAO_SCR_TL = -3
    gblngAO_SCR_BF = -4
    gblngAO_SCR_BR = -5
    gblngAO_MFC1 = -11
    gblngAO_MFC2 = -12
    gblngAO_MFC3 = -13
    gblngAO_MFC4 = -14
    gblngAO_MFC5 = -14
    gblngAO_MFC6 = -14
    '120713 Josh
    gblngAO_SCR_6 = -6
    gblngAO_SCR_7 = -7
    gblngAO_SCR_8 = -8
    gblngAO_SCR_9 = -9
    gblngAO_SCR_10 = -10
    gblngAO_SCR_11 = -11
    gblngAO_SCR_12 = -12
    gblngAO_SCR_13 = -13
    gblngAO_SCR_14 = -14
    gblngAO_SCR_15 = -15
    gblngAO_SCR_16 = -16
    gblngAO_SCR_17 = -17
    
    Call ReloadIO
    
    gblngCheckPC = gblngDO_PC_Check1
        
End Sub

Public Sub InitialForm()
    Dim i As Integer
    
    With fgDI
        .Cols = 2
        .Rows = 50
        .ColWidth(0) = 800
        .ColWidth(1) = 3000
        .TextMatrix(0, 0) = "Port"
        .TextMatrix(0, 1) = "Description"
        For i = 1 To 34
            .RowHeight(i) = 360
            .TextMatrix(i, 0) = CStr(i - 1) & "     "
            .TextMatrix(i, 1) = "NA"
            .RowHeight(i) = 360
        Next i
        If gblngDI_ChamberOverheat >= 0 Then .TextMatrix(1 + gblngDI_ChamberOverheat, 1) = NAME_CHAMBER_OVERHEAT
        If gblngDI_AirAlarm >= 0 Then .TextMatrix(1 + gblngDI_AirAlarm, 1) = NAME_AIR_ALARM
        If gblngDI_WaterAlarm >= 0 Then .TextMatrix(1 + gblngDI_WaterAlarm, 1) = NAME_WATER_ALARM
        If gblngDI_SystemReady >= 0 Then .TextMatrix(1 + gblngDI_SystemReady, 1) = NAME_SYSTEM_READY
        If gblngDI_EMS_Alarm >= 0 Then .TextMatrix(1 + gblngDI_EMS_Alarm, 1) = NAME_EMS_ALARM
        If gblngDI_MC_Input >= 0 Then .TextMatrix(1 + gblngDI_MC_Input, 1) = NAME_MC_STATUS
        If gblngDI_DoorOpen >= 0 Then .TextMatrix(1 + gblngDI_DoorOpen, 1) = NAME_DOOR_OPEN
        If gblngDI_DoorClose >= 0 Then .TextMatrix(1 + gblngDI_DoorClose, 1) = NAME_DOOR_CLOSE
        If gblngDI_DoorOpenSensor >= 0 Then .TextMatrix(1 + gblngDI_DoorOpenSensor, 1) = NAME_DOOR_OPEN_SENSOR
        If gblngDI_DoorCloseSensor >= 0 Then .TextMatrix(1 + gblngDI_DoorCloseSensor, 1) = NAME_DOOR_CLOSE_SENSOR
        If gblngDI_DoorClamp >= 0 Then .TextMatrix(1 + gblngDI_DoorClamp, 1) = NAME_DI_DOOR_CLAMP
        If gblngDI_VacuumGateR >= 0 Then .TextMatrix(1 + gblngDI_VacuumGateR, 1) = NAME_VAC_SENSOR_R
        If gblngDI_VacuumGateL >= 0 Then .TextMatrix(1 + gblngDI_VacuumGateL, 1) = NAME_VAC_SENSOR_L
        If gblngDI_CT1 >= 0 Then .TextMatrix(1 + gblngDI_CT1, 1) = NAME_CT1
        If gblngDI_CT2 >= 0 Then .TextMatrix(1 + gblngDI_CT2, 1) = NAME_CT2
        If gblngDI_CT3 >= 0 Then .TextMatrix(1 + gblngDI_CT3, 1) = NAME_CT3
        If gblngDI_CT4 >= 0 Then .TextMatrix(1 + gblngDI_CT4, 1) = NAME_CT4
        If gblngDI_CT5 >= 0 Then .TextMatrix(1 + gblngDI_CT5, 1) = NAME_CT5
        If gblngDI_CT6 >= 0 Then .TextMatrix(1 + gblngDI_CT6, 1) = NAME_CT6
        'Add Chamber Safe
        If gblngDI_ChamberGateR >= 0 Then .TextMatrix(1 + gblngDI_ChamberGateR, 1) = NAME_CHAMBER_SENSOR_R
        If gblngDI_ChamberGateL >= 0 Then .TextMatrix(1 + gblngDI_ChamberGateL, 1) = NAME_CHAMBER_SENSOR_L
        'Add Cover Status
        If gblngDI_CoverAlarm1 >= 0 Then .TextMatrix(1 + gblngDI_CoverAlarm1, 1) = NAME_COVER_ALARM1
        If gblngDI_CoverServoRdy1 >= 0 Then .TextMatrix(1 + gblngDI_CoverServoRdy1, 1) = NAME_COVER_SERVO_RDY1
        If gblngDI_CoverOrigRdy1 >= 0 Then .TextMatrix(1 + gblngDI_CoverOrigRdy1, 1) = NAME_COVER_ORIG_RDY1
        If gblngDI_CoverIsMoving1 >= 0 Then .TextMatrix(1 + gblngDI_CoverOrigRdy1, 1) = NAME_COVER_ISMOVING1
        If gblngDI_CoverUpInpos1 >= 0 Then .TextMatrix(1 + gblngDI_CoverUpInpos1, 1) = NAME_COVER_UP_INP1
        If gblngDI_CoverDownInpos1 >= 0 Then .TextMatrix(1 + gblngDI_CoverDownInpos1, 1) = NAME_COVER_DOWN_INP1
        If gblngDI_CoverAlarm2 >= 0 Then .TextMatrix(1 + gblngDI_CoverAlarm2, 1) = NAME_COVER_ALARM2
        If gblngDI_CoverServoRdy2 >= 0 Then .TextMatrix(1 + gblngDI_CoverServoRdy2, 1) = NAME_COVER_SERVO_RDY2
        If gblngDI_CoverOrigRdy2 >= 0 Then .TextMatrix(1 + gblngDI_CoverOrigRdy2, 1) = NAME_COVER_ORIG_RDY2
        If gblngDI_CoverIsMoving2 >= 0 Then .TextMatrix(1 + gblngDI_CoverOrigRdy2, 1) = NAME_COVER_ISMOVING2
        If gblngDI_CoverUpInpos2 >= 0 Then .TextMatrix(1 + gblngDI_CoverUpInpos2, 1) = NAME_COVER_UP_INP2
        If gblngDI_CoverDownInpos2 >= 0 Then .TextMatrix(1 + gblngDI_CoverDownInpos2, 1) = NAME_COVER_DOWN_INP2
        'Add pumpalarm
        If gblngDI_PumpAlarm >= 0 Then .TextMatrix(1 + gblngDI_PumpAlarm, 1) = NAME_PUMP_ALARM
    End With
    
    cmbIOList(0).AddItem NAME_SYSTEM_ALARM
    cmbIOList(0).AddItem NAME_CHAMBER_OVERHEAT
    cmbIOList(0).AddItem NAME_AIR_ALARM
    cmbIOList(0).AddItem NAME_WATER_ALARM
    cmbIOList(0).AddItem NAME_SYSTEM_READY
    cmbIOList(0).AddItem NAME_EMS_ALARM
    cmbIOList(0).AddItem NAME_MC_STATUS
    cmbIOList(0).AddItem NAME_DOOR_OPEN
    cmbIOList(0).AddItem NAME_DOOR_CLOSE
    cmbIOList(0).AddItem NAME_DOOR_OPEN_SENSOR
    cmbIOList(0).AddItem NAME_DOOR_CLOSE_SENSOR
    cmbIOList(0).AddItem NAME_DI_DOOR_CLAMP
    cmbIOList(0).AddItem NAME_VAC_SENSOR_R
    cmbIOList(0).AddItem NAME_VAC_SENSOR_L
    cmbIOList(0).AddItem NAME_CHAMBER_SENSOR_R
    cmbIOList(0).AddItem NAME_CHAMBER_SENSOR_L
    cmbIOList(0).AddItem NAME_CT1
    cmbIOList(0).AddItem NAME_CT2
    cmbIOList(0).AddItem NAME_CT3
    cmbIOList(0).AddItem NAME_CT4
    cmbIOList(0).AddItem NAME_CT5
    cmbIOList(0).AddItem NAME_CT6
    cmbIOList(0).AddItem NAME_ARM_INFRONT
    cmbIOList(0).AddItem NAME_COVER_ALARM1
    cmbIOList(0).AddItem NAME_COVER_SERVO_RDY1
    cmbIOList(0).AddItem NAME_COVER_ORIG_RDY1
    cmbIOList(0).AddItem NAME_COVER_ISMOVING1
    cmbIOList(0).AddItem NAME_COVER_UP_INP1
    cmbIOList(0).AddItem NAME_COVER_DOWN_INP1
    cmbIOList(0).AddItem NAME_COVER_ALARM2
    cmbIOList(0).AddItem NAME_COVER_SERVO_RDY2
    cmbIOList(0).AddItem NAME_COVER_ORIG_RDY2
    cmbIOList(0).AddItem NAME_COVER_ISMOVING2
    cmbIOList(0).AddItem NAME_COVER_UP_INP2
    cmbIOList(0).AddItem NAME_COVER_DOWN_INP2
    cmbIOList(0).AddItem NAME_PUMP_ALARM
    cmbIOList(0).AddItem "NA"
    
    With fgDO
        .Cols = 2
        .Rows = 33
         .ColWidth(0) = 800
        .ColWidth(1) = 3000
       .TextMatrix(0, 0) = "Port"
        .TextMatrix(0, 1) = "Description"
        For i = 1 To 32
            .RowHeight(i) = 360
            .TextMatrix(i, 0) = CStr(i - 1) & "     "
            .TextMatrix(i, 1) = "NA"
        Next i
        If gblngDO_PC_Check1 >= 0 Then .TextMatrix(1 + gblngDO_PC_Check1, 1) = NAME_PC_CHECK1
        If gblngDO_PC_Check2 >= 0 Then .TextMatrix(1 + gblngDO_PC_Check2, 1) = NAME_PC_CHECK2
        If gblngDO_MFC_ValveServo1 >= 0 Then .TextMatrix(1 + gblngDO_MFC_ValveServo1, 1) = NAME_MFC_SERVO_1
        If gblngDO_MFC_ValveServo2 >= 0 Then .TextMatrix(1 + gblngDO_MFC_ValveServo2, 1) = NAME_MFC_SERVO_2
        If gblngDO_MFC_ValveServo3 >= 0 Then .TextMatrix(1 + gblngDO_MFC_ValveServo3, 1) = NAME_MFC_SERVO_3
        If gblngDO_MFC_ValveServo4 >= 0 Then .TextMatrix(1 + gblngDO_MFC_ValveServo4, 1) = NAME_MFC_SERVO_4
        If gblngDO_PumpPower >= 0 Then .TextMatrix(1 + gblngDO_PumpPower, 1) = NAME_PUMP_POWER
        If gblngDO_SystemAlarm >= 0 Then .TextMatrix(1 + gblngDO_SystemAlarm, 1) = NAME_SYSTEM_ALARM
        If gblngDO_GasValve1 > 0 Then .TextMatrix(1 + gblngDO_GasValve1, 1) = NAME_GAS_VALVE_1
        If gblngDO_GasValve2 >= 0 Then .TextMatrix(1 + gblngDO_GasValve2, 1) = NAME_GAS_VALVE_2
        If gblngDO_GasValve3 >= 0 Then .TextMatrix(1 + gblngDO_GasValve3, 1) = NAME_GAS_VALVE_3
        If gblngDO_GasValve4 >= 0 Then .TextMatrix(1 + gblngDO_GasValve4, 1) = NAME_GAS_VALVE_4
        If gblngDO_GasValve5 >= 0 Then .TextMatrix(1 + gblngDO_GasValve4, 1) = NAME_GAS_VALVE_6
        If gblngDO_GasValve6 >= 0 Then .TextMatrix(1 + gblngDO_GasValve4, 1) = NAME_GAS_VALVE_6
        'CDA
        If gblngDO_ValveCDA >= 0 Then .TextMatrix(1 + gblngDO_ValveCDA, 1) = NAME_CDA_VALVE
        'Exhaust
        If gblngDO_Exhaust >= 0 Then .TextMatrix(1 + gblngDO_Exhaust, 1) = NAME_EXHAUST
        If gblngDO_AngleValve >= 0 Then .TextMatrix(1 + gblngDO_AngleValve, 1) = NAME_ANGLE_VALVE
        If gblngDO_ReleaseValve >= 0 Then .TextMatrix(1 + gblngDO_ReleaseValve, 1) = NAME_RELEASE_VALVE
        If gblngDO_APCGaugeValve >= 0 Then .TextMatrix(1 + gblngDO_APCGaugeValve, 1) = NAME_APC_GAUGE_VALVE
        If gblngDO_APCGaugeAngle >= 0 Then .TextMatrix(1 + gblngDO_APCGaugeAngle, 1) = NAME_APC_GAUGE_ANGLE
        If lngDO_DoorOpenValve >= 0 Then .TextMatrix(1 + lngDO_DoorOpenValve, 1) = NAME_DOOR_OPEN_VALVE
        If lngDO_DoorCloseValve >= 0 Then .TextMatrix(1 + lngDO_DoorCloseValve, 1) = NAME_DOOR_CLOSE_VALVE
        If lngDO_DoorClamp >= 0 Then .TextMatrix(1 + lngDO_DoorClamp, 1) = NAME_DOOR_CLAMP
        'REV4.1.2
        If gblngDO_AlarmRed > 0 Then .TextMatrix(1 + gblngDO_AlarmRed, 1) = NAME_ALARM_LIGHT_RED
        If gblngDO_AlarmYellow > 0 Then .TextMatrix(1 + gblngDO_AlarmYellow, 1) = NAME_ALARM_LIGHT_YELLOW
        If gblngDO_AlarmBlue > 0 Then .TextMatrix(1 + gblngDO_AlarmBlue, 1) = NAME_ALARM_LIGHT_BLUE
        If gblngDO_AlarmGreen > 0 Then .TextMatrix(1 + gblngDO_AlarmGreen, 1) = NAME_ALARM_LIGHT_GREEN
        If gblngDO_ARM_FRONT >= 0 Then .TextMatrix(1 + gblngDO_ARM_FRONT, 1) = NAME_ARM_FRONT
        If gblngDO_COVER_ARESET >= 0 Then .TextMatrix(1 + gblngDO_COVER_ARESET, 1) = NAME_COVER_ARESET
        If gblngDO_COVER_SERVO >= 0 Then .TextMatrix(1 + gblngDO_COVER_SERVO, 1) = NAME_COVER_SERVO
        If gblngDO_COVER_ORGIN >= 0 Then .TextMatrix(1 + gblngDO_COVER_ORGIN, 1) = NAME_COVER_ORGIN
        If gblngDO_COVER_MOVE >= 0 Then .TextMatrix(1 + gblngDO_COVER_MOVE, 1) = NAME_COVER_MOVE
        If gblngDO_COVER_POS_01 >= 0 Then .TextMatrix(1 + gblngDO_COVER_POS_01, 1) = NAME_COVER_POS_01
    End With
    cmbIOList(1).AddItem NAME_PC_CHECK1
    cmbIOList(1).AddItem NAME_PC_CHECK2
    cmbIOList(1).AddItem NAME_MFC_SERVO_1
    cmbIOList(1).AddItem NAME_MFC_SERVO_2
    cmbIOList(1).AddItem NAME_MFC_SERVO_3
    cmbIOList(1).AddItem NAME_MFC_SERVO_4
    cmbIOList(1).AddItem NAME_MFC_SERVO_5
    cmbIOList(1).AddItem NAME_MFC_SERVO_6
    cmbIOList(1).AddItem NAME_PUMP_POWER
    cmbIOList(1).AddItem NAME_SYSTEM_ALARM
    cmbIOList(1).AddItem NAME_GAS_VALVE_1
    cmbIOList(1).AddItem NAME_GAS_VALVE_2
    cmbIOList(1).AddItem NAME_GAS_VALVE_3
    cmbIOList(1).AddItem NAME_GAS_VALVE_4
    cmbIOList(1).AddItem NAME_GAS_VALVE_5
    cmbIOList(1).AddItem NAME_GAS_VALVE_6
    cmbIOList(1).AddItem NAME_CDA_VALVE
    cmbIOList(1).AddItem NAME_EXHAUST 'Exhaust
    cmbIOList(1).AddItem NAME_ANGLE_VALVE
    cmbIOList(1).AddItem NAME_RELEASE_VALVE
    cmbIOList(1).AddItem NAME_APC_GAUGE_VALVE
    cmbIOList(1).AddItem NAME_APC_GAUGE_ANGLE
    cmbIOList(1).AddItem NAME_DOOR_OPEN_VALVE
    cmbIOList(1).AddItem NAME_DOOR_CLOSE_VALVE
    cmbIOList(1).AddItem NAME_DOOR_CLAMP
    cmbIOList(1).AddItem NAME_ALARM_LIGHT_RED
    cmbIOList(1).AddItem NAME_ALARM_LIGHT_YELLOW
    cmbIOList(1).AddItem NAME_ALARM_LIGHT_GREEN
    cmbIOList(1).AddItem NAME_ALARM_LIGHT_BLUE
    cmbIOList(1).AddItem NAME_ARM_FRONT
    cmbIOList(1).AddItem NAME_ARM_REAR
    cmbIOList(1).AddItem NAME_COVER_ARESET
    cmbIOList(1).AddItem NAME_COVER_SERVO
    cmbIOList(1).AddItem NAME_COVER_ORGIN
    cmbIOList(1).AddItem NAME_COVER_MOVE
    cmbIOList(1).AddItem NAME_COVER_POS_01
    cmbIOList(1).AddItem "NA"
        
    With fgAI
        .Cols = 4
        .Rows = 33
        .ColWidth(0) = 800
        .ColWidth(1) = 2000
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .TextMatrix(0, 0) = "Port"
        .TextMatrix(0, 1) = "Description"
        .TextMatrix(0, 2) = "Input (V)"
        .TextMatrix(0, 3) = "Error (V)"
        For i = 1 To 32
            .RowHeight(i) = 360
            .TextMatrix(i, 0) = CStr(i - 1) & "     "
            .TextMatrix(i, 1) = "NA"
            .TextMatrix(i, 2) = "0"
            .TextMatrix(i, 3) = "0"
        Next i
        If gblngAI_CT_01 >= 0 Then .TextMatrix(1 + gblngAI_CT_01, 1) = NAME_CT_01
        If gblngAI_CT_02 >= 0 Then .TextMatrix(1 + gblngAI_CT_02, 1) = NAME_CT_02
        If gblngAI_CT_03 >= 0 Then .TextMatrix(1 + gblngAI_CT_03, 1) = NAME_CT_03
        If gblngAI_CT_04 >= 0 Then .TextMatrix(1 + gblngAI_CT_04, 1) = NAME_CT_04
        If gblngAI_CT_05 >= 0 Then .TextMatrix(1 + gblngAI_CT_05, 1) = NAME_CT_05
        If gblngAI_CT_06 >= 0 Then .TextMatrix(1 + gblngAI_CT_06, 1) = NAME_CT_06
        If gblngAI_CT_07 >= 0 Then .TextMatrix(1 + gblngAI_CT_07, 1) = NAME_CT_07
        If gblngAI_MFC_Read(0) >= 0 Then .TextMatrix(1 + gblngAI_MFC_Read(0), 1) = NAME_MFC_READ_1
        If gblngAI_MFC_Read(1) >= 0 Then .TextMatrix(1 + gblngAI_MFC_Read(1), 1) = NAME_MFC_READ_2
        If gblngAI_MFC_Read(2) >= 0 Then .TextMatrix(1 + gblngAI_MFC_Read(2), 1) = NAME_MFC_READ_3
        If gblngAI_MFC_Read(3) >= 0 Then .TextMatrix(1 + gblngAI_MFC_Read(3), 1) = NAME_MFC_READ_4
        If gblngAI_Vacuum_Gauge >= 0 Then .TextMatrix(1 + gblngAI_Vacuum_Gauge, 1) = NAME_VACUUM_GAUGE
        If gblngAI_Vacuum_Gauge2 >= 0 Then .TextMatrix(1 + gblngAI_Vacuum_Gauge2, 1) = NAME_VACUUM_GAUGE2
        If gblngAI_Oxygen_Gauge >= 0 Then .TextMatrix(1 + gblngAI_Oxygen_Gauge, 1) = NAME_OXYGEN_GAUGE
        If gblngAI_TC_Cvt1 >= 0 Then .TextMatrix(1 + gblngAI_TC_Cvt1, 1) = NAME_TC_CVT_1
        'Rev4.1.4
        If gblngAI_TCWafer1 >= 0 Then .TextMatrix(1 + gblngAI_TCWafer1, 1) = NAME_TC_WAF_1
        If gblngAI_TCWafer2 >= 0 Then .TextMatrix(1 + gblngAI_TCWafer2, 1) = NAME_TC_WAF_2
        If gblngAI_TCWafer3 >= 0 Then .TextMatrix(1 + gblngAI_TCWafer3, 1) = NAME_TC_WAF_3
        If gblngAI_TCWafer4 >= 0 Then .TextMatrix(1 + gblngAI_TCWafer4, 1) = NAME_TC_WAF_4
        If gblngAI_TCWafer5 >= 0 Then .TextMatrix(1 + gblngAI_TCWafer5, 1) = NAME_TC_WAF_5
    End With

    cmbIOList(2).AddItem NAME_MFC_READ_1
    cmbIOList(2).AddItem NAME_MFC_READ_2
    cmbIOList(2).AddItem NAME_MFC_READ_3
    cmbIOList(2).AddItem NAME_MFC_READ_4
    cmbIOList(2).AddItem NAME_MFC_READ_5
    cmbIOList(2).AddItem NAME_MFC_READ_6
    cmbIOList(2).AddItem NAME_VACUUM_GAUGE
    cmbIOList(2).AddItem NAME_VACUUM_GAUGE2
    cmbIOList(2).AddItem NAME_OXYGEN_GAUGE
    cmbIOList(2).AddItem NAME_TC_CVT_1
    'Rev4.1.4
    cmbIOList(2).AddItem NAME_TC_WAF_1
    cmbIOList(2).AddItem NAME_TC_WAF_2
    cmbIOList(2).AddItem NAME_TC_WAF_3
    cmbIOList(2).AddItem NAME_TC_WAF_4
    cmbIOList(2).AddItem NAME_TC_WAF_5
    cmbIOList(2).AddItem "NA"

    With fgAO
        .Cols = 4
        .Rows = 33
        .ColWidth(0) = 600
        .ColWidth(1) = 2000
        .ColWidth(2) = 700
        .ColWidth(3) = 700
        .TextMatrix(0, 0) = "Port"
        .TextMatrix(0, 1) = "Description"
        .TextMatrix(0, 2) = "MaxV"
        .TextMatrix(0, 3) = "OutV"
        
        For i = 1 To 32
            .RowHeight(i) = 360
            .TextMatrix(i, 0) = CStr(i - 1) & "     "
            .TextMatrix(i, 1) = "NA"
            .TextMatrix(i, 2) = "0"
            .TextMatrix(i, 3) = "0"
        Next i
        If gblngAO_SCR_TBC >= 0 Then .TextMatrix(1 + gblngAO_SCR_TBC, 1) = NAME_SET_SCR_TBC
        If gblngAO_SCR_TR >= 0 Then .TextMatrix(1 + gblngAO_SCR_TR, 1) = NAME_SET_SCR_TR
        If gblngAO_SCR_TL >= 0 Then .TextMatrix(1 + gblngAO_SCR_TL, 1) = NAME_SET_SCR_TL
        If gblngAO_SCR_BF >= 0 Then .TextMatrix(1 + gblngAO_SCR_BF, 1) = NAME_SET_SCR_BF
        If gblngAO_SCR_BR >= 0 Then .TextMatrix(1 + gblngAO_SCR_BR, 1) = NAME_SET_SCR_BR
        If gblngAO_SCR_6 >= 0 Then .TextMatrix(1 + gblngAO_SCR_6, 1) = NAME_SET_SCR_6
        If gblngAO_SCR_7 >= 0 Then .TextMatrix(1 + gblngAO_SCR_7, 1) = NAME_SET_SCR_7
        If gblngAO_SCR_8 >= 0 Then .TextMatrix(1 + gblngAO_SCR_8, 1) = NAME_SET_SCR_8
        If gblngAO_SCR_9 >= 0 Then .TextMatrix(1 + gblngAO_SCR_9, 1) = NAME_SET_SCR_9
        If gblngAO_SCR_10 >= 0 Then .TextMatrix(1 + gblngAO_SCR_10, 1) = NAME_SET_SCR_10
        If gblngAO_SCR_11 >= 0 Then .TextMatrix(1 + gblngAO_SCR_11, 1) = NAME_SET_SCR_11
        If gblngAO_SCR_12 >= 0 Then .TextMatrix(1 + gblngAO_SCR_12, 1) = NAME_SET_SCR_12
        If gblngAO_SCR_13 >= 0 Then .TextMatrix(1 + gblngAO_SCR_13, 1) = NAME_SET_SCR_13
        If gblngAO_SCR_14 >= 0 Then .TextMatrix(1 + gblngAO_SCR_14, 1) = NAME_SET_SCR_14
        If gblngAO_SCR_15 >= 0 Then .TextMatrix(1 + gblngAO_SCR_14, 1) = NAME_SET_SCR_15
        If gblngAO_SCR_16 >= 0 Then .TextMatrix(1 + gblngAO_SCR_14, 1) = NAME_SET_SCR_16
        If gblngAO_SCR_17 >= 0 Then .TextMatrix(1 + gblngAO_SCR_14, 1) = NAME_SET_SCR_17
        
        If gblngAO_MFC1 >= 0 Then .TextMatrix(1 + gblngAO_MFC1, 1) = NAME_MFC_SET_1
        If gblngAO_MFC2 >= 0 Then .TextMatrix(1 + gblngAO_MFC2, 1) = NAME_MFC_SET_2
        If gblngAO_MFC3 >= 0 Then .TextMatrix(1 + gblngAO_MFC3, 1) = NAME_MFC_SET_3
        If gblngAO_MFC4 >= 0 Then .TextMatrix(1 + gblngAO_MFC4, 1) = NAME_MFC_SET_4
        If gblngAO_MFC5 >= 0 Then .TextMatrix(1 + gblngAO_MFC5, 1) = NAME_MFC_SET_5
        If gblngAO_MFC6 >= 0 Then .TextMatrix(1 + gblngAO_MFC6, 1) = NAME_MFC_SET_6
    End With
    cmbIOList(3).AddItem NAME_SET_SCR_TBC
    cmbIOList(3).AddItem NAME_SET_SCR_TR
    cmbIOList(3).AddItem NAME_SET_SCR_TL
    cmbIOList(3).AddItem NAME_SET_SCR_BF
    cmbIOList(3).AddItem NAME_SET_SCR_BR
    cmbIOList(3).AddItem NAME_MFC_SET_1
    cmbIOList(3).AddItem NAME_MFC_SET_2
    cmbIOList(3).AddItem NAME_MFC_SET_3
    cmbIOList(3).AddItem NAME_MFC_SET_4
    cmbIOList(3).AddItem NAME_MFC_SET_5
    cmbIOList(3).AddItem NAME_MFC_SET_6
    cmbIOList(3).AddItem NAME_SET_SCR_6
    cmbIOList(3).AddItem NAME_SET_SCR_7
    cmbIOList(3).AddItem NAME_SET_SCR_8
    cmbIOList(3).AddItem NAME_SET_SCR_9
    cmbIOList(3).AddItem NAME_SET_SCR_10
    cmbIOList(3).AddItem NAME_SET_SCR_11
    cmbIOList(3).AddItem NAME_SET_SCR_12
    cmbIOList(3).AddItem NAME_SET_SCR_13
    cmbIOList(3).AddItem NAME_SET_SCR_14
    cmbIOList(3).AddItem NAME_SET_SCR_15
    cmbIOList(3).AddItem NAME_SET_SCR_16
    cmbIOList(3).AddItem NAME_SET_SCR_17
    cmbIOList(3).AddItem "NA"
    
    'Rta9.0.0.0 Add the advantech pci-1719hgu ai
    With fgAdvDaqAI
        .Cols = 11
        .Rows = 25
        .ColWidth(0) = 800
        .ColWidth(1) = 2000
        .ColWidth(2) = 1250
        .ColWidth(3) = 1250
        .ColWidth(4) = 1250
        .ColWidth(5) = 1250
        .ColWidth(6) = 1250
        .ColWidth(7) = 1250
        .ColWidth(8) = 1250
        .ColWidth(9) = 1250
        .ColWidth(10) = 2050
        
        .TextMatrix(0, 0) = "Port"
        .TextMatrix(0, 1) = "Name"
        .TextMatrix(0, 2) = "X^2"
        .TextMatrix(0, 3) = "Error"
        .TextMatrix(0, 4) = "Ratio"
        .TextMatrix(0, 5) = "Orig."
        .TextMatrix(0, 6) = "X^3"
        .TextMatrix(0, 7) = "X^4"
        .TextMatrix(0, 8) = "X^5"
        .TextMatrix(0, 9) = "Loop No"
        .TextMatrix(0, 10) = "Precision Digit"
        For i = 1 To 24
            .RowHeight(i) = 360
            .TextMatrix(i, 0) = CStr(i - 1) & "     "
            .TextMatrix(i, 1) = "NA"
        Next i
        If gblngAI_TC_01 >= 0 Then .TextMatrix(1 + gblngAI_TC_01, 1) = NAME_TC_01
        If gblngAI_TC_02 >= 0 Then .TextMatrix(1 + gblngAI_TC_02, 1) = NAME_TC_02
        If gblngAI_TC_03 >= 0 Then .TextMatrix(1 + gblngAI_TC_03, 1) = NAME_TC_03
        If gblngAI_TC_04 >= 0 Then .TextMatrix(1 + gblngAI_TC_04, 1) = NAME_TC_04
        If gblngAI_TC_05 >= 0 Then .TextMatrix(1 + gblngAI_TC_05, 1) = NAME_TC_05
        If gblngAI_TC_06 >= 0 Then .TextMatrix(1 + gblngAI_TC_06, 1) = NAME_TC_06
        If gblngAI_TC_07 >= 0 Then .TextMatrix(1 + gblngAI_TC_07, 1) = NAME_TC_07
        If gblngAI_TC_08 >= 0 Then .TextMatrix(1 + gblngAI_TC_08, 1) = NAME_TC_08
        If gblngAI_TC_09 >= 0 Then .TextMatrix(1 + gblngAI_TC_09, 1) = NAME_TC_09
        If gblngAI_TC_10 >= 0 Then .TextMatrix(1 + gblngAI_TC_10, 1) = NAME_TC_10
        If gblngAI_TC_11 >= 0 Then .TextMatrix(1 + gblngAI_TC_11, 1) = NAME_TC_11
        If gblngAI_TC_12 >= 0 Then .TextMatrix(1 + gblngAI_TC_12, 1) = NAME_TC_12
        If gblngAI_TC_13 >= 0 Then .TextMatrix(1 + gblngAI_TC_13, 1) = NAME_TC_13
        If gblngAI_TC_14 >= 0 Then .TextMatrix(1 + gblngAI_TC_14, 1) = NAME_TC_14
        If gblngAI_TC_15 >= 0 Then .TextMatrix(1 + gblngAI_TC_15, 1) = NAME_TC_15
        If gblngAI_TC_16 >= 0 Then .TextMatrix(1 + gblngAI_TC_16, 1) = NAME_TC_16
        
    End With
    cmbIOList(4).AddItem NAME_TC_01
    cmbIOList(4).AddItem NAME_TC_02
    cmbIOList(4).AddItem NAME_TC_03
    cmbIOList(4).AddItem NAME_TC_04
    cmbIOList(4).AddItem NAME_TC_05
    cmbIOList(4).AddItem NAME_TC_06
    cmbIOList(4).AddItem NAME_TC_07
    cmbIOList(4).AddItem NAME_TC_08
    cmbIOList(4).AddItem NAME_TC_09
    cmbIOList(4).AddItem NAME_TC_10
    cmbIOList(4).AddItem NAME_TC_11
    cmbIOList(4).AddItem NAME_TC_12
    cmbIOList(4).AddItem NAME_TC_13
    cmbIOList(4).AddItem NAME_TC_14
    cmbIOList(4).AddItem NAME_TC_15
    cmbIOList(4).AddItem NAME_TC_16
    cmbIOList(4).AddItem NAME_TC_17
    cmbIOList(4).AddItem NAME_TC_18
    cmbIOList(4).AddItem NAME_TC_19
    cmbIOList(4).AddItem NAME_TC_20
    cmbIOList(4).AddItem NAME_TC_21
    cmbIOList(4).AddItem NAME_TC_22
    cmbIOList(4).AddItem NAME_TC_23
    cmbIOList(4).AddItem NAME_TC_24
    cmbIOList(4).AddItem NAME_O2_01
    cmbIOList(4).AddItem NAME_PS_01
    cmbIOList(4).AddItem NAME_P
    For i = 0 To cmbIOList.UBound
        cmbIOList(i).Visible = False
    Next i
        
    frmConfiguration.tabConfiguration.TabVisible(1) = False
    frmConfiguration.tabConfiguration.TabVisible(2) = False
    frmConfiguration.tabConfiguration.TabVisible(3) = False
    frmConfiguration.tabConfiguration.TabVisible(4) = False
    frmConfiguration.tabConfiguration.TabVisible(5) = False
    frmConfiguration.tabConfiguration.TabVisible(6) = False
        
    'Ver 4.1.2 Modify the control loop mode
    cmbControlMode.Clear
    cmbControlMode.AddItem "Single"
    cmbControlMode.AddItem "Multiple"
    cmbControlMode.AddItem "Uniformity"
    cmbControlMode.ListIndex = 0
    
    
    ' Add selectable items of Thermocouple type
    cmbTCType.AddItem "J type"            ' 0
    cmbTCType.AddItem "K type"            ' 1
    cmbTCType.AddItem "S type"            ' 2
    cmbTCType.AddItem "T type"            ' 3
    cmbTCType.AddItem "B type"            ' 4
    cmbTCType.AddItem "R type"            ' 5
    cmbTCType.AddItem "E type"            ' 6
    For i = 0 To 11
        cmbTCVoltageRange.AddItem CStr(i)
    Next i
    
    With fgAlarm
        .Cols = 3
        .Rows = 101
        .ColWidth(0) = 1500
        .ColWidth(1) = 9000
        .ColWidth(2) = 1500
                
        .TextMatrix(0, 0) = "編號"
        .TextMatrix(0, 1) = "警報訊息"
        .TextMatrix(0, 2) = "處置方式"
        .TextMatrix(1, 0) = "4001"
        .TextMatrix(1, 1) = Alarm4001

    End With
    
    Para.AlarmName(1) = "系統異常,請連絡原廠"
    Para.AlarmName(2) = "金屬腔體過熱"
    Para.AlarmName(3) = "控溫超過高溫限制"
    Para.AlarmName(4) = "能量輸出超過設定限制"
    Para.AlarmName(5) = "腔體壓力處於真空(負壓)狀態"
    Para.AlarmName(6) = "腔體壓力仍未達到真空標準值"
    Para.AlarmName(7) = "腔體壓力正壓超過上限"
    Para.AlarmName(8) = "溫度過低"
    Para.AlarmName(9) = "溫度過高"
    Para.AlarmName(10) = "機台未備妥"
    Para.AlarmName(11) = "腔門狀態異常"
    Para.AlarmName(12) = "CDA異常"
    Para.AlarmName(13) = "冷卻水異常"
    Para.AlarmName(14) = "燈管連線檢查異常"
    Para.AlarmName(15) = "燈管壽命超過設定限制"
    Para.AlarmName(16) = "監控溫度超過設定限制"
    Para.AlarmName(17) = "伺服器連線異常"
    Para.AlarmName(18) = "EMO 急停按鍵作用中"
    Para.AlarmName(19) = "MFC 流量異常"
    Para.AlarmName(20) = "Power Meter 連線異常"
    Para.AlarmName(21) = "自動機連線異常"
    Para.AlarmName(22) = "燈管電流異常(大於上限)"
    Para.AlarmName(23) = "燈管電流異常(小於下限)"
    Para.AlarmName(24) = "氧氣流量異常"
    Para.AlarmName(25) = "伺服端異常警告(Abnormal)"
    Para.AlarmName(26) = "伺服端循環停止(Stop)"
    Para.AlarmName(27) = "伺服端強制停止(Abort)"
    Para.AlarmName(28) = "電動缸伺服異常"
    Para.AlarmName(29) = "MTC空接"
    Para.AlarmName(30) = "R值不在指定範圍內"
    Para.AlarmName(31) = "真空計故障"
    Para.AlarmName(32) = "SCR連線異常"
    Para.AlarmName(33) = "泵異常/DI"
    With fgTeach
        .Rows = 51
        .Cols = 5
        .ColWidth(0) = 800
        .ColWidth(1) = 1200
        .ColWidth(2) = 1200
        .ColWidth(3) = 1200
        .ColWidth(4) = 1200
        .TextMatrix(0, 1) = "A1"
        .TextMatrix(0, 2) = "A2"
        .TextMatrix(0, 3) = "A3"
        .TextMatrix(0, 4) = "A4"
        For i = 0 To 49
            .TextMatrix(i + 1, 0) = CInt(i) & "     "
        Next i
    End With
    
End Sub

Public Sub ReloadIO()
    Dim i As Integer
    
    For i = 1 To 32
        If fgDO.TextMatrix(i, 1) = NAME_PC_CHECK1 Then
            gblngDO_PC_Check1 = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_PC_CHECK2 Then
            gblngDO_PC_Check2 = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_MFC_SERVO_1 Then
            gblngDO_MFC_ValveServo1 = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_MFC_SERVO_2 Then
            gblngDO_MFC_ValveServo2 = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_MFC_SERVO_3 Then
            gblngDO_MFC_ValveServo3 = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_MFC_SERVO_4 Then
            gblngDO_MFC_ValveServo4 = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_MFC_SERVO_5 Then
            gblngDO_MFC_ValveServo5 = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_MFC_SERVO_6 Then
            gblngDO_MFC_ValveServo6 = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_CDA_VALVE Then
            gblngDO_ValveCDA = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_EXHAUST Then
            gblngDO_Exhaust = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_GAS_VALVE_1 Then
            gblngDO_GasValve1 = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_GAS_VALVE_2 Then
            gblngDO_GasValve2 = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_GAS_VALVE_3 Then
            gblngDO_GasValve3 = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_GAS_VALVE_4 Then
            gblngDO_GasValve4 = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_GAS_VALVE_5 Then
            gblngDO_GasValve5 = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_GAS_VALVE_6 Then
            gblngDO_GasValve6 = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_ANGLE_VALVE Then
            gblngDO_AngleValve = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_RELEASE_VALVE Then
            gblngDO_ReleaseValve = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_PUMP_POWER Then
            gblngDO_PumpPower = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_APC_GAUGE_VALVE Then
            gblngDO_APCGaugeValve = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_APC_GAUGE_ANGLE Then
            gblngDO_APCGaugeAngle = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_DOOR_OPEN_VALVE Then
            lngDO_DoorOpenValve = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_DOOR_CLOSE_VALVE Then
            lngDO_DoorCloseValve = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_DOOR_CLAMP Then
            lngDO_DoorClamp = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_SYSTEM_ALARM Then
            gblngDO_SystemAlarm = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_ALARM_LIGHT_RED Then
            gblngDO_AlarmRed = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_ALARM_LIGHT_YELLOW Then
            gblngDO_AlarmYellow = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_ALARM_LIGHT_GREEN Then
            gblngDO_AlarmGreen = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_ALARM_LIGHT_BLUE Then
            gblngDO_AlarmBlue = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_ARM_FRONT Then
            gblngDO_ARM_FRONT = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_COVER_ARESET Then
            gblngDO_COVER_ARESET = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_COVER_SERVO Then
            gblngDO_COVER_SERVO = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_COVER_ORGIN Then
            gblngDO_COVER_ORGIN = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_COVER_MOVE Then
            gblngDO_COVER_MOVE = i - 1
        End If
        If fgDO.TextMatrix(i, 1) = NAME_COVER_POS_01 Then
            gblngDO_COVER_POS_01 = i - 1
        End If

    Next i
    
    For i = 1 To 33
        If fgDI.TextMatrix(i, 1) = NAME_CHAMBER_OVERHEAT Then
            gblngDI_ChamberOverheat = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_AIR_ALARM Then
            gblngDI_AirAlarm = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_WATER_ALARM Then
            gblngDI_WaterAlarm = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_EMS_ALARM Then
            gblngDI_EMS_Alarm = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_MC_STATUS Then
            gblngDI_MC_Input = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_DOOR_OPEN Then
            gblngDI_DoorOpen = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_DOOR_CLOSE Then
            gblngDI_DoorClose = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_DOOR_OPEN_SENSOR Then
            gblngDI_DoorOpenSensor = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_DOOR_CLOSE_SENSOR Then
            gblngDI_DoorCloseSensor = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_DI_DOOR_CLAMP Then
            gblngDI_DoorClamp = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_SYSTEM_ALARM Then
            gblngDI_SystemAlarm = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_VAC_SENSOR_R Then
            gblngDI_VacuumGateR = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_VAC_SENSOR_L Then
            gblngDI_VacuumGateL = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_CHAMBER_SENSOR_R Then
            gblngDI_ChamberGateR = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_CHAMBER_SENSOR_L Then
            gblngDI_ChamberGateL = i - 1
        End If
        
        If fgDI.TextMatrix(i, 1) = NAME_CT1 Then
            gblngDI_CT1 = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_CT2 Then
            gblngDI_CT2 = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_CT3 Then
            gblngDI_CT3 = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_CT4 Then
            gblngDI_CT4 = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_CT5 Then
            gblngDI_CT5 = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_CT6 Then
            gblngDI_CT6 = i - 1
        End If
        'Cover
        If fgDI.TextMatrix(i, 1) = NAME_COVER_ALARM1 Then
            gblngDI_CoverAlarm1 = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_COVER_SERVO_RDY1 Then
            gblngDI_CoverServoRdy1 = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_COVER_ORIG_RDY1 Then
            gblngDI_CoverOrigRdy1 = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_COVER_ISMOVING1 Then
            gblngDI_CoverIsMoving1 = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_COVER_UP_INP1 Then
            gblngDI_CoverUpInpos1 = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_COVER_DOWN_INP1 Then
            gblngDI_CoverDownInpos1 = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_COVER_ALARM2 Then
            gblngDI_CoverAlarm2 = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_COVER_SERVO_RDY2 Then
            gblngDI_CoverServoRdy2 = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_COVER_ORIG_RDY2 Then
            gblngDI_CoverOrigRdy2 = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_COVER_ISMOVING2 Then
            gblngDI_CoverIsMoving2 = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_COVER_UP_INP2 Then
            gblngDI_CoverUpInpos2 = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_COVER_DOWN_INP2 Then
            gblngDI_CoverDownInpos2 = i - 1
        End If
        If fgDI.TextMatrix(i, 1) = NAME_PUMP_ALARM Then
            gblngDI_PumpAlarm = i - 1
        End If
    Next i
    
    For i = 1 To 16
        If fgAI.TextMatrix(i, 1) = NAME_CT_01 Then
            gblngAI_CT_01 = i - 1
        End If
        If fgAI.TextMatrix(i, 1) = NAME_CT_02 Then
            gblngAI_CT_02 = i - 1
        End If
        If fgAI.TextMatrix(i, 1) = NAME_CT_03 Then
            gblngAI_CT_03 = i - 1
        End If
        If fgAI.TextMatrix(i, 1) = NAME_CT_04 Then
            gblngAI_CT_04 = i - 1
        End If
        If fgAI.TextMatrix(i, 1) = NAME_CT_05 Then
            gblngAI_CT_05 = i - 1
        End If
        If fgAI.TextMatrix(i, 1) = NAME_CT_06 Then
            gblngAI_CT_06 = i - 1
        End If
        If fgAI.TextMatrix(i, 1) = NAME_CT_07 Then
            gblngAI_CT_07 = i - 1
        End If
        If fgAI.TextMatrix(i, 1) = NAME_MFC_READ_1 Then
            gblngAI_MFC_Read(0) = i - 1
        End If
        If fgAI.TextMatrix(i, 1) = NAME_MFC_READ_2 Then
            gblngAI_MFC_Read(1) = i - 1
        End If
        If fgAI.TextMatrix(i, 1) = NAME_MFC_READ_3 Then
            gblngAI_MFC_Read(2) = i - 1
        End If
        If fgAI.TextMatrix(i, 1) = NAME_MFC_READ_4 Then
            gblngAI_MFC_Read(3) = i - 1
        End If
        If fgAI.TextMatrix(i, 1) = NAME_MFC_READ_5 Then
            gblngAI_MFC_Read(4) = i - 1
        End If
        If fgAI.TextMatrix(i, 1) = NAME_MFC_READ_6 Then
            gblngAI_MFC_Read(5) = i - 1
        End If
        If fgAI.TextMatrix(i, 1) = NAME_VACUUM_GAUGE Then
            gblngAI_Vacuum_Gauge = i - 1
        End If
        If fgAI.TextMatrix(i, 1) = NAME_VACUUM_GAUGE2 Then
            gblngAI_Vacuum_Gauge2 = i - 1
        End If
        If fgAI.TextMatrix(i, 1) = NAME_OXYGEN_GAUGE Then
            gblngAI_Oxygen_Gauge = i - 1
        End If
        If fgAI.TextMatrix(i, 1) = NAME_TC_CVT_1 Then
            gblngAI_TC_Cvt1 = i - 1
        End If
        'Rev4.1.4
        If fgAI.TextMatrix(i, 1) = NAME_TC_WAF_1 Then
            gblngAI_TCWafer1 = i - 1
        End If
        If fgAI.TextMatrix(i, 1) = NAME_TC_WAF_2 Then
            gblngAI_TCWafer2 = i - 1
        End If
        If fgAI.TextMatrix(i, 1) = NAME_TC_WAF_3 Then
            gblngAI_TCWafer3 = i - 1
        End If
        If fgAI.TextMatrix(i, 1) = NAME_TC_WAF_4 Then
            gblngAI_TCWafer4 = i - 1
        End If
        If fgAI.TextMatrix(i, 1) = NAME_TC_WAF_5 Then
            gblngAI_TCWafer5 = i - 1
        End If
    Next i
    
    For i = 1 To 32
        If fgAO.TextMatrix(i, 1) = NAME_SET_SCR_TBC Then
            gblngAO_SCR_TBC = i - 1
        End If
        If fgAO.TextMatrix(i, 1) = NAME_SET_SCR_TR Then
            gblngAO_SCR_TR = i - 1
        End If
        If fgAO.TextMatrix(i, 1) = NAME_SET_SCR_TL Then
            gblngAO_SCR_TL = i - 1
        End If
        If fgAO.TextMatrix(i, 1) = NAME_SET_SCR_BF Then
            gblngAO_SCR_BF = i - 1
        End If
        If fgAO.TextMatrix(i, 1) = NAME_SET_SCR_BR Then
            gblngAO_SCR_BR = i - 1
        End If
        '120713 Josh
        If fgAO.TextMatrix(i, 1) = NAME_SET_SCR_6 Then
            gblngAO_SCR_6 = i - 1
        End If
        If fgAO.TextMatrix(i, 1) = NAME_SET_SCR_7 Then
            gblngAO_SCR_7 = i - 1
        End If
        If fgAO.TextMatrix(i, 1) = NAME_SET_SCR_8 Then
            gblngAO_SCR_8 = i - 1
        End If
        If fgAO.TextMatrix(i, 1) = NAME_SET_SCR_9 Then
            gblngAO_SCR_9 = i - 1
        End If
        If fgAO.TextMatrix(i, 1) = NAME_SET_SCR_10 Then
            gblngAO_SCR_10 = i - 1
        End If
        If fgAO.TextMatrix(i, 1) = NAME_SET_SCR_11 Then
            gblngAO_SCR_11 = i - 1
        End If
        If fgAO.TextMatrix(i, 1) = NAME_SET_SCR_12 Then
            gblngAO_SCR_12 = i - 1
        End If
        If fgAO.TextMatrix(i, 1) = NAME_SET_SCR_13 Then
            gblngAO_SCR_13 = i - 1
        End If
        If fgAO.TextMatrix(i, 1) = NAME_SET_SCR_14 Then
            gblngAO_SCR_14 = i - 1
        End If
        If fgAO.TextMatrix(i, 1) = NAME_SET_SCR_15 Then
            gblngAO_SCR_15 = i - 1
        End If
        If fgAO.TextMatrix(i, 1) = NAME_SET_SCR_16 Then
            gblngAO_SCR_16 = i - 1
        End If
        If fgAO.TextMatrix(i, 1) = NAME_SET_SCR_17 Then
            gblngAO_SCR_17 = i - 1
        End If
        If fgAO.TextMatrix(i, 1) = NAME_MFC_SET_1 Then
            gblngAO_MFC1 = i - 1
        End If
        If fgAO.TextMatrix(i, 1) = NAME_MFC_SET_2 Then
            gblngAO_MFC2 = i - 1
        End If
        If fgAO.TextMatrix(i, 1) = NAME_MFC_SET_3 Then
            gblngAO_MFC3 = i - 1
        End If
        If fgAO.TextMatrix(i, 1) = NAME_MFC_SET_4 Then
            gblngAO_MFC4 = i - 1
        End If
        If fgAO.TextMatrix(i, 1) = NAME_MFC_SET_5 Then
            gblngAO_MFC5 = i - 1
        End If
        If fgAO.TextMatrix(i, 1) = NAME_MFC_SET_6 Then
            gblngAO_MFC6 = i - 1
        End If
        
    Next i
            
    gblngCheckPC = gblngDO_PC_Check1
    
    'RTA9.0.0.0
    For i = 1 To 24
        If fgAdvDaqAI.TextMatrix(i, 1) = NAME_TC_01 Then
            gblngAI_TC_01 = i - 1
        End If
        If fgAdvDaqAI.TextMatrix(i, 1) = NAME_TC_02 Then
            gblngAI_TC_02 = i - 1
        End If
        If fgAdvDaqAI.TextMatrix(i, 1) = NAME_TC_03 Then
            gblngAI_TC_03 = i - 1
        End If
        If fgAdvDaqAI.TextMatrix(i, 1) = NAME_TC_04 Then
            gblngAI_TC_04 = i - 1
        End If
        If fgAdvDaqAI.TextMatrix(i, 1) = NAME_TC_05 Then
            gblngAI_TC_05 = i - 1
        End If
        If fgAdvDaqAI.TextMatrix(i, 1) = NAME_TC_06 Then
            gblngAI_TC_06 = i - 1
        End If
        If fgAdvDaqAI.TextMatrix(i, 1) = NAME_TC_07 Then
            gblngAI_TC_07 = i - 1
        End If
        If fgAdvDaqAI.TextMatrix(i, 1) = NAME_TC_08 Then
            gblngAI_TC_08 = i - 1
        End If
        If fgAdvDaqAI.TextMatrix(i, 1) = NAME_TC_09 Then
            gblngAI_TC_09 = i - 1
        End If
        If fgAdvDaqAI.TextMatrix(i, 1) = NAME_TC_10 Then
            gblngAI_TC_10 = i - 1
        End If
        If fgAdvDaqAI.TextMatrix(i, 1) = NAME_TC_11 Then
            gblngAI_TC_11 = i - 1
        End If
        If fgAdvDaqAI.TextMatrix(i, 1) = NAME_TC_12 Then
            gblngAI_TC_12 = i - 1
        End If
        If fgAdvDaqAI.TextMatrix(i, 1) = NAME_TC_13 Then
            gblngAI_TC_13 = i - 1
        End If
        If fgAdvDaqAI.TextMatrix(i, 1) = NAME_TC_14 Then
            gblngAI_TC_14 = i - 1
        End If
        If fgAdvDaqAI.TextMatrix(i, 1) = NAME_TC_15 Then
            gblngAI_TC_15 = i - 1
        End If
        If fgAdvDaqAI.TextMatrix(i, 1) = NAME_TC_16 Then
            gblngAI_TC_16 = i - 1
        End If
        If fgAdvDaqAI.TextMatrix(i, 1) = NAME_TC_17 Then
            gblngAI_TC_17 = i - 1
        End If
        If fgAdvDaqAI.TextMatrix(i, 1) = NAME_TC_18 Then
            gblngAI_TC_18 = i - 1
        End If
        If fgAdvDaqAI.TextMatrix(i, 1) = NAME_TC_19 Then
            gblngAI_TC_19 = i - 1
        End If
        If fgAdvDaqAI.TextMatrix(i, 1) = NAME_TC_20 Then
            gblngAI_TC_20 = i - 1
        End If
        If fgAdvDaqAI.TextMatrix(i, 1) = NAME_TC_21 Then
            gblngAI_TC_21 = i - 1
        End If
        If fgAdvDaqAI.TextMatrix(i, 1) = NAME_TC_22 Then
            gblngAI_TC_22 = i - 1
        End If
        If fgAdvDaqAI.TextMatrix(i, 1) = NAME_TC_23 Then
            gblngAI_TC_23 = i - 1
        End If
        If fgAdvDaqAI.TextMatrix(i, 1) = NAME_TC_24 Then
            gblngAI_TC_24 = i - 1
        End If
    Next i
    
End Sub

Public Sub ParameterSave()
    Dim i               As Integer
    Dim j               As Integer
    Dim lngRet          As Long
    Dim StrFileName     As String
    Dim iInputDevice    As Integer
    Dim ExtendPath As String
    NewGasCount = 0
    
    On Error GoTo ERR_PARAMETER_SAVE
    
    StrFileName = gbSystemPath & "\System\system.cfg"
    ExtendPath = App.Path + ProcDict_Path

    If txtParaNormal(0).text = "" Then txtParaNormal(0).text = "1500"
    If txtParaNormal(1).text = "" Then txtParaNormal(1).text = "90"
    If txtParaNormal(2).text = "" Then txtParaNormal(2).text = "0"
    If txtParaNormal(3).text = "" Then txtParaNormal(3).text = "500"
    If txtParaNormal(4).text = "" Then txtParaNormal(4).text = "1"
    If txtParaNormal(5).text = "" Then txtParaNormal(5).text = "5"
    If txtParaNormal(6).text = "" Then txtParaNormal(6).text = "10"
    If txtParaNormal(7).text = "" Then txtParaNormal(7).text = "100"
    If txtParaNormal(8).text = "" Then txtParaNormal(8).text = "100"
    If txtParaNormal(9).text = "" Then txtParaNormal(9).text = "0"
    If txtParaNormal(10).text = "" Then txtParaNormal(10).text = "0"
    If txtParaNormal(14).text = "" Then txtParaNormal(14).text = "0"
    If txtParaNormal(15).text = "" Then txtParaNormal(15).text = "0"
    If txtParaNormal(16).text = "" Then txtParaNormal(16).text = "0"
    If txtParaNormal(17).text = "" Then txtParaNormal(17).text = "0"
    If txtParaNormal(18).text = "" Then txtParaNormal(18).text = "0"
    If txtParaNormal(19).text = "" Then txtParaNormal(19).text = "0"
    
    If txtParaHeat(0).text = "" Then txtParaHeat(0).text = "90"
    If txtParaHeat(1).text = "" Then txtParaHeat(1).text = "600"
    If txtParaHeat(2).text = "" Then txtParaHeat(2).text = "1"
    If txtParaHeat(3).text = "" Then txtParaHeat(3).text = "-18"
    If txtParaHeat(4).text = "" Then txtParaHeat(4).text = "60000"
    If txtParaHeat(5).text = "" Then txtParaHeat(5).text = "30"
    If txtParaHeat(6).text = "" Then txtParaHeat(6).text = "1"
    If txtParaHeat(7).text = "" Then txtParaHeat(7).text = "-200"
    If txtParaHeat(8).text = "" Then txtParaHeat(8).text = "5"

    If txtParaVacuum(0).text = "" Then txtParaVacuum(0).text = "0"
    If txtParaVacuum(1).text = "" Then txtParaVacuum(1).text = "1000"
    If txtParaVacuum(2).text = "" Then txtParaVacuum(2).text = "3000"
    If txtParaVacuum(3).text = "" Then txtParaVacuum(3).text = "200"
    If txtParaVacuum(4).text = "" Then txtParaVacuum(4).text = "0"
    If txtParaVacuum(5).text = "" Then txtParaVacuum(5).text = "2"
    If txtParaVacuum(6).text = "" Then txtParaVacuum(6).text = "0"
    If txtParaVacuum(7).text = "" Then txtParaVacuum(7).text = "0"
    If txtParaVacuum(8).text = "" Then txtParaVacuum(8).text = "0"
    If txtParaVacuum(9).text = "" Then txtParaVacuum(9).text = "0"
    If txtParaVacuum(10).text = "" Then txtParaVacuum(10).text = "0"
    If txtParaVacuum(11).text = "" Then txtParaVacuum(11).text = "1"
    If txtParaVacuum(12).text = "" Then txtParaVacuum(12).text = "1"
    
    For i = 0 To txtParaGasAlias.UBound
        If txtParaGasAlias(i).text = "" Then txtParaGasAlias(i).text = "NA"
    Next i
    For i = 0 To txtParaGasUnit.UBound
        If txtParaGasUnit(i).text = "" Then txtParaGasUnit(i).text = "SLPM"
    Next i
    
    For i = 0 To txtParaGasValue.UBound
        If txtParaGasValue(i).text = "" Then txtParaGasValue(i).text = "1"
    Next i
    For i = 0 To txtParaGasBias.UBound
        If txtParaGasBias(i).text = "" Then txtParaGasBias(i).text = "0"
    Next i
       
    If txtIntensityWeight(0).text = "" Then txtIntensityWeight(0).text = "100"
    If txtIntensityWeight(1).text = "" Then txtIntensityWeight(1).text = "100"
    If txtIntensityWeight(2).text = "" Then txtIntensityWeight(2).text = "100"
    If txtIntensityWeight(3).text = "" Then txtIntensityWeight(3).text = "100"
    If txtIntensityWeight(4).text = "" Then txtIntensityWeight(4).text = "100"
    
    If txtIntensityWeightS(0).text = "" Then txtIntensityWeightS(0).text = "100"
    If txtIntensityWeightS(1).text = "" Then txtIntensityWeightS(1).text = "100"
    If txtIntensityWeightS(2).text = "" Then txtIntensityWeightS(2).text = "100"
    If txtIntensityWeightS(3).text = "" Then txtIntensityWeightS(3).text = "100"
    If txtIntensityWeightS(4).text = "" Then txtIntensityWeightS(4).text = "100"
    
    
    For i = 0 To txtIntensityWeight.UBound
        lngRet = WritePrivateProfileString("PARAMETER", "IntensityWeight" & CStr(i + 1), txtIntensityWeight(i).text, StrFileName)
        gbsngIntensityWeight(i) = CSng(txtIntensityWeight(i).text) / 100
    Next i
    
    For i = 0 To txtIntensityWeightS.UBound
        lngRet = WritePrivateProfileString("PARAMETER", "IntensityWeightS" & CStr(i + 1), txtIntensityWeightS(i).text, StrFileName)
        gbsngIntensityWeightS(i) = CSng(txtIntensityWeightS(i).text) / 100
    Next i
   
    If txtPropertyCoeff(0).text = "" Then txtPropertyCoeff(0).text = "1"
    If txtPropertyCoeff(1).text = "" Then txtPropertyCoeff(1).text = "1"
    If txtPropertyCoeff(2).text = "" Then txtPropertyCoeff(2).text = "1"
    If txtPropertyCoeff(3).text = "" Then txtPropertyCoeff(3).text = "1"
    If txtPropertyCoeff(4).text = "" Then txtPropertyCoeff(4).text = "1"
        
    'Rev4.1.5
    If txtSubWeight1.text = "" Then txtSubWeight1.text = "0"
    If txtSubWeight2.text = "" Then txtSubWeight2.text = "0"
    
    'The property of PID coefficient
    For i = 0 To txtPropertyCoeff.UBound
        If i = 4 Then
'            lngRet = WritePrivateProfileString("PARAMETER", "PropertyCoefficient" & CStr(i + 1), EncryptDecrypt(txtPropertyCoeff(i).text, 123), StrFileName)
                
            lngRet = WritePrivateProfileString("PARAMETER", "PropertyCoefficient" & CStr(i + 1), txtPropertyCoeff(i).text, StrFileName)
        Else
            lngRet = WritePrivateProfileString("PARAMETER", "PropertyCoefficient" & CStr(i + 1), txtPropertyCoeff(i).text, StrFileName)
        End If
        
        gbsngPropertyCoefficient(i) = CSng(txtPropertyCoeff(i).text)
    Next i
        
    lngRet = WritePrivateProfileString("PARAMETER", "MaxTemperature", txtParaNormal(0).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "PreheatIntensity", txtParaNormal(1).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "PumpTimeout", txtParaNormal(2).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "PumpDown", txtParaNormal(3).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "APCGaugePressureValue", txtParaNormal(4).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "PUMPINGDELAY", txtParaNormal(5).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "TCDifferential", txtParaNormal(6).text, StrFileName)
    'Rev 7.1.0.3
    lngRet = WritePrivateProfileString("PARAMETER", "VentGate", txtParaNormal(7).text, StrFileName)
    'Rev 12.0.0.2 add intensity limit
    lngRet = WritePrivateProfileString("PARAMETER", "IntensityLimit", txtParaNormal(8).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "FinishedBeep", txtParaNormal(9).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "MinTemperature", txtParaNormal(10).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "OpenTemperature", txtParaNormal(28).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "MaxMonitorError", txtParaNormal(14).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "MaxMonitorTime", txtParaNormal(15).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "FinishedLight", txtParaNormal(16).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CycleRun", txtParaNormal(17).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "GaugeValue", txtParaNormal(18).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "IdleWarning", txtParaNormal(19).text, StrFileName)
    'Rev4.1.4
    lngRet = WritePrivateProfileString("PARAMETER", "UniformityTest", CStr(chkUniformity.value), StrFileName)
    
    lngRet = WritePrivateProfileString("PARAMETER", "Overheat", txtParaHeat(0).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "PMTemperature", txtParaHeat(1).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "PMCVT1", txtParaHeat(2).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "PMCVT2", txtParaHeat(3).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "PreheatTimeout", txtParaHeat(4).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "IntensityKeep", txtParaHeat(5).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "TCCVT1", txtParaHeat(6).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "TCCVT2", txtParaHeat(7).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "NumOfBanks", txtParaHeat(8).text, StrFileName)
    
    lngRet = WritePrivateProfileString("PARAMETER", "Gas1Active", CStr(chkGasEnable(0).value), StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas2Active", CStr(chkGasEnable(1).value), StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas3Active", CStr(chkGasEnable(2).value), StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas4Active", CStr(chkGasEnable(3).value), StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas5Active", CStr(chkGasEnable(4).value), StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas6Active", CStr(chkGasEnable(5).value), StrFileName)
    lngRet = WritePrivateProfileString("Gas7", "Gas7Active", CStr(chkGasEnable(6).value), ExtendPath)
    For i = 0 To 6
    If chkGasEnable(i).value = 1 Then NewGasCount = NewGasCount + 1
    Next i
    lngRet = WritePrivateProfileString("PARAMETER", "Gas1Alias", txtParaGasAlias(0).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas2Alias", txtParaGasAlias(1).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas3Alias", txtParaGasAlias(2).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas4Alias", txtParaGasAlias(3).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas5Alias", txtParaGasAlias(4).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas6Alias", txtParaGasAlias(5).text, StrFileName)
    lngRet = WritePrivateProfileString("Gas7", "Gas7Alias", txtParaGasAlias(6).text, ExtendPath)
    
    lngRet = WritePrivateProfileString("PARAMETER", "Gas1Unit", txtParaGasUnit(0).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas2Unit", txtParaGasUnit(1).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas3Unit", txtParaGasUnit(2).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas4Unit", txtParaGasUnit(3).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas5Unit", txtParaGasUnit(4).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas6Unit", txtParaGasUnit(5).text, StrFileName)
    lngRet = WritePrivateProfileString("Gas7", "Gas7Unit", txtParaGasUnit(6).text, ExtendPath)
    
    lngRet = WritePrivateProfileString("PARAMETER", "Gas1SLMP", txtParaGasValue(0).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas2SLMP", txtParaGasValue(1).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas3SLMP", txtParaGasValue(2).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas4SLMP", txtParaGasValue(3).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas5SLMP", txtParaGasValue(4).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas6SLMP", txtParaGasValue(5).text, StrFileName)
    lngRet = WritePrivateProfileString("Gas7", "Gas7SLMP", txtParaGasValue(6).text, ExtendPath)
    
    lngRet = WritePrivateProfileString("PARAMETER", "Gas1Bias", txtParaGasBias(0).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas2Bias", txtParaGasBias(1).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas3Bias", txtParaGasBias(2).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas4Bias", txtParaGasBias(3).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas5Bias", txtParaGasBias(4).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas6Bias", txtParaGasBias(5).text, StrFileName)
    lngRet = WritePrivateProfileString("Gas7", "Gas7Bias", txtParaGasBias(6).text, ExtendPath)
    
    lngRet = WritePrivateProfileString("PARAMETER", "Gas1Error", txtParaGasError(0).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas2Error", txtParaGasError(1).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas3Error", txtParaGasError(2).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas4Error", txtParaGasError(3).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas5Error", txtParaGasError(4).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas6Error", txtParaGasError(5).text, StrFileName)
    lngRet = WritePrivateProfileString("Gas7", "Gas7Error", txtParaGasError(6).text, ExtendPath)
    
    lngRet = WritePrivateProfileString("PARAMETER", "Gas1NError", txtParaGasErrorN(0).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas2NError", txtParaGasErrorN(1).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas3NError", txtParaGasErrorN(2).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas4NError", txtParaGasErrorN(3).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas5NError", txtParaGasErrorN(4).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas6NError", txtParaGasErrorN(5).text, StrFileName)
    lngRet = WritePrivateProfileString("Gas7", "Gas7NError", txtParaGasErrorN(6).text, ExtendPath)
    
    lngRet = WritePrivateProfileString("PARAMETER", "Gas1Unit", txtParaGasUnit(0).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas2Unit", txtParaGasUnit(1).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas3Unit", txtParaGasUnit(2).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas4Unit", txtParaGasUnit(3).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas5Unit", txtParaGasUnit(4).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "Gas6Unit", txtParaGasUnit(5).text, StrFileName)
    lngRet = WritePrivateProfileString("Gas7", "Gas7Unit", txtParaGasUnit(6).text, ExtendPath)
    
    lngRet = WritePrivateProfileString("PARAMETER", "VacuumGaugePara", txtParaVacuum(0).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "AngleOpenDelay", txtParaVacuum(1).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "ReleaseOpenDelay", txtParaVacuum(2).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "ThrottleFullOpenDelay", txtParaVacuum(3).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "ThrottleInitialPos", txtParaVacuum(4).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "APCGaugeValveLimit", txtParaVacuum(5).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "GaugeZoomIn", txtParaVacuum(6).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "APCInterval", txtParaVacuum(7).text, StrFileName)
    
    lngRet = WritePrivateProfileString("PARAMETER", "KeepPurge", txtParaVacuum(9).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "APC_MFC_Port", txtParaVacuum(11).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "MFC_Ratio", txtParaVacuum(12).text, StrFileName)
            
    lngRet = WritePrivateProfileString("PARAMETER", "IntensityRef1", txtIntensityRef(0).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "IntensityRef2", txtIntensityRef(1).text, StrFileName)
    
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate1", txtParaCTGate(0).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate2", txtParaCTGate(1).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate3", txtParaCTGate(2).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate4", txtParaCTGate(3).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate5", txtParaCTGate(4).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate6", txtParaCTGate(5).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate7", txtParaCTGate(6).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate8", txtParaCTGate(7).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate9", txtParaCTGate(8).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate10", txtParaCTGate(9).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate11", txtParaCTGate(10).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate12", txtParaCTGate(11).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate13", txtParaCTGate(12).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate14", txtParaCTGate(13).text, StrFileName)
    
    lngRet = WritePrivateProfileString("PARAMETER", "CTAlertGateWeight", txtCTAlertGateWeight.text, StrFileName)

    lngRet = WritePrivateProfileString("PARAMETER", "ResetIntegral", CStr(chkResetIntegral.value), StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "LampMonitor", CStr(chkLampMonitor.value), StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "RampSmooth", CStr(chkSmoothRamp.value), StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "SmoothDisplay", CStr(chkSmoothDisplay.value), StrFileName)
    'lngRet = WritePrivateProfileString("PARAMETER", "SmoothTime", txtSmoothTime.Text, strFileName) 'ms
    'lngRet = WritePrivateProfileString("PARAMETER", "MultiLoop", CStr(chkMultiLoop.value), strFileName)
    'Ver 4.1.2 Modify Control Mode
    lngRet = WritePrivateProfileString("PARAMETER", "MultiLoop", CStr(cmbControlMode.ListIndex), StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "MonitorTC", CStr(chkMonitorTC.value), StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "AlarmBuzzer", CStr(chkAlarmBuzzer.value), StrFileName)
  
    'Rev8.0.1.7
    lngRet = WritePrivateProfileString("PARAMETER", "MonitorTC1Active", CStr(chkMonitorTCActive(0).value), StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "MonitorTC2Active", CStr(chkMonitorTCActive(1).value), StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "MonitorTC3Active", CStr(chkMonitorTCActive(2).value), StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "MonitorTC4Active", CStr(chkMonitorTCActive(3).value), StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "MonitorTC5Active", CStr(chkMonitorTCActive(4).value), StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "MonitorTC6Active", CStr(chkMonitorTCActive(5).value), StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "MonitorTC7Active", CStr(chkMonitorTCActive(6).value), StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "MonitorTC8Active", CStr(chkMonitorTCActive(7).value), StrFileName)
'    lngRet = WritePrivateProfileString("FuncSwitch", "StopTCM", CStr(ckeStopTCM.value), App.Path + Function_Path)
'    lngRet = WritePrivateProfileString("FuncSwitch", "ProcStepInUse", CStr(ChkDefineProcStep.value), App.Path + Function_Path)
'    lngRet = WritePrivateProfileString("FuncSwitch", "OffsetWriteToTcm", CStr(ChkWriteOffsetToTCM.value), App.Path + Function_Path)
    
    
'    lngRet = WritePrivateProfileString("FuncSwitch", "ForcePreheat", CStr(ckeForcePreheat.value), App.Path + Function_Path)
'    lngRet = WritePrivateProfileString("FuncSwitch", "CTDisplay", CStr(ckeCTDisplay.value), App.Path + Function_Path)
    If (CInt(txtNumber1.text) > 20) Then txtNumber1.text = 20
    lngRet = WritePrivateProfileString("CTConfig", "CTNumber1", txtNumber1.text, App.Path + DeviceInfo_Path)
    If (CInt(txtNumber2.text) > 20) Then txtNumber2.text = 20
    lngRet = WritePrivateProfileString("CTConfig", "CTNumber2", txtNumber2.text, App.Path + DeviceInfo_Path)
    If (CInt(txtNumber3.text) > 20) Then txtNumber3.text = 20
    lngRet = WritePrivateProfileString("CTConfig", "CTNumber3", txtNumber3.text, App.Path + DeviceInfo_Path)
    If (CInt(txtNumber4.text) > 20) Then txtNumber4.text = 20
    lngRet = WritePrivateProfileString("CTConfig", "CTNumber4", txtNumber4.text, App.Path + DeviceInfo_Path)
    
    If txtCTNumber.text = "" Then
        txtCTNumber.text = CInt(txtNumber1.text) + CInt(txtNumber2.text) + CInt(txtNumber3.text) + CInt(txtNumber4.text)
    End If
    
    lngRet = WritePrivateProfileString("CTConfig", "CTName1", cmbCTName1(0).text, App.Path + DeviceInfo_Path)
    lngRet = WritePrivateProfileString("CTConfig", "CTName2", cmbCTName2(0).text, App.Path + DeviceInfo_Path)
    lngRet = WritePrivateProfileString("CTConfig", "CTName3", cmbCTName3(0).text, App.Path + DeviceInfo_Path)
    lngRet = WritePrivateProfileString("CTConfig", "CTName4", cmbCTName4(0).text, App.Path + DeviceInfo_Path)
    
    lngRet = WritePrivateProfileString("CTConfig", "CTOrder1", cmbOrder1(0).text, App.Path + DeviceInfo_Path)
    lngRet = WritePrivateProfileString("CTConfig", "CTOrder2", cmbOrder2(0).text, App.Path + DeviceInfo_Path)
    lngRet = WritePrivateProfileString("CTConfig", "CTOrder3", cmbOrder3(0).text, App.Path + DeviceInfo_Path)
    lngRet = WritePrivateProfileString("CTConfig", "CTOrder4", cmbOrder4(0).text, App.Path + DeviceInfo_Path)
    
    lngRet = WritePrivateProfileString("CTConfig", "CTNumbers", txtCTNumber.text, App.Path + DeviceInfo_Path)
    lngRet = WritePrivateProfileString("PARAMETER", "TCTYPE", CStr(cmbTCType.ListIndex), StrFileName)
    gbintTCType = cmbTCType.ListIndex
    lngRet = WritePrivateProfileString("PARAMETER", "TCVoltageRange", CStr(cmbTCVoltageRange.ListIndex), StrFileName)
    gbintTCVoltageRange = cmbTCVoltageRange.ListIndex
    
    'Assign parameter to variable
    gbsngMaxTemperature = Val(txtParaNormal(0).text)
    If (gbsngMaxTemperature > 1500) Then gbsngMaxTemperature = 1500
    gbintPreheatIntensity = CSng(txtParaNormal(1).text)
    gbintPumpTimeout = CInt(txtParaNormal(2).text)
'    If gbintPumpTimeout > 65 Then
'        gbintPumpTimeout = 65
'        txtParaNormal(2).Text = "65"
'    End If
    gbsngPumpDownGate = Val(txtParaNormal(3).text)
    gbsngAPCGaugePressureValue = Val(txtParaNormal(4).text)
    gbsngPumpingDelay = Val(txtParaNormal(5).text)
    gbsngTCDifferentialRange = Val(txtParaNormal(6).text)
    If (Val(txtParaNormal(7).text) > 500) Then txtParaNormal(7).text = "500"
    If (Val(txtParaNormal(7).text) < 100) Then txtParaNormal(7).text = "100"
    gbsngVentGate = Val(txtParaNormal(7).text)
    gbsngIntensityLimit = Val(txtParaNormal(8).text)
    gbintFinishedBeep = Val(txtParaNormal(9).text)
    gbsngMinTemperature = Val(txtParaNormal(10).text)
    gbsngOpenTemperature = Val(txtParaNormal(28).text)
    gbsngLifeLamp = Val(txtParaNormal(11).text) * 3600
    gbsngUsedLamp = Val(txtParaNormal(12).text) * 3600
    gbsngMaxMonitorError = Val(txtParaNormal(14).text)
    gbsngMaxMonitorTime = Val(txtParaNormal(15).text)
    gbintFinishedLight = Val(txtParaNormal(16).text)
    gbintCycleRun = Val(txtParaNormal(17).text)
    gbsngGaugeValue = Val(txtParaNormal(18).text)
    gbsngIdleWarning = Val(txtParaNormal(19).text) * 60
    
        
    gbsngChamberOverheat = Val(txtParaHeat(0).text)
    gbsngValidPMTempature = Val(txtParaHeat(1).text)
    gbsngPMCVT1 = Val(txtParaHeat(2).text)
    gbsngPMCVT2 = Val(txtParaHeat(3).text)
    gbdblProcessPreheatTimerout = CDbl(txtParaHeat(4).text)
    gbsngIntensityKeep = CSng(Val(txtParaHeat(5).text) / 10)
    gbsngTCCVT1 = Val(txtParaHeat(6).text)
    gbsngTCCVT2 = Val(txtParaHeat(7).text)
    gbintNumOfBanks = CInt(txtParaHeat(8).text)
    Txt_RValueRange.text = CommnonReadini("Loop1-5", "RValRanage", App.Path + ProcDict_Path)
    gbRValRange = Txt_RValueRange.text
    'Rev9.0.0.0
    advThermo.TC_CVT1 = gbsngTCCVT1
    advThermo.TC_CVT2 = gbsngTCCVT2
    'Rev8.0.1.7
    gbintMonitorTCActive(0) = CInt(chkMonitorTCActive(0).value)
    gbintMonitorTCActive(1) = CInt(chkMonitorTCActive(1).value)
    gbintMonitorTCActive(2) = CInt(chkMonitorTCActive(2).value)
    gbintMonitorTCActive(3) = CInt(chkMonitorTCActive(3).value)
    gbintMonitorTCActive(4) = CInt(chkMonitorTCActive(4).value)
    gbintMonitorTCActive(5) = CInt(chkMonitorTCActive(5).value)
    gbintMonitorTCActive(6) = CInt(chkMonitorTCActive(6).value)
    gbintMonitorTCActive(7) = CInt(chkMonitorTCActive(7).value)
    
    gbintMaxGasEnable = -1
    For j = 0 To txtParaGasAlias.UBound
        gbintGasEnable(j) = chkGasEnable(j).value
        gbstrGasAlias(j) = txtParaGasAlias(j).text
        gbstrGasUnit(j) = txtParaGasUnit(j).text
        gbsngMaxGasSLMP(j) = Val(txtParaGasValue(j).text)
        gbsngGasBias(j) = Val(txtParaGasBias(j).text)
        gbsngGasError(j) = Val(txtParaGasError(j).text)
        gbsngGasErrorN(j) = Val(txtParaGasErrorN(j).text)
        If (chkGasEnable(j).value = 1) Then
            gbintMaxGasEnable = gbintMaxGasEnable + 1
        End If
    Next j

    gbsngVacuumGaugeCompensation = Val(txtParaVacuum(0).text)
    gbintAngleOpenDelay = CInt(txtParaVacuum(1).text)
    gbintReleaseOpenDelay = CInt(txtParaVacuum(2).text)
    gbintThrottleFullOpenDelay = CInt(txtParaVacuum(3).text)
    gbintThrottleInitialPos = CInt(txtParaVacuum(4).text)
    gbsngAPCGaugeValveLimit = Val(txtParaVacuum(5).text)
    gbsngGaugeZoomIn = Val(txtParaVacuum(6).text)
    gbsngAPCInterval = Val(txtParaVacuum(7).text)
   
    gbsngKeepPurge = Val(txtParaVacuum(9).text)
    gbintAPC_MFC_Port = Val(txtParaVacuum(11).text)
    gbintMFC_Ratio = Val(txtParaVacuum(12).text)
        
    gbdblProcessPumpDownTimerout = CDbl(CSng(gbintPumpTimeout) * 1000)
    
    gbblnResetInteral = IIf(chkResetIntegral.value = 1, True, False)
    gbintLampMonitor = CInt(chkLampMonitor.value)
    gbintAlarmBuzzer = CInt(chkAlarmBuzzer.value)
    gbintMonitorTC = CInt(chkMonitorTC.value)
    
    
       
    gbintRampSmooth = CInt(chkSmoothRamp.value)
    'gbsngSmoothTime = CSng(Val(txtSmoothTime.Text))
    gbintSmoothDisplay = CInt(chkSmoothDisplay.value)
    
    lngRet = WritePrivateProfileString("PARAMETER", "IntensityRef1", txtIntensityRef(0).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "IntensityRef2", txtIntensityRef(1).text, StrFileName)
    
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate1", txtParaCTGate(0).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate2", txtParaCTGate(1).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate3", txtParaCTGate(2).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate4", txtParaCTGate(3).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate5", txtParaCTGate(4).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate6", txtParaCTGate(5).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate7", txtParaCTGate(6).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate8", txtParaCTGate(7).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate9", txtParaCTGate(8).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate10", txtParaCTGate(9).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate11", txtParaCTGate(10).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate12", txtParaCTGate(11).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate13", txtParaCTGate(12).text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "CTGate14", txtParaCTGate(13).text, StrFileName)
    
    lngRet = WritePrivateProfileString("PARAMETER", "CTAlertGateWeight", txtCTAlertGateWeight.text, StrFileName)
    gbsngIntensityRef(0) = Val(txtIntensityRef(0).text)
    gbsngIntensityRef(1) = Val(txtIntensityRef(1).text)
    For i = 0 To 6
        gbsngCTGate1(i) = txtParaCTGate(i)
        gbsngCTGate2(i) = txtParaCTGate(i + 7)
    Next i
    
    gbsngCTGateWeight = Val(txtCTAlertGateWeight.text)
    
    
    lngRet = WritePrivateProfileString("PARAMETER", "UniStartPointRamp", txtUniformityRampStartPoint.text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "UniSubWeightD1", txtSubWeightD1.text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "UniSubWeightD2", txtSubWeightD2.text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "UniStartPointHold", txtUniformityHoldStartPoint.text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "UniSubWeight1", txtSubWeight1.text, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "UniSubWeight2", txtSubWeight2.text, StrFileName)
    
    
    gbsngUniformitySubWeightD1 = Val(txtSubWeightD1.text)
    gbsngUniformitySubWeightD2 = Val(txtSubWeightD2.text)
    gbsngUniformityStartPointHold = Val(txtUniformityHoldStartPoint.text)
    gbsngUniformitySubWeight1 = Val(txtSubWeight1.text)
    gbsngUniformitySubWeight2 = Val(txtSubWeight2.text)
        
    lngRet = WritePrivateProfileString("ALARM_ACTIVE", "System", CStr(chkAlarmEnable(0).value), StrFileName)
    lngRet = WritePrivateProfileString("ALARM_ACTIVE", "Ready", CStr(chkAlarmEnable(1).value), StrFileName)
    lngRet = WritePrivateProfileString("ALARM_ACTIVE", "EMS", CStr(chkAlarmEnable(2).value), StrFileName)
    lngRet = WritePrivateProfileString("ALARM_ACTIVE", "BZR", CStr(chkAlarmEnable(3).value), StrFileName)
    lngRet = WritePrivateProfileString("ALARM_ACTIVE", "OH", CStr(chkAlarmEnable(4).value), StrFileName)
    lngRet = WritePrivateProfileString("ALARM_ACTIVE", "CT", CStr(chkAlarmEnable(5).value), StrFileName)
    lngRet = WritePrivateProfileString("ALARM_ACTIVE", "TC", CStr(chkAlarmEnable(6).value), StrFileName)
    lngRet = WritePrivateProfileString("ALARM_ACTIVE", "Water", CStr(chkAlarmEnable(7).value), StrFileName)
    lngRet = WritePrivateProfileString("ALARM_ACTIVE", "CDA", CStr(chkAlarmEnable(8).value), StrFileName)
    lngRet = WritePrivateProfileString("ALARM_ACTIVE", "VG", CStr(chkAlarmEnable(9).value), StrFileName)
    lngRet = WritePrivateProfileString("ALARM_ACTIVE", "VS", CStr(chkAlarmEnable(10).value), StrFileName)
    lngRet = WritePrivateProfileString("ALARM_ACTIVE", "Door", CStr(chkAlarmEnable(11).value), StrFileName)
    lngRet = WritePrivateProfileString("ALARM_ACTIVE", "APC", CStr(chkAlarmEnable(12).value), StrFileName)
    lngRet = WritePrivateProfileString("ALARM_ACTIVE", "Chamber", CStr(chkAlarmEnable(13).value), StrFileName)
'    lngRet = WritePrivateProfileString("FuncSwitch", "CheckTcWafer", CStr(chkAlarmEnable(14).value), App.Path + Function_Path)
'    gbintActiveAlarm_TcWafer = chkAlarmEnable(14).value
    
'    lngRet = WritePrivateProfileString("FuncSwitch", "CheckRValue", CStr(chkAlarmEnable(15).value), App.Path + Function_Path)
'    gbintActiveAlarm_RValue = chkAlarmEnable(15).value
     '--------------------------------------------------------------------
    'Define Alarm Activity Enable/Disable
    '--------------------------------------------------------------------
    '***System status***
    gbintActiveAlarm_System = chkAlarmEnable(0).value
    gbintActiveAlarm_Ready = chkAlarmEnable(1).value
    
    gbintActiveAlarm_EMS = chkAlarmEnable(2).value
    gbintActiveAlarm_Buzzer = chkAlarmEnable(3).value
    
    '***The Heating Module***
    gbintActiveAlarm_Overheat = chkAlarmEnable(4).value
    gbintActiveAlarm_CT = chkAlarmEnable(5).value
    gbintActiveAlarm_TC = chkAlarmEnable(6).value
    
    '***The Gas Module***
    '***The Cooling Module***
    gbintActiveAlarm_Water = chkAlarmEnable(7).value
    gbintActiveAlarm_Air = chkAlarmEnable(8).value
    '***The Vacuum Module***
    gbintActiveAlarm_VacuumGauge = chkAlarmEnable(9).value
    gbintActiveAlarm_VacuumSwitch = chkAlarmEnable(10).value
    gbintActiveAlarm_Door = chkAlarmEnable(11).value
    gbintActiveAlarm_APC = chkAlarmEnable(12).value
    gbintActiveAlarm_ChamberGate = chkAlarmEnable(13).value
    
    lngRet = WritePrivateProfileString("MODULE_ACTIVE", "H", CStr(chkModuleEnable(0).value), StrFileName)
    lngRet = WritePrivateProfileString("MODULE_ACTIVE", "C", CStr(chkModuleEnable(1).value), StrFileName)
    lngRet = WritePrivateProfileString("MODULE_ACTIVE", "G", CStr(chkModuleEnable(2).value), StrFileName)
    lngRet = WritePrivateProfileString("MODULE_ACTIVE", "V", CStr(chkModuleEnable(3).value), StrFileName)
    lngRet = WritePrivateProfileString("MODULE_ACTIVE", "F", CStr(chkModuleEnable(4).value), StrFileName)
    lngRet = WritePrivateProfileString("MODULE_ACTIVE", "O", CStr(chkModuleEnable(5).value), StrFileName)
    lngRet = WritePrivateProfileString("MODULE_ACTIVE", "B", CStr(chkModuleEnable(6).value), StrFileName)
    lngRet = WritePrivateProfileString("MODULE_ACTIVE", "D", CStr(chkModuleEnable(7).value), StrFileName)
    lngRet = WritePrivateProfileString("MODULE_ACTIVE", "P", CStr(chkModuleEnable(8).value), StrFileName)
    lngRet = WritePrivateProfileString("MODULE_ACTIVE", "A", CStr(chkModuleEnable(9).value), StrFileName)
    lngRet = WritePrivateProfileString("MODULE_ACTIVE", "R", CStr(chkModuleEnable(10).value), StrFileName)
    lngRet = WritePrivateProfileString("MODULE_ACTIVE", "S", CStr(chkModuleEnable(13).value), StrFileName)
    lngRet = WritePrivateProfileString("MODULE_ACTIVE", "M", CStr(chkModuleEnable(14).value), StrFileName)
    
    '--------------------------------------------------------------------
    'Define Module Activity Enable/Disable
    '--------------------------------------------------------------------
    gbintActiveModule_Heating = chkModuleEnable(0).value
    gbintActiveModule_Cooling = chkModuleEnable(1).value
    gbintActiveModule_Gas = chkModuleEnable(2).value
    gbintActiveModule_Vacuum = chkModuleEnable(3).value
    gbintActiveModule_Database = chkModuleEnable(4).value
    gbintActiveModule_Oxygen = chkModuleEnable(5).value
    gbintActiveModule_Barcode = chkModuleEnable(6).value
    gbintActiveModule_Door = chkModuleEnable(7).value
    gbintActiveModule_PNRecipe = chkModuleEnable(8).value
    gbintActiveModule_APC = chkModuleEnable(9).value
    gbintActiveModule_Auto = chkModuleEnable(10).value
    gbintActiveModule_CIM = chkModuleEnable(13).value
    gbintActiveModule_MLoop = chkModuleEnable(14).value
    Para.UseCT = chkModuleEnable(11).value
    Para.UseMTC = chkModuleEnable(12).value
    Para.UseMTCB = chkModuleEnable(19).value
    
       
    
    For i = 0 To 32
        lngRet = WritePrivateProfileString("IO91141", "DI" & CStr(i), fgDI.TextMatrix(i + 1, 1), StrFileName)
    Next i
    For i = 0 To 31
        lngRet = WritePrivateProfileString("IO91141", "DO" & CStr(i), fgDO.TextMatrix(i + 1, 1), StrFileName)
    Next i
    For i = 0 To 31
        lngRet = WritePrivateProfileString("IO91141", "AI" & CStr(i), fgAI.TextMatrix(i + 1, 1), StrFileName)
        lngRet = WritePrivateProfileString("IO91141", "AI_ErrorV" & CStr(i), fgAI.TextMatrix(i + 1, 3), StrFileName)
        SysAI.ErrorV(i) = Val(fgAI.TextMatrix(i + 1, 3))
    Next i
    For i = 0 To 31
        lngRet = WritePrivateProfileString("IO62081", "AO" & CStr(i), fgAO.TextMatrix(i + 1, 1), StrFileName)
        lngRet = WritePrivateProfileString("IO62081", "AO_Type" & CStr(i), fgAO.TextMatrix(i + 1, 2), StrFileName)
        gbintAO_Type(i) = CInt(fgAO.TextMatrix(i + 1, 2))
    Next i
    For i = 0 To 23
        lngRet = WritePrivateProfileString("ADVIO17101", "AI" & CStr(i), fgAdvDaqAI.TextMatrix(i + 1, 1), StrFileName)
        lngRet = WritePrivateProfileString("ADVIO17101", "Power" & CStr(i), fgAdvDaqAI.TextMatrix(i + 1, 2), StrFileName)
        gbsngPowerTC(i) = CSng(fgAdvDaqAI.TextMatrix(i + 1, 2))
        lngRet = WritePrivateProfileString("ADVIO17101", "Error" & CStr(i), fgAdvDaqAI.TextMatrix(i + 1, 3), StrFileName)
        gbsngErrorTC(i) = CSng(fgAdvDaqAI.TextMatrix(i + 1, 3))
        lngRet = WritePrivateProfileString("ADVIO17101", "Ratio" & CStr(i), fgAdvDaqAI.TextMatrix(i + 1, 4), StrFileName)
        gbsngRatioTC(i) = CSng(fgAdvDaqAI.TextMatrix(i + 1, 4))
        lngRet = WritePrivateProfileString("ADVIO17101", "Power3C" & CStr(i), fgAdvDaqAI.TextMatrix(i + 1, 6), StrFileName)
        gbsngPower3C(i) = CSng(fgAdvDaqAI.TextMatrix(i + 1, 6))
        lngRet = WritePrivateProfileString("ADVIO17101", "Power4C" & CStr(i), fgAdvDaqAI.TextMatrix(i + 1, 7), StrFileName)
        gbsngPower4C(i) = CSng(fgAdvDaqAI.TextMatrix(i + 1, 7))
        lngRet = WritePrivateProfileString("ADVIO17101", "Power5C" & CStr(i), fgAdvDaqAI.TextMatrix(i + 1, 8), StrFileName)
        gbsngPower5C(i) = CSng(fgAdvDaqAI.TextMatrix(i + 1, 8))
        lngRet = WritePrivateProfileString("ADVIO17101", "LoopNo" & CStr(i), fgAdvDaqAI.TextMatrix(i + 1, 9), StrFileName)
        gbintLoopNo(i) = CSng(fgAdvDaqAI.TextMatrix(i + 1, 9))
        lngRet = WritePrivateProfileString("PrecisionDigit", "PrecisionDigit" & CStr(i), fgAdvDaqAI.TextMatrix(i + 1, 10), App.Path + ProcDict_Path)
        gbintPrecisionDigit(i) = CSng(fgAdvDaqAI.TextMatrix(i + 1, 10))
        
    Next i
    
      
    gbintCTCheck = chkCTCheck.value
    lngRet = WritePrivateProfileString("Utility", "CTCheck", CStr(chkCTCheck.value), StrFileName)
    gbintRtaType = Val(txtRtaType.text)
    lngRet = WritePrivateProfileString("Utility", "RtaType", txtRtaType.text, StrFileName)
    gbintAutoDeleteRecipe = chkAutoDeleteRecipe.value
    lngRet = WritePrivateProfileString("Utility", "AutoDeleteRecipe", CStr(chkAutoDeleteRecipe.value), StrFileName)
        
    
    gbsngLifeLamp = Val(txtParaNormal(11).text) * 3600
    gbsngUsedLamp = Val(txtParaNormal(12).text) * 3600
    gbsngMaxMonitorError = Val(txtParaNormal(14).text)
    gbsngMaxMonitorTime = Val(txtParaNormal(15).text)
    gbstrLogFilePath = txtParaNormal(13).text
    gbstrRecipeFilePath = txtParaNormal(20).text
    lngRet = WritePrivateProfileString("PARAMETER", "LifeLamp", CStr(gbsngLifeLamp), StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "UsedLamp", CStr(gbsngUsedLamp), StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "LogFilePath", gbstrLogFilePath, StrFileName)
    lngRet = WritePrivateProfileString("PARAMETER", "RecipeFilePath", gbstrRecipeFilePath, StrFileName)
    
    '120822 Josh
    lngRet = WritePrivateProfileString("User", "ad", txtAdmin.text, StrFileName)
    lngRet = WritePrivateProfileString("User", "eg", txtEngineer.text, StrFileName)
    lngRet = WritePrivateProfileString("User", "op", txtOperator.text, StrFileName)
    lngRet = WritePrivateProfileString("User", "Page0", txtActivePage(0).text, StrFileName)
    lngRet = WritePrivateProfileString("User", "Page1", txtActivePage(1).text, StrFileName)
    lngRet = WritePrivateProfileString("User", "Page2", txtActivePage(2).text, StrFileName)

    
    gbintActivePage(0) = Val(txtActivePage(0).text)
    gbintActivePage(1) = Val(txtActivePage(1).text)
    gbintActivePage(2) = Val(txtActivePage(2).text)
    gbstrAdminPW = txtAdmin.text
    gbstrEngineerPW = txtEngineer.text
    gbstrOperatorPW = txtOperator.text
    
'    For i = 0 To GB_MAX_LOOPS - 1
'        gbsngRatioEX(i) = Val(txtRatioEX(i).text)
'        lngRet = WritePrivateProfileString("PARAMETER", "RatioEX" & CStr(i), txtRatioEX(i).text, strFileName)
'    Next i
    
'    For i = 0 To GB_MAX_LOOPS - 1
'
'        If Val(txtRatioEX(i).text) >= -100 And Val(txtRatioEX(i).text) <= 100 Then
'
'            gbsngRatioEX(i) = Val(txtRatioEX(i).text)
'            lngRet = WritePrivateProfileString("PARAMETER", "RatioEX" & CStr(i), txtRatioEX(i).text, strFileName)
'
'            If i <> GB_MAX_LOOPS - 1 Then
'            gbsngErrorTC(i) = Val(txtRatioEX(i + 1).text)
'            lngRet = WritePrivateProfileString("ADVIO17101", "Error" & CStr(i), txtRatioEX(i + 1).text, strFileName)
'            End If
'        Else
'            MsgBox "TC" + CStr(i) + "對應offset請輸入-100到100之間的數值！"
'        End If
'    Next i
    
    gbsngRecipeTempDownTimeout = Val(txtTempDownTimeout.text)
    lngRet = WritePrivateProfileString("InterLock", "TempDownTimeout", CStr(gbsngRecipeTempDownTimeout), StrFileName)
    gbsngRecipeGatePS1 = Val(txtGatePS1.text)
    lngRet = WritePrivateProfileString("InterLock", "GatePS1", CStr(gbsngRecipeGatePS1), StrFileName)
    gbsngRecipeGatePS2 = Val(txtGatePS2.text)
    lngRet = WritePrivateProfileString("InterLock", "GatePS2", CStr(gbsngRecipeGatePS2), StrFileName)
    
    
    gbintRobotPort = Val(txtRobotPort.text)
    lngRet = WritePrivateProfileString("Robot", "RobotPort", txtRobotPort.text, StrFileName)
    gbsngPickH = Val(txtPickH.text)
    lngRet = WritePrivateProfileString("Robot", "RobotPickH", txtPickH.text, StrFileName)
    gbsngPlaceH = Val(txtPlaceH.text)
    lngRet = WritePrivateProfileString("Robot", "RobotPlaceH", txtPlaceH.text, StrFileName)
    
    For i = 0 To 3
        If optRobotSpeed(i).value = True Then gbintRobotSpeed = i
    Next i
    lngRet = WritePrivateProfileString("Robot", "RobotSpeed", CStr(gbintRobotSpeed), StrFileName)
    
    For i = 0 To 49
        For j = 0 To 3
            gbsngTeach(i, j) = Val(fgTeach.TextMatrix(i + 1, j + 1))
            lngRet = WritePrivateProfileString("Robot", "Teach" & CStr(i) & "_" & CStr(j), fgTeach.TextMatrix(i + 1, j + 1), StrFileName)
        Next j
    Next i
    
    
    Para.intOnlyRecipe = chkOnlyRecipe.value
    Para.sngGaugeAngle = Val(txtParaNormal(21).text)
    Para.intMonitorIndex = Val(txtParaNormal(22).text)
    Para.UseBarcodeServer = chkBarcodeServer.value
    Para.intCycleRuns = Val(txtParaNormal(17).text)
    Para.RtaType = Val(txtRtaType.text)
    Para.strServerPath = txtServerPath.text
    Para.strServerPath = txtServerPath.text
    Para.intMonitorRuns = Val(txtMonitorRuns.text)
    Para.strTestRunKey = txtTestRunKey.text
    Para.IsHoldSafety = chkHoldSafety.value
    Para.IsCali = chkCalibration.value
    Para.intComCT = Val(txtComCT.text)
    Para.UseAutoMode = chkAutoMode.value
    Para.sngO2Gate = Val(txtParaNormal(23).text)
    Para.intLampAlarmTime = Val(txtParaNormal(24).text)
    Para.intOpenDoorTime = Val(txtParaNormal(25).text)
    Para.intAutoPort = Val(txtParaNormal(26).text)
    Para.intPumpDelay = Val(txtParaNormal(27).text)
    
    Para.UseCIM = chkCIM.value
    'Para.intCIMPort = Val(txtCIMPort.Text)
    Para.intCIMPort = cmbCIMPort.ListIndex
    Para.intPMbig = Val(txtPMbig.text)
    Para.intPMsmall = Val(txtPMsmall.text)
    
    Para.strAzIP1 = txtAzIP1.text
    Para.strAzIP2 = txtAzIP2.text
    Para.UseAz1 = chkModuleEnable(15).value
    Para.UseAz2 = chkModuleEnable(16).value
    Para.useTPump = chkModuleEnable(17).value
    Para.strRobotIP = txtRobotIP.text
        
    Para.sngGaugeD = Val(txtGaugeD.text)
    Para.sngGaugeVP = Val(txtGaugeVP.text)
    Para.sngGaugeVN = Val(txtGaugeVN.text)
    
    Para.UseCover = chkModuleEnable(18).value
    
    For i = 1 To 33
        Para.AlarmDo(i) = Val(fgAlarm.TextMatrix(i, 2))
    Next i
    SavePara
    
    Call frmHistory.AppendLogAlert(1, "Manual", 1101, "機台參數儲存", 1)
    Call InitialIO
    Call SetCTSlope
    Exit Sub

ERR_PARAMETER_SAVE:
    Call AlertShow("Save Parameter Failed!!", ERRORTYPE)
End Sub

Public Sub ParameterOpen()
    Dim i                   As Integer
    Dim j                   As Integer
    Dim lngRet                As Long
    Dim StrFileName         As String
    Dim iInputDevice        As Integer
    Dim strDataIntensity(20)    As String * 30
    Dim strDataIntensityS(20)    As String * 30
    Dim strPropertyCoeff(20) As String * 30
    
    Dim StrData(50)    As String * 30
    Dim StrTC(50)    As String * 4
    Dim strCoffData(20)    As String * 30
    Dim strDataKeepIntensity(1)    As String * 30
    
    Dim strDataResetIntegral    As String * 30
    
    Dim strDataNormal(30)           As String * 30
    Dim strDataHeating(10)          As String * 30
    Dim strDataHeatingSmooth(3)     As String * 30
    Dim strDataGas(50)              As String * 30
    Dim strDataVacuum(20)           As String * 30
    Dim strDataIntensityRef(2)      As String * 30
    Dim strDataCTGate(20)           As String * 30
    Dim strDataCTGateWeight(1)      As String * 30
    Dim strDataSpecial(10)          As String * 30
    Dim strAlarmEnable(20)          As String * 30
    Dim strModuleEnable(20)         As String * 30
    Dim strDataUniformity(10)       As String * 30 'Rev4.1.5
    Dim strDataMonitorTCActive(10)  As String * 30 'Rev8.0.1.7
    Dim strDataTCType               As String * 30
    Dim strDataTCVoltageRange       As String * 30
    Dim strPath                     As String * 100
    Dim strPage(3)    As String * 30
    Dim DefineSelf As Integer
    Dim ExtendPath As String
    
    ExtendPath = App.Path + ProcDict_Path
    
    On Error GoTo ERR_PARAMETER_OPEN
    
    StrFileName = gbSystemPath & "\System\system.cfg"
    If dir(StrFileName) = "" Then GoTo ERR_PARAMETER_OPEN
       
    
    'The Intensity weight for dynamic parameter
    For i = 0 To txtIntensityWeight.UBound
        lngRet = GetPrivateProfileString("PARAMETER", "IntensityWeight" & CStr(i + 1), "100", strDataIntensity(i), 20, StrFileName)
        txtIntensityWeight(i).text = strDataIntensity(i)
        gbsngIntensityWeight(i) = CSng(txtIntensityWeight(i).text) / 100
    Next i
    'The Intensity weight for stead state parameter
    For i = 0 To txtIntensityWeightS.UBound
        lngRet = GetPrivateProfileString("PARAMETER", "IntensityWeightS" & CStr(i + 1), "100", strDataIntensityS(i), 20, StrFileName)
        txtIntensityWeightS(i).text = strDataIntensityS(i)
        gbsngIntensityWeightS(i) = CSng(txtIntensityWeightS(i).text) / 100
    Next i
    
    'The property of PID coefficient
    Dim newParam As String
    For i = 0 To txtPropertyCoeff.UBound
        lngRet = GetPrivateProfileString("PARAMETER", "PropertyCoefficient" & CStr(i + 1), "1", strPropertyCoeff(i), 20, StrFileName)
'        If i = 4 Then
'            newParam = EncryptDecrypt(strPropertyCoeff(i), 123)
'            txtPropertyCoeff(i).text = Replace(newParam, "{", "")
'        Else
'            txtPropertyCoeff(i).text = strPropertyCoeff(i)
'        End If
        txtPropertyCoeff(i).text = strPropertyCoeff(i)
        gbsngPropertyCoefficient(i) = CSng(txtPropertyCoeff(i).text)
    Next i
    
    lngRet = GetPrivateProfileString("PARAMETER", "MaxTemperature", "0", strDataNormal(0), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "PreheatIntensity", "0", strDataNormal(1), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "PumpTimeout", "0", strDataNormal(2), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "PumpDown", "0", strDataNormal(3), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "APCGaugePressureValue", "0", strDataNormal(4), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "PUMPINGDELAY", "0", strDataNormal(5), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "TCDifferential", "0", strDataNormal(6), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "VentGate", "0", strDataNormal(7), 20, StrFileName)
    'Rev 12.0.0.2 add intensity limit
    lngRet = GetPrivateProfileString("PARAMETER", "IntensityLimit", "0", strDataNormal(8), 20, StrFileName)
    'Rev4.1.4
    lngRet = GetPrivateProfileString("PARAMETER", "UniformityTest", "0", strDataNormal(9), 20, StrFileName)
                
    lngRet = GetPrivateProfileString("PARAMETER", "Overheat", "0", strDataHeating(0), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "PMTemperature", "0", strDataHeating(1), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "PMCVT1", "1", strDataHeating(2), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "PMCVT2", "0", strDataHeating(3), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "PreheatTimeout", "0", strDataHeating(4), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "IntensityKeep", "0", strDataHeating(5), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "TCCVT1", "1", strDataHeating(6), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "TCCVT2", "0", strDataHeating(7), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "NumOfBanks", "5", strDataHeating(8), 20, StrFileName)
          
       
   lngRet = GetPrivateProfileString("PARAMETER", "Gas1Alias", "NA", strDataGas(0), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas2Alias", "NA", strDataGas(1), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas3Alias", "NA", strDataGas(2), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas4Alias", "NA", strDataGas(3), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas5Alias", "NA", strDataGas(4), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas6Alias", "NA", strDataGas(5), 20, StrFileName)
    lngRet = GetPrivateProfileString("Gas7", "Gas7Alias", "NA", strDataGas(6), 20, ExtendPath)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas1SLMP", "0", strDataGas(7), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas2SLMP", "0", strDataGas(8), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas3SLMP", "0", strDataGas(9), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas4SLMP", "0", strDataGas(10), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas5SLMP", "0", strDataGas(11), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas6SLMP", "0", strDataGas(12), 20, StrFileName)
    lngRet = GetPrivateProfileString("Gas7", "Gas7SLMP", "0", strDataGas(13), 20, ExtendPath)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas1Bias", "0", strDataGas(14), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas2Bias", "0", strDataGas(15), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas3Bias", "0", strDataGas(16), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas4Bias", "0", strDataGas(17), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas5Bias", "0", strDataGas(18), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas6Bias", "0", strDataGas(19), 20, StrFileName)
    lngRet = GetPrivateProfileString("Gas7", "Gas7Bias", "0", strDataGas(20), 20, ExtendPath)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas1Active", "0", strDataGas(21), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas2Active", "0", strDataGas(22), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas3Active", "0", strDataGas(23), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas4Active", "0", strDataGas(24), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas5Active", "0", strDataGas(25), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas6Active", "0", strDataGas(26), 20, StrFileName)
    lngRet = GetPrivateProfileString("Gas7", "Gas7Active", "0", strDataGas(27), 20, ExtendPath)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas1Unit", "0", strDataGas(28), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas2Unit", "0", strDataGas(29), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas3Unit", "0", strDataGas(30), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas4Unit", "0", strDataGas(31), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas5Unit", "0", strDataGas(32), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas6Unit", "0", strDataGas(33), 20, StrFileName)
    lngRet = GetPrivateProfileString("Gas7", "Gas7Unit", "0", strDataGas(34), 20, ExtendPath)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas1Error", "0", strDataGas(35), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas2Error", "0", strDataGas(36), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas3Error", "0", strDataGas(37), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas4Error", "0", strDataGas(38), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas5Error", "0", strDataGas(39), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas6Error", "0", strDataGas(40), 20, StrFileName)
    lngRet = GetPrivateProfileString("Gas7", "Gas7Error", "0", strDataGas(41), 20, ExtendPath)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas1NError", "0", strDataGas(42), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas2NError", "0", strDataGas(43), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas3NError", "0", strDataGas(44), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas4NError", "0", strDataGas(45), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas5NError", "0", strDataGas(46), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "Gas6NError", "0", strDataGas(47), 20, StrFileName)
    lngRet = GetPrivateProfileString("Gas7", "Gas7NError", "0", strDataGas(48), 20, ExtendPath)
    
    'Vacuum Parameter
    lngRet = GetPrivateProfileString("PARAMETER", "VacuumGaugePara", "0", strDataVacuum(0), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "AngleOpenDelay", "0", strDataVacuum(1), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "ReleaseOpenDelay", "0", strDataVacuum(2), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "ThrottleFullOpenDelay", "0", strDataVacuum(3), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "ThrottleInitialPos", "0", strDataVacuum(4), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "APCGaugeValveLimit", "0", strDataVacuum(5), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "GaugeZoomIn", "0", strDataVacuum(6), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "APCInterval", "0", strDataVacuum(7), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "KeepPurge", "60", strDataVacuum(9), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "APC_MFC_Port", "1", strDataVacuum(11), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "MFC_Ratio", "1", strDataVacuum(12), 20, StrFileName)
    'CT Parameter
    lngRet = GetPrivateProfileString("PARAMETER", "IntensityRef1", "0", strDataIntensityRef(0), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "IntensityRef2", "0", strDataIntensityRef(1), 20, StrFileName)

    lngRet = GetPrivateProfileString("PARAMETER", "CTGate1", "0", strDataCTGate(0), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "CTGate2", "0", strDataCTGate(1), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "CTGate3", "0", strDataCTGate(2), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "CTGate4", "0", strDataCTGate(3), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "CTGate5", "0", strDataCTGate(4), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "CTGate6", "0", strDataCTGate(5), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "CTGate7", "0", strDataCTGate(6), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "CTGate8", "0", strDataCTGate(7), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "CTGate9", "0", strDataCTGate(8), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "CTGate10", "0", strDataCTGate(9), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "CTGate11", "0", strDataCTGate(10), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "CTGate12", "0", strDataCTGate(11), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "CTGate13", "0", strDataCTGate(12), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "CTGate14", "0", strDataCTGate(13), 20, StrFileName)

    lngRet = GetPrivateProfileString("PARAMETER", "CTAlertGateWeight", "100", strDataCTGateWeight(0), 20, StrFileName)
    'Special
    lngRet = GetPrivateProfileString("PARAMETER", "ResetIntegral", "0", strDataSpecial(0), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "LampMonitor", "0", strDataSpecial(1), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "MonitorTC", "0", strDataSpecial(2), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "AlarmBuzzer", "0", strDataSpecial(3), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "MultiLoop", "0", strDataSpecial(4), 20, StrFileName)
    
    lngRet = GetPrivateProfileString("PARAMETER", "RampSmooth", "0", strDataHeatingSmooth(0), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "SmoothDisplay", "0", strDataHeatingSmooth(1), 20, StrFileName)
    'lngRet = GetPrivateProfileString("PARAMETER", "SmoothTime", "0", strDataHeatingSmooth(2), 20, strFileName)
                
                
    '<- Pyrometer Coefficient setting
    lngRet = GetPrivateProfileString("CALIBRATION", "PMWAFERCNT", "5", strCoffData(0), 30, StrFileName)
    lngRet = GetPrivateProfileString("CALIBRATION", "PMSPTCNT", "5", strCoffData(1), 30, StrFileName)
    gbintCompensationCountWAF = CInt(Val(strCoffData(0)))
    gbintCompensationCountSPT = CInt(Val(strCoffData(1)))
    ReDim gbdblPMCoffForWafer(0 To gbintCompensationCountWAF)
    gbdblPMCoffForWafer(0) = 0
    For i = 0 To gbintCompensationCountWAF - 1
        lngRet = GetPrivateProfileString("CALIBRATION1", CStr(i), "-999", strCoffData(i), 30, StrFileName)
        If (Val(strCoffData(i)) = -999) Then Exit For
        gbdblPMCoffForWafer(i) = Val(strCoffData(i))
    Next i
    ReDim gbdblPMCoffForSPT(0 To gbintCompensationCountSPT)
    gbdblPMCoffForSPT(0) = 0
    For i = 0 To gbintCompensationCountSPT - 1
        lngRet = GetPrivateProfileString("CALIBRATION2", CStr(i), "-999", strCoffData(i), 30, StrFileName)
        If (Val(strCoffData(i)) = -999) Then Exit For
        gbdblPMCoffForSPT(i) = Val(strCoffData(i))
    Next i
    
    'Rev8.0.1.7
    lngRet = GetPrivateProfileString("PARAMETER", "MonitorTC1Active", "0", strDataMonitorTCActive(0), 30, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "MonitorTC2Active", "0", strDataMonitorTCActive(1), 30, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "MonitorTC3Active", "0", strDataMonitorTCActive(2), 30, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "MonitorTC4Active", "0", strDataMonitorTCActive(3), 30, StrFileName)
    '120713 Josh
    lngRet = GetPrivateProfileString("PARAMETER", "MonitorTC5Active", "0", strDataMonitorTCActive(4), 30, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "MonitorTC6Active", "0", strDataMonitorTCActive(5), 30, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "MonitorTC7Active", "0", strDataMonitorTCActive(6), 30, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "MonitorTC8Active", "0", strDataMonitorTCActive(7), 30, StrFileName)
    StopTCM = CommnonReadini("FuncSwitch", "StopTCM", App.Path + Function_Path)
    ckeStopTCM.value = CInt(StopTCM)
    DefineprocStep = CommnonReadini("FuncSwitch", "ProcStepInUse", App.Path + Function_Path)
    ChkDefineProcStep.value = CInt(DefineprocStep)
    OffsetWriteToTcm = CommnonReadini("FuncSwitch", "OffsetWriteToTcm", App.Path + Function_Path)
    ChkWriteOffsetToTCM.value = CInt(OffsetWriteToTcm)
    
    CTNumbers = CommnonReadini("CTConfig", "CTNumbers", App.Path + DeviceInfo_Path)
    txtCTNumber.text = CTNumbers
    ForcePreheat = CommnonReadini("FuncSwitch", "ForcePreheat", App.Path + Function_Path)
    ckeForcePreheat.value = CInt(ForcePreheat)
    CTDisplay = CommnonReadini("FuncSwitch", "CTDisplay", App.Path + Function_Path)
    ckeCTDisplay.value = CInt(CTDisplay)
    CTNumber1 = CommnonReadini("CTConfig", "CTNumber1", App.Path + DeviceInfo_Path)
    txtNumber1.text = CTNumber1
    CTNumber2 = CommnonReadini("CTConfig", "CTNumber2", App.Path + DeviceInfo_Path)
    txtNumber2.text = CTNumber2
    CTNumber3 = CommnonReadini("CTConfig", "CTNumber3", App.Path + DeviceInfo_Path)
    txtNumber3.text = CTNumber3
    CTNumber4 = CommnonReadini("CTConfig", "CTNumber4", App.Path + DeviceInfo_Path)
    txtNumber4.text = CTNumber4
    
    CTName1 = CommnonReadini("CTConfig", "CTName1", App.Path + DeviceInfo_Path)
    cmbCTName1(0).text = CTName1
    CTName2 = CommnonReadini("CTConfig", "CTName2", App.Path + DeviceInfo_Path)
    cmbCTName2(0).text = CTName2
    CTName3 = CommnonReadini("CTConfig", "CTName3", App.Path + DeviceInfo_Path)
    cmbCTName3(0).text = CTName3
    CTName4 = CommnonReadini("CTConfig", "CTName4", App.Path + DeviceInfo_Path)
    cmbCTName4(0).text = CTName4
    
    CTOrder1 = CommnonReadini("CTConfig", "CTOrder1", App.Path + DeviceInfo_Path)
    cmbOrder1(0).text = CTOrder1
    CTOrder2 = CommnonReadini("CTConfig", "CTOrder2", App.Path + DeviceInfo_Path)
    cmbOrder2(0).text = CTOrder2
    CTOrder3 = CommnonReadini("CTConfig", "CTOrder3", App.Path + DeviceInfo_Path)
    cmbOrder3(0).text = CTOrder3
    CTOrder4 = CommnonReadini("CTConfig", "CTOrder4", App.Path + DeviceInfo_Path)
    cmbOrder4(0).text = CTOrder4
    
    IsUsedSCR = CommnonReadini("Device2", "IsUsed", App.Path + ModbusRtu_Path)
    ckSCREable.value = CInt(IsUsedSCR)
    PortSCR = CommnonReadini("Device2", "Port", App.Path + ModbusRtu_Path)
    txtCommSCR.text = Replace(PortSCR, "COM", "")
    
    lngRet = GetPrivateProfileString("PARAMETER", "TCTYPE", "1", strDataTCType, 30, StrFileName)
    gbintTCType = CInt(Val(strDataTCType))
    cmbTCType.ListIndex = gbintTCType
    lngRet = GetPrivateProfileString("PARAMETER", "TCVoltageRange", "1", strDataTCVoltageRange, 30, StrFileName)
    gbintTCVoltageRange = CInt(Val(strDataTCVoltageRange))
    cmbTCVoltageRange.ListIndex = gbintTCVoltageRange
    '->
    'Normal parameter
    For i = 0 To txtParaNormal.UBound
        txtParaNormal(i).text = strDataNormal(i)
    Next i
    chkUniformity.value = CInt(Val(strDataNormal(9)))
    If chkUniformity.value = 1 Then
        chkUniformity.Caption = "Enable"
    Else
        chkUniformity.Caption = "Disable"
    End If
    
    'Heating parameter
    For i = 0 To txtParaHeat.UBound
        txtParaHeat(i).text = strDataHeating(i)
    Next i
    'Gas parameter
    gbintMaxGasEnable = -1
    For i = 0 To txtParaGasAlias.UBound
        txtParaGasAlias(i).text = strDataGas(i)
        txtParaGasValue(i).text = strDataGas(i + 7)
        txtParaGasBias(i).text = strDataGas(i + 14)
        chkGasEnable(i).value = CInt(Val(strDataGas(i + 21)))
        txtParaGasUnit(i).text = strDataGas(i + 28)
        txtParaGasError(i).text = strDataGas(i + 35)
        txtParaGasErrorN(i).text = strDataGas(i + 42)
        If (chkGasEnable(i).value = 1) Then
            gbintMaxGasEnable = gbintMaxGasEnable + 1
        End If
    Next i
    
    'Vacuum parameter
    For i = 0 To txtParaVacuum.UBound
        txtParaVacuum(i).text = strDataVacuum(i)
    Next i
    'CT reference parameter
    For i = 0 To txtIntensityRef.UBound
        txtIntensityRef(i).text = strDataIntensityRef(i)
    Next i
    'CT gate parameter
    For i = 0 To txtParaCTGate.UBound
        txtParaCTGate(i).text = strDataCTGate(i)
    Next i
    txtCTAlertGateWeight.text = strDataCTGateWeight(0)
    
    gbsngIntensityRef(0) = Val(txtIntensityRef(0).text)
    gbsngIntensityRef(1) = Val(txtIntensityRef(1).text)
    For i = 0 To 6
        gbsngCTGate1(i) = txtParaCTGate(i)
        gbsngCTGate2(i) = txtParaCTGate(i + 7)
    Next i
    gbsngCTGateWeight = Val(txtCTAlertGateWeight.text)
    


    'Special parameter
    chkResetIntegral.value = CInt(Val(strDataSpecial(0)))
    chkLampMonitor.value = CInt(Val(strDataSpecial(1)))
    chkMonitorTC.value = CInt(Trim(strDataSpecial(2)))
    chkAlarmBuzzer.value = CInt(Trim(strDataSpecial(3)))
    'chkMultiLoop.value = CInt(Trim(strDataSpecial(4)))
    'Rev4.1.2
    cmbControlMode.ListIndex = CInt(Val(strDataSpecial(4)))
    'Heating smooth parameter
    chkSmoothRamp.value = CInt(Val(strDataHeatingSmooth(0)))
    chkSmoothDisplay.value = CInt(Val(strDataHeatingSmooth(1)))
    'txtSmoothTime.Text = CStr(CSng(Val(strDataHeatingSmooth(2))))
    ckeStopTCM.value = CInt(StopTCM)
    ChkDefineProcStep.value = CInt(DefineprocStep)
    ChkWriteOffsetToTCM.value = CInt(OffsetWriteToTcm)
    'Rev8.0.1.7
    chkMonitorTCActive(0).value = CInt(Trim(strDataMonitorTCActive(0)))
    chkMonitorTCActive(1).value = CInt(Trim(strDataMonitorTCActive(1)))
    chkMonitorTCActive(2).value = CInt(Trim(strDataMonitorTCActive(2)))
    chkMonitorTCActive(3).value = CInt(Trim(strDataMonitorTCActive(3)))
    '120713 Josh
    chkMonitorTCActive(4).value = CInt(Trim(strDataMonitorTCActive(4)))
    chkMonitorTCActive(5).value = CInt(Trim(strDataMonitorTCActive(5)))
    chkMonitorTCActive(6).value = CInt(Trim(strDataMonitorTCActive(6)))
    chkMonitorTCActive(7).value = CInt(Trim(strDataMonitorTCActive(7)))
    
    gbintMonitorTCActive(0) = CInt(chkMonitorTCActive(0).value)
    gbintMonitorTCActive(1) = CInt(chkMonitorTCActive(1).value)
    gbintMonitorTCActive(2) = CInt(chkMonitorTCActive(2).value)
    gbintMonitorTCActive(3) = CInt(chkMonitorTCActive(3).value)
    '120713 Josh
    gbintMonitorTCActive(4) = CInt(chkMonitorTCActive(4).value)
    gbintMonitorTCActive(5) = CInt(chkMonitorTCActive(5).value)
    gbintMonitorTCActive(6) = CInt(chkMonitorTCActive(6).value)
    gbintMonitorTCActive(7) = CInt(chkMonitorTCActive(7).value)
           
    'Assign parameter to variable
    gbsngMaxTemperature = Val(txtParaNormal(0).text)
    If (gbsngMaxTemperature > 1500) Then gbsngMaxTemperature = 1500
    gbintPreheatIntensity = CSng(txtParaNormal(1).text)
    gbintPumpTimeout = CInt(txtParaNormal(2).text)
'    If gbintPumpTimeout > 65 Then
'        gbintPumpTimeout = 65
'        txtParaNormal(2).Text = "65"
'    End If
    gbsngPumpDownGate = Val(txtParaNormal(3).text)
    gbsngAPCGaugePressureValue = Val(txtParaNormal(4).text)
    gbsngPumpingDelay = Val(txtParaNormal(5).text)
    gbsngTCDifferentialRange = Val(txtParaNormal(6).text)
    
    If (Val(txtParaNormal(7).text) > 500) Then txtParaNormal(7).text = "500"
    If (Val(txtParaNormal(7).text) < 100) Then txtParaNormal(7).text = "100"
    
    gbsngVentGate = Val(txtParaNormal(7).text)
    
    gbsngIntensityLimit = Val(txtParaNormal(8).text)
        
    gbsngChamberOverheat = Val(txtParaHeat(0).text)
    gbsngValidPMTempature = Val(txtParaHeat(1).text)
    gbsngPMCVT1 = Val(txtParaHeat(2).text)
    gbsngPMCVT2 = Val(txtParaHeat(3).text)
    gbdblProcessPreheatTimerout = CDbl(txtParaHeat(4).text)
    gbsngIntensityKeep = CSng(Val(txtParaHeat(5).text) / 10)
    gbsngTCCVT1 = Val(txtParaHeat(6).text)
    gbsngTCCVT2 = Val(txtParaHeat(7).text)
    gbintNumOfBanks = Val(txtParaHeat(8).text)
    
    'Rev9.0.0.0
    
    For j = 0 To txtParaGasAlias.UBound
        gbintGasEnable(j) = chkGasEnable(j).value
        gbstrGasAlias(j) = txtParaGasAlias(j).text
        gbstrGasUnit(j) = txtParaGasUnit(j).text
        gbsngMaxGasSLMP(j) = Val(txtParaGasValue(j).text)
        gbsngGasBias(j) = Val(txtParaGasBias(j).text)
        gbsngGasError(j) = Val(txtParaGasError(j).text)
        gbsngGasErrorN(j) = Val(txtParaGasErrorN(j).text)
    Next j
    
    gbsngVacuumGaugeCompensation = Val(txtParaVacuum(0).text)
    gbintAngleOpenDelay = CInt(txtParaVacuum(1).text)
    gbintReleaseOpenDelay = CInt(txtParaVacuum(2).text)
    gbintThrottleFullOpenDelay = CInt(txtParaVacuum(3).text)
    gbintThrottleInitialPos = CInt(txtParaVacuum(4).text)
    gbsngAPCGaugeValveLimit = Val(txtParaVacuum(5).text)
    gbsngGaugeZoomIn = Val(txtParaVacuum(6).text)
    gbsngAPCInterval = Val(txtParaVacuum(7).text)
    
    gbsngKeepPurge = Val(txtParaVacuum(9).text)
    gbintAPC_MFC_Port = Val(txtParaVacuum(11).text)
    If gbintAPC_MFC_Port <= 0 Or gbintAPC_MFC_Port > 6 Then gbintAPC_MFC_Port = 1
    gbintMFC_Ratio = Val(txtParaVacuum(12).text)
    If gbintMFC_Ratio <= 0 Then gbintMFC_Ratio = 1
            
    gbdblProcessPumpDownTimerout = CDbl(CSng(gbintPumpTimeout) * 1000)
    
    gbblnResetInteral = IIf(chkResetIntegral.value = 1, True, False)
    gbintLampMonitor = CInt(chkLampMonitor.value)
    gbintAlarmBuzzer = CInt(chkAlarmBuzzer.value)
    gbintMonitorTC = CInt(chkMonitorTC.value)
    
    gbintRampSmooth = CInt(chkSmoothRamp.value)
    'gbsngSmoothTime = CSng(Val(txtSmoothTime.Text))
    gbintSmoothDisplay = CInt(chkSmoothDisplay.value)
        
    'Rev4.1.5
    lngRet = GetPrivateProfileString("PARAMETER", "UniRampActive", "0", strDataUniformity(0), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "UniStartPointRamp", "0", strDataUniformity(1), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "UniSubWeightD1", "0", strDataUniformity(2), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "UniSubWeightD2", "0", strDataUniformity(3), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "UniStartPointHold", "0", strDataUniformity(4), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "UniSubWeight1", "0", strDataUniformity(5), 20, StrFileName)
    lngRet = GetPrivateProfileString("PARAMETER", "UniSubWeight2", "0", strDataUniformity(6), 20, StrFileName)
    
    txtUniformityRampStartPoint.text = strDataUniformity(1)
    txtSubWeightD1.text = strDataUniformity(2)
    txtSubWeightD2.text = strDataUniformity(3)
    txtUniformityHoldStartPoint.text = strDataUniformity(4)
    txtSubWeight1.text = strDataUniformity(5)
    txtSubWeight2.text = strDataUniformity(6)
    
    
    gbsngUniformitySubWeightD1 = CSng(Val(txtSubWeightD1.text))
    gbsngUniformitySubWeightD2 = CSng(Val(txtSubWeightD2.text))
    gbsngUniformityStartPointHold = CSng(Val(txtUniformityHoldStartPoint.text))
    gbsngUniformitySubWeight1 = CSng(Val(txtSubWeight1.text))
    gbsngUniformitySubWeight2 = CSng(Val(txtSubWeight2.text))
    
        
    'System alarm notice enable / disable
    lngRet = GetPrivateProfileString("ALARM_ACTIVE", "System", "0", strAlarmEnable(0), 20, StrFileName)
    lngRet = GetPrivateProfileString("ALARM_ACTIVE", "Ready", "0", strAlarmEnable(1), 20, StrFileName)
    lngRet = GetPrivateProfileString("ALARM_ACTIVE", "EMS", "0", strAlarmEnable(2), 20, StrFileName)
    lngRet = GetPrivateProfileString("ALARM_ACTIVE", "BZR", "0", strAlarmEnable(3), 20, StrFileName)
    lngRet = GetPrivateProfileString("ALARM_ACTIVE", "OH", "0", strAlarmEnable(4), 20, StrFileName)
    lngRet = GetPrivateProfileString("ALARM_ACTIVE", "CT", "0", strAlarmEnable(5), 20, StrFileName)
    lngRet = GetPrivateProfileString("ALARM_ACTIVE", "TC", "0", strAlarmEnable(6), 20, StrFileName)
    lngRet = GetPrivateProfileString("ALARM_ACTIVE", "Water", "0", strAlarmEnable(7), 20, StrFileName)
    lngRet = GetPrivateProfileString("ALARM_ACTIVE", "CDA", "0", strAlarmEnable(8), 20, StrFileName)
    lngRet = GetPrivateProfileString("ALARM_ACTIVE", "VG", "0", strAlarmEnable(9), 20, StrFileName)
    lngRet = GetPrivateProfileString("ALARM_ACTIVE", "VS", "0", strAlarmEnable(10), 20, StrFileName)
    lngRet = GetPrivateProfileString("ALARM_ACTIVE", "Door", "0", strAlarmEnable(11), 20, StrFileName)
    lngRet = GetPrivateProfileString("ALARM_ACTIVE", "APC", "0", strAlarmEnable(12), 20, StrFileName)
    lngRet = GetPrivateProfileString("ALARM_ACTIVE", "Chamber", "0", strAlarmEnable(13), 20, StrFileName)

    For i = 0 To chkAlarmEnable.UBound
        chkAlarmEnable(i).value = CInt(Val(strAlarmEnable(i)))
    Next i
    
    chkAlarmEnable(14).value = CommnonReadini("FuncSwitch", "CheckTcWafer", App.Path + Function_Path)
    gbintActiveAlarm_TcWafer = chkAlarmEnable(14).value
    
    chkAlarmEnable(15).value = CommnonReadini("FuncSwitch", "CheckRValue", App.Path + Function_Path)
    gbintActiveAlarm_RValue = chkAlarmEnable(15).value
    
    Txt_RValueRange.text = CommnonReadini("Loop1-5", "RValRanage", App.Path + ProcDict_Path)
    gbRValRange = Txt_RValueRange.text
    '--------------------------------------------------------------------
    'Assign Alarm Activity Enable/Disable to variable
    '--------------------------------------------------------------------
    '***System status***
    gbintActiveAlarm_System = chkAlarmEnable(0).value
    gbintActiveAlarm_Ready = chkAlarmEnable(1).value
    'gbintActiveAlarm_DC = chkAlarmEnable(1).Value
    gbintActiveAlarm_EMS = chkAlarmEnable(2).value
    gbintActiveAlarm_Buzzer = chkAlarmEnable(3).value
    
    '***The Heating Module***
    gbintActiveAlarm_Overheat = chkAlarmEnable(4).value
    gbintActiveAlarm_CT = chkAlarmEnable(5).value
    gbintActiveAlarm_TC = chkAlarmEnable(6).value
    
    '***The Gas Module***
    '***The Cooling Module***
    gbintActiveAlarm_Water = chkAlarmEnable(7).value
    gbintActiveAlarm_Air = chkAlarmEnable(8).value
    '***The Vacuum Module***
    gbintActiveAlarm_VacuumGauge = chkAlarmEnable(9).value
    gbintActiveAlarm_VacuumSwitch = chkAlarmEnable(10).value
    gbintActiveAlarm_Door = chkAlarmEnable(11).value
    gbintActiveAlarm_APC = chkAlarmEnable(12).value
    gbintActiveAlarm_ChamberGate = chkAlarmEnable(13).value
    'Module enable / disable
    lngRet = GetPrivateProfileString("MODULE_ACTIVE", "H", "0", strModuleEnable(0), 20, StrFileName)
    lngRet = GetPrivateProfileString("MODULE_ACTIVE", "C", "0", strModuleEnable(1), 20, StrFileName)
    lngRet = GetPrivateProfileString("MODULE_ACTIVE", "G", "0", strModuleEnable(2), 20, StrFileName)
    lngRet = GetPrivateProfileString("MODULE_ACTIVE", "V", "0", strModuleEnable(3), 20, StrFileName)
    lngRet = GetPrivateProfileString("MODULE_ACTIVE", "F", "0", strModuleEnable(4), 20, StrFileName)
    lngRet = GetPrivateProfileString("MODULE_ACTIVE", "O", "0", strModuleEnable(5), 20, StrFileName)
    lngRet = GetPrivateProfileString("MODULE_ACTIVE", "B", "0", strModuleEnable(6), 20, StrFileName)
    lngRet = GetPrivateProfileString("MODULE_ACTIVE", "D", "0", strModuleEnable(7), 20, StrFileName)
    lngRet = GetPrivateProfileString("MODULE_ACTIVE", "P", "0", strModuleEnable(8), 20, StrFileName)
    lngRet = GetPrivateProfileString("MODULE_ACTIVE", "A", "0", strModuleEnable(9), 20, StrFileName)
    lngRet = GetPrivateProfileString("MODULE_ACTIVE", "R", "0", strModuleEnable(10), 20, StrFileName)
    lngRet = GetPrivateProfileString("MODULE_ACTIVE", "T", "0", strModuleEnable(11), 20, StrFileName)
    lngRet = GetPrivateProfileString("MODULE_ACTIVE", "S", "0", strModuleEnable(13), 20, StrFileName)
    lngRet = GetPrivateProfileString("MODULE_ACTIVE", "M", "0", strModuleEnable(14), 20, StrFileName)
    For i = 0 To chkModuleEnable.UBound
        chkModuleEnable(i).value = CInt(Val(strModuleEnable(i)))
    Next i
    '--------------------------------------------------------------------
    'Assign Module Activity Enable/Disable
    '--------------------------------------------------------------------
    gbintActiveModule_Heating = chkModuleEnable(0).value
    gbintActiveModule_Cooling = chkModuleEnable(1).value
    gbintActiveModule_Gas = chkModuleEnable(2).value
    gbintActiveModule_Vacuum = chkModuleEnable(3).value
    gbintActiveModule_Database = chkModuleEnable(4).value
    gbintActiveModule_Oxygen = chkModuleEnable(5).value
    gbintActiveModule_Barcode = chkModuleEnable(6).value
    gbintActiveModule_Door = chkModuleEnable(7).value
    gbintActiveModule_PNRecipe = chkModuleEnable(8).value
    gbintActiveModule_APC = chkModuleEnable(9).value
    gbintActiveModule_Auto = chkModuleEnable(10).value
    gbintActiveModule_CIM = chkModuleEnable(13).value
    gbintActiveModule_MLoop = chkModuleEnable(14).value
    
    If gbdblPMCoffForWafer(0) <> 0 Then
        gbblnCompensationPM = True
    End If
    
    For i = 0 To 31
        lngRet = GetPrivateProfileString("IO91141", "DI" & CStr(i), "NA", StrData(i), 20, StrFileName)
        fgDI.TextMatrix(i + 1, 1) = StrData(i)
    Next i
    For i = 0 To 31
        lngRet = GetPrivateProfileString("IO91141", "DO" & CStr(i), "NA", StrData(i), 20, StrFileName)
        fgDO.TextMatrix(i + 1, 1) = StrData(i)
    Next i
    For i = 0 To 31
        lngRet = GetPrivateProfileString("IO91141", "AI" & CStr(i), "NA", StrData(i), 20, StrFileName)
        fgAI.TextMatrix(i + 1, 1) = StrData(i)
        lngRet = GetPrivateProfileString("IO91141", "AI_ErrorV" & CStr(i), "0", StrData(i), 20, StrFileName)
        fgAI.TextMatrix(i + 1, 3) = StrData(i)
        SysAI.ErrorV(i) = Val(StrData(i))
    Next i
    For i = 0 To 31
        lngRet = GetPrivateProfileString("IO62081", "AO" & CStr(i), "NA", StrData(i), 20, StrFileName)
        fgAO.TextMatrix(i + 1, 1) = StrData(i)
        lngRet = GetPrivateProfileString("IO62081", "AO_Type" & CStr(i), "0", StrData(i), 20, StrFileName)
        fgAO.TextMatrix(i + 1, 2) = StrData(i)
        gbintAO_Type(i) = CInt(StrData(i))
    Next i
    
    iniPara.Path = gbSystemPath & "\System\system.cfg"
    iniPara.Section = "ADVIO17101"
    gbintTCO2 = -1
    For i = 0 To 23
'        lngRet = GetPrivateProfileString("ADVIO17101", "AI" & CStr(i), "NA", StrTC(i), 5, strFileName)
'        fgAdvDaqAI.TextMatrix(i + 1, 1) = StrTC(i)
'        gbstrNameTC(i) = StrTC(i)
        iniPara.Key = "AI" & CStr(i)
        gbstrNameTC(i) = iniPara.value
        If gbstrNameTC(i) = "O2" Then
            gbintTCO2 = i
        End If
        fgAdvDaqAI.TextMatrix(i + 1, 1) = gbstrNameTC(i)
        lngRet = GetPrivateProfileString("ADVIO17101", "Power" & CStr(i), "0", StrData(i), 20, StrFileName)
        fgAdvDaqAI.TextMatrix(i + 1, 2) = StrData(i)
        gbsngPowerTC(i) = CSng(StrData(i))
        lngRet = GetPrivateProfileString("ADVIO17101", "Error" & CStr(i), "0", StrData(i), 20, StrFileName)
        fgAdvDaqAI.TextMatrix(i + 1, 3) = StrData(i)
        gbsngErrorTC(i) = CSng(StrData(i))
        lngRet = GetPrivateProfileString("ADVIO17101", "Ratio" & CStr(i), "1", StrData(i), 20, StrFileName)
        fgAdvDaqAI.TextMatrix(i + 1, 4) = StrData(i)
        gbsngRatioTC(i) = CSng(StrData(i))
        lngRet = GetPrivateProfileString("ADVIO17101", "Power3C" & CStr(i), "0", StrData(i), 20, StrFileName)
        fgAdvDaqAI.TextMatrix(i + 1, 6) = StrData(i)
        gbsngPower3C(i) = CSng(StrData(i))
        lngRet = GetPrivateProfileString("ADVIO17101", "Power4C" & CStr(i), "0", StrData(i), 20, StrFileName)
        fgAdvDaqAI.TextMatrix(i + 1, 7) = StrData(i)
        gbsngPower4C(i) = CSng(StrData(i))
        lngRet = GetPrivateProfileString("ADVIO17101", "Power5C" & CStr(i), "0", StrData(i), 20, StrFileName)
        fgAdvDaqAI.TextMatrix(i + 1, 8) = StrData(i)
        gbsngPower5C(i) = CSng(StrData(i))
        lngRet = GetPrivateProfileString("ADVIO17101", "LoopNo" & CStr(i), "0", StrData(i), 20, StrFileName)
        fgAdvDaqAI.TextMatrix(i + 1, 9) = StrData(i)
        gbintLoopNo(i) = CInt(StrData(i))
        Dim DigitValue As String
        DigitValue = CommnonReadini("PrecisionDigit", "PrecisionDigit" & CStr(i), App.Path + ProcDict_Path)
'        lngRet = GetPrivateProfileString("ADVIO17101", "PrecisionDigit" & CStr(i), "1", StrData(i), 20, strFileName)
        If DigitValue = "0" Then
        DigitValue = "1"
        End If
        fgAdvDaqAI.TextMatrix(i + 1, 10) = DigitValue
        gbintPrecisionDigit(i) = CInt(DigitValue)
    Next i
     
    
    lngRet = GetPrivateProfileString("PARAMETER", "FinishedBeep", "0", StrData(0), 20, StrFileName)
    txtParaNormal(9).text = StrData(0)
    gbintFinishedBeep = Val(txtParaNormal(9).text)
    lngRet = GetPrivateProfileString("PARAMETER", "LifeLamp", "0", StrData(0), 20, StrFileName)
    txtParaNormal(11).text = CStr(Round(Val(StrData(0)) / 3600, 3))
    gbsngLifeLamp = Val(StrData(0))
    lngRet = GetPrivateProfileString("PARAMETER", "UsedLamp", "0", StrData(0), 20, StrFileName)
    txtParaNormal(12).text = CStr(Round(Val(StrData(0)) / 3600, 3))
    gbsngUsedLamp = Val(StrData(0))
    lngRet = GetPrivateProfileString("PARAMETER", "LogFilePath", "C:\Program Files\eRTA100", strPath, 100, StrFileName)
    txtParaNormal(13).text = strPath
    gbstrLogFilePath = txtParaNormal(13).text
    lngRet = GetPrivateProfileString("PARAMETER", "RecipeFilePath", "C:\Program Files\eRTA100\Recipe\op\", strPath, 100, StrFileName)
    txtParaNormal(20).text = strPath
    gbstrRecipeFilePath = txtParaNormal(20).text
    
    lngRet = GetPrivateProfileString("PARAMETER", "MaxMonitorError", "0", StrData(0), 20, StrFileName)
    txtParaNormal(14).text = StrData(0)
    gbsngMaxMonitorError = Val(txtParaNormal(14).text)
    lngRet = GetPrivateProfileString("PARAMETER", "MaxMonitorTime", "0", StrData(0), 20, StrFileName)
    txtParaNormal(15).text = StrData(0)
    gbsngMaxMonitorTime = Val(txtParaNormal(15).text)
    lngRet = GetPrivateProfileString("PARAMETER", "FinishedLight", "0", StrData(0), 20, StrFileName)
    txtParaNormal(16).text = StrData(0)
    gbintFinishedLight = Val(txtParaNormal(16).text)
    lngRet = GetPrivateProfileString("PARAMETER", "CycleRun", "0", StrData(0), 20, StrFileName)
    txtParaNormal(17).text = StrData(0)
    gbintCycleRun = Val(txtParaNormal(17).text)
    lngRet = GetPrivateProfileString("PARAMETER", "GaugeValue", "0", StrData(0), 20, StrFileName)
    txtParaNormal(18).text = StrData(0)
    gbsngGaugeValue = Val(txtParaNormal(18).text)
    lngRet = GetPrivateProfileString("PARAMETER", "IdleWarning", "0", StrData(0), 20, StrFileName)
    txtParaNormal(19).text = StrData(0)
    gbsngIdleWarning = Val(txtParaNormal(19).text) * 60
    
    lngRet = GetPrivateProfileString("PARAMETER", "MinTemperature", "0", StrData(0), 20, StrFileName)
    txtParaNormal(10).text = StrData(0)
    gbsngMinTemperature = Val(txtParaNormal(10).text)
    
    lngRet = GetPrivateProfileString("PARAMETER", "OpenTemperature", "0", StrData(0), 20, StrFileName)
    txtParaNormal(28).text = StrData(0)
    gbsngOpenTemperature = Val(txtParaNormal(28).text)
    
    lngRet = GetPrivateProfileString("Utility", "CTCheck", "1", StrData(0), 20, StrFileName)
    chkCTCheck.value = CInt(Val(StrData(0)))
    gbintCTCheck = chkCTCheck.value
    lngRet = GetPrivateProfileString("Utility", "RtaType", "1", StrData(0), 20, StrFileName)
    txtRtaType.text = StrData(0)
    gbintRtaType = Val(txtRtaType.text)
    lngRet = GetPrivateProfileString("Utility", "AutoDeleteRecipe", "0", StrData(0), 20, StrFileName)
    chkAutoDeleteRecipe.value = CInt(Val(StrData(0)))
    gbintAutoDeleteRecipe = chkAutoDeleteRecipe.value
    
    
    lngRet = GetPrivateProfileString("InterLock", "TempDownTimeout", "0", StrData(0), 20, StrFileName)
    gbsngRecipeTempDownTimeout = Val(StrData(0))
    txtTempDownTimeout.text = CStr(Val(StrData(0)))
    
    lngRet = GetPrivateProfileString("InterLock", "GatePS1", "0", StrData(0), 20, StrFileName)
    gbsngRecipeGatePS1 = Val(StrData(0))
    txtGatePS1.text = CStr(Val(StrData(0)))
    lngRet = GetPrivateProfileString("InterLock", "GatePS2", "0", StrData(0), 20, StrFileName)
    gbsngRecipeGatePS2 = Val(StrData(0))
    txtGatePS2.text = CStr(Val(StrData(0)))
    
    
    lngRet = GetPrivateProfileString("Robot", "RobotPort", "1", StrData(0), 20, StrFileName)
    txtRobotPort.text = StrData(0)
    gbintRobotPort = Val(txtRobotPort.text)
    lngRet = GetPrivateProfileString("Robot", "RobotPickH", "1", StrData(0), 20, StrFileName)
    txtPickH.text = StrData(0)
    gbsngPickH = Val(txtPickH.text)
    lngRet = GetPrivateProfileString("Robot", "RobotPlaceH", "1", StrData(0), 20, StrFileName)
    txtPlaceH.text = StrData(0)
    gbsngPlaceH = Val(txtPlaceH.text)
    lngRet = GetPrivateProfileString("Robot", "RobotSpeed", "0", StrData(0), 20, StrFileName)
    gbintRobotSpeed = Val(StrData(0))
        
    For i = 0 To 3
        If gbintRobotSpeed = i Then optRobotSpeed(i).value = True
    Next i
    
    
    For i = 0 To 49
        For j = 0 To 3
            lngRet = GetPrivateProfileString("Robot", "Teach" & CStr(i) & "_" & CStr(j), "0", StrData(0), 20, StrFileName)
            gbsngTeach(i, j) = Val(StrData(0))
            fgTeach.TextMatrix(i + 1, j + 1) = StrData(0)
            
            
        Next j
    Next i
    
    
    '120822 Josh
    lngRet = GetPrivateProfileString("User", "ad", "123", StrData(0), 20, StrFileName)
    gbstrAdminPW = Mid(StrData(0), 1, 3)
    txtAdmin.text = gbstrAdminPW
    lngRet = GetPrivateProfileString("User", "eg", "456", StrData(0), 20, StrFileName)
    gbstrEngineerPW = Mid(StrData(0), 1, 3)
    txtEngineer.text = gbstrEngineerPW
    lngRet = GetPrivateProfileString("User", "op", "789", StrData(0), 20, StrFileName)
    gbstrOperatorPW = Mid(StrData(0), 1, 3)
    txtOperator.text = gbstrOperatorPW
    
    lngRet = GetPrivateProfileString("User", "Page0", "0", strPage(0), 20, StrFileName)
    gbintActivePage(0) = Val(strPage(0))
    txtActivePage(0).text = gbintActivePage(0)
    lngRet = GetPrivateProfileString("User", "Page1", "0", strPage(1), 20, StrFileName)
    gbintActivePage(1) = Val(strPage(1))
    txtActivePage(1).text = gbintActivePage(1)
    lngRet = GetPrivateProfileString("User", "Page2", "0", strPage(2), 20, StrFileName)
    gbintActivePage(2) = Val(strPage(2))
    txtActivePage(2).text = gbintActivePage(2)
    
    For i = 0 To 23
        gbsngRatioEX(i) = 1
    Next i
    
'    For i = 0 To GB_MAX_LOOPS - 1
'        lngRet = GetPrivateProfileString("PARAMETER", "RatioEX" & CStr(i), "1", StrData(0), 20, strFileName)
'        txtRatioEX(i).text = StrData(0)
'        gbsngRatioEX(i) = Val(txtRatioEX(i).text)
'    Next i
              
    LoadPara
    
    chkOnlyRecipe.value = Para.intOnlyRecipe
    txtRtaType.text = Para.RtaType
    chkBarcodeServer.value = Para.UseBarcodeServer
    txtParaNormal(17).text = CStr(Para.intCycleRuns)
    txtServerPath.text = Para.strServerPath
    txtMonitorRuns.text = CStr(Para.intMonitorRuns)
    txtTestRunKey.text = Para.strTestRunKey
    chkHoldSafety.value = Para.IsHoldSafety
    chkCalibration.value = Para.IsCali
    
    'chkModuleEnable(10).value = Para.UseAutoMode
    chkAutoMode.value = Para.UseAutoMode
    
    chkModuleEnable(11).value = Para.UseCT
    txtComCT.text = Para.intComCT
    chkModuleEnable(12).value = Para.UseMTC
    chkModuleEnable(19).value = Para.UseMTCB
    chkCIM.value = Para.UseCIM
    txtAzIP1.text = Para.strAzIP1
    txtAzIP2.text = Para.strAzIP2
    chkModuleEnable(15).value = Para.UseAz1
    chkModuleEnable(16).value = Para.UseAz2
    chkModuleEnable(17).value = Para.useTPump
    txtRobotIP.text = Para.strRobotIP
    
    txtParaNormal(21).text = Para.sngGaugeAngle
    txtParaNormal(22).text = Para.intMonitorIndex
    txtParaNormal(23).text = Para.sngO2Gate
    txtParaNormal(24).text = Para.intLampAlarmTime
    txtParaNormal(25).text = Para.intOpenDoorTime
    txtParaNormal(26).text = Para.intAutoPort
    txtParaNormal(27).text = Para.intPumpDelay
    cmbCIMPort.ListIndex = Para.intCIMPort
    txtPMbig = Para.intPMbig
    txtPMsmall = Para.intPMsmall
    
    txtGaugeD.text = Para.sngGaugeD
    txtGaugeVP.text = Para.sngGaugeVP
    txtGaugeVN.text = Para.sngGaugeVN
    
    chkModuleEnable(18).value = Para.UseCover
               
    For i = 1 To 33
        fgAlarm.TextMatrix(i, 0) = CStr(4000 + i)
        fgAlarm.TextMatrix(i, 1) = Para.AlarmName(i)
        fgAlarm.TextMatrix(i, 2) = CStr(Para.AlarmDo(i))
    Next i
    
    Call RefreshFunctionRight
    Call InitialIO
    Call SetCTSlope
    Exit Sub
    
ERR_PARAMETER_OPEN:
    Call AlertShow("Open Parameter Failed!!", ERRORTYPE)
End Sub

Private Sub cmbIOList_Click(Index As Integer)
    cmbIOList(Index).Visible = False
    Select Case Index
        Case 0
            fgDI.TextMatrix(fgDI.RowSel, fgDI.ColSel) = cmbIOList(Index).text
            'fgDI.Refresh
        Case 1
            fgDO.TextMatrix(fgDO.RowSel, fgDO.ColSel) = cmbIOList(Index).text
            'fgDO.Refresh
        Case 2
            fgAI.TextMatrix(fgAI.RowSel, fgAI.ColSel) = cmbIOList(Index).text
            'fgAI.Refresh
        Case 3
            fgAO.TextMatrix(fgAO.RowSel, fgAO.ColSel) = cmbIOList(Index).text
            'fgAO.Refresh
        Case 4
            fgAdvDaqAI.TextMatrix(fgAdvDaqAI.RowSel, fgAdvDaqAI.ColSel) = cmbIOList(Index).text
        
    End Select
    
    
    blnSave = True
End Sub

Private Sub cmbIOList_LostFocus(Index As Integer)
    cmbIOList(Index).Visible = False
End Sub

Private Sub cmdLogPath_Click()
    frmPathSetting.Show vbModal
End Sub

Private Sub cmdOpen_Click()
    ParameterOpen
    OpenFunctionIni
    OpenOffsetValue
    
    frmRecipeEdit.RefreshRecipeGridTitle
    frmDiagnosis.RefreshGasDefiniation
End Sub

Private Sub cmdSave_Click()
    'Call txtParaNormal_KeyPress(0, 13)
'    On Error GoTo ERR_SAVE:
    ParameterSave
    SaveFunctionIni
    SaveOffsetValue
'    frmRecipeEdit.RefreshRecipeGridTitle
     Dim OldGasCount As Integer
    OldGasCount = frmRecipeEdit.GasNames.Count
    If NewGasCount <> OldGasCount Then frmRecipeEdit.InitForm
    frmDiagnosis.RefreshGasDefiniation
'    MsgBox "保存成功"
'    Exit Sub
'ERR_SAVE:
'    WriteLog ("保存Configuration失敗")
End Sub

Private Sub fgAdvDaqAI_Click()
    fgAdvDaqAI.Refresh
    If fgAdvDaqAI.col = 1 Then
        With cmbIOList(4)
            .text = cmbIOList(4).text
            .Move tabConfiguration.Left + fraAdvDaqAI.Left + fgAdvDaqAI.Left + fgAdvDaqAI.ColPos(fgAdvDaqAI.ColSel), _
                  tabConfiguration.Top + fraAdvDaqAI.Top + fgAdvDaqAI.Top + fgAdvDaqAI.RowPos(fgAdvDaqAI.RowSel), _
                  fgAdvDaqAI.ColWidth(fgAdvDaqAI.ColSel)
            .Visible = True
            .ZOrder
            .SetFocus
        End With
    End If
    
    If fgAdvDaqAI.col > 1 And fgAdvDaqAI.col < 11 Then
        With txtErrorMap
            .text = fgAdvDaqAI.TextMatrix(fgAdvDaqAI.row, fgAdvDaqAI.col)
            .Move tabConfiguration.Left + fraAdvDaqAI.Left + fgAdvDaqAI.Left + fgAdvDaqAI.ColPos(fgAdvDaqAI.ColSel), _
                  tabConfiguration.Top + fraAdvDaqAI.Top + fgAdvDaqAI.Top + fgAdvDaqAI.RowPos(fgAdvDaqAI.RowSel), _
                  fgAdvDaqAI.ColWidth(fgAdvDaqAI.ColSel)
            .Visible = True
            .ZOrder
            .SetFocus
        End With
    End If
End Sub

Private Sub fgAI_Click()
    fgAI.Refresh
    If fgAI.col = 1 Then
        With cmbIOList(2)
            .text = cmbIOList(2).text
            .Move tabConfiguration.Left + fraAI.Left + fgAI.Left + fgAI.ColPos(fgAI.ColSel), _
                  tabConfiguration.Top + fraAI.Top + fgAI.Top + fgAI.RowPos(fgAI.RowSel), _
                  fgAI.ColWidth(fgAI.ColSel)
            .Visible = True
            .ZOrder
            .SetFocus
        End With
    End If
    
    If fgAI.col = 3 And fgAI.row > 0 Then
        With txtErrorV
            .text = fgAI.TextMatrix(fgAI.row, fgAI.col)
            .Move tabConfiguration.Left + fraAI.Left + fgAI.Left + fgAI.ColPos(fgAI.ColSel), _
                  tabConfiguration.Top + fraAI.Top + fgAI.Top + fgAI.RowPos(fgAI.RowSel), _
                  fgAI.ColWidth(fgAI.ColSel)
            .Visible = True
            .ZOrder
            .SetFocus
        End With
    End If
    
   
End Sub

Private Sub fgAlarm_Click()
    fgAlarm.Refresh
    If fgAlarm.col = 2 Then
    
        With txtAlarm
            .text = fgAlarm.TextMatrix(fgAlarm.row, fgAlarm.col)
            .Move tabConfiguration.Left + tabMain.Left + fgAlarm.Left + fgAlarm.ColPos(fgAlarm.ColSel), _
                  tabConfiguration.Top + tabMain.Top + fgAlarm.Top + fgAlarm.RowPos(fgAlarm.RowSel), _
                  fgAlarm.ColWidth(fgAlarm.ColSel)
            .Visible = True
            .ZOrder
            .SetFocus
        End With
    End If
        
End Sub

Private Sub fgAO_Click()
        fgAO.Refresh
    If fgAO.col = 1 Then
        With cmbIOList(3)
            .Move tabConfiguration.Left + fraAO.Left + fgAO.Left + fgAO.ColPos(fgAO.ColSel), _
                  tabConfiguration.Top + fraAO.Top + fgAO.Top + fgAO.RowPos(fgAO.RowSel), _
                  fgAO.ColWidth(fgAO.ColSel)
            .Visible = True
            .ZOrder
            .SetFocus
        End With
    ElseIf fgAO.col = 2 Then
        With txtAO_Type
            .text = fgAO.TextMatrix(fgAO.RowSel, fgAO.ColSel)
            .Move tabConfiguration.Left + fraAO.Left + fgAO.Left + fgAO.ColPos(fgAO.ColSel), _
                  tabConfiguration.Top + fraAO.Top + fgAO.Top + fgAO.RowPos(fgAO.RowSel), _
                  fgAO.ColWidth(fgAO.ColSel)
            .Visible = True
            .ZOrder
            .SetFocus
        End With
    
    ElseIf fgAO.col = 3 Then
        With txtAO
            .text = fgAO.TextMatrix(fgAO.RowSel, fgAO.ColSel)
            .Move tabConfiguration.Left + fraAO.Left + fgAO.Left + fgAO.ColPos(fgAO.ColSel), _
                  tabConfiguration.Top + fraAO.Top + fgAO.Top + fgAO.RowPos(fgAO.RowSel), _
                  fgAO.ColWidth(fgAO.ColSel)
            .Visible = True
            .ZOrder
            .SetFocus
        End With
    End If
    
End Sub

Private Sub fgDI_Click()
        fgDI.Refresh
        With cmbIOList(0)
            .Move tabConfiguration.Left + fraDI.Left + fgDI.Left + fgDI.ColPos(fgDI.ColSel), _
                  tabConfiguration.Top + fraDI.Top + fgDI.Top + fgDI.RowPos(fgDI.RowSel), _
                  fgDI.ColWidth(fgDI.ColSel)
            .Visible = True
            .ZOrder
            .SetFocus
        End With
End Sub

Private Sub fgDO_Click()
        fgDO.Refresh
        With cmbIOList(1)
            .Move tabConfiguration.Left + fraDO.Left + fgDO.Left + fgDO.ColPos(fgDO.ColSel), _
                  tabConfiguration.Top + fraDO.Top + fgDO.Top + fgDO.RowPos(fgDO.RowSel), _
                  fgDO.ColWidth(fgDO.ColSel)
            .Visible = True
            .ZOrder
            .SetFocus
        End With
End Sub

Private Sub fgTeach_Click()
    Dim Index As Integer
    Dim i As Integer
    
    Index = fgTeach.RowSel
    txtTeachIndex = Index - 1
    For i = 0 To 3
        txtWriteTeach(i) = fgTeach.TextMatrix(Index, i + 1)
    Next i
    
End Sub

Private Sub imgDO_Click(Index As Integer)
        
    If SysDO.value(Index) = 0 Then
        SetDO CLng(Index), True
    Else
        SetDO CLng(Index), False
    End If

End Sub


Private Sub sldAO_Click(Index As Integer)
    Dim iRet As Integer
    Dim iValue As Integer

    iValue = sldAO(Index).value * 20
    If Index < 20 Then
        SetAO CLng(Index), CSng(iValue)
        'Call Ixud_WriteAOVoltage(gbintPIODA8A, Index, iValue)
    End If
        
    fgAO.TextMatrix(Index + 1, 2) = CStr(sldAO(Index).value)
End Sub



Private Sub tabConfiguration_Click(PreviousTab As Integer)
If tabConfiguration.Tab = 6 Then
fgAdvDaqAI.Visible = True
Else
fgAdvDaqAI.Visible = False
End If
End Sub


'Private Sub tabMain_Click(PreviousTab As Integer)
'If tabMain.Tab = 0 Then
'Frame1.Visible = True
''txtParaNormal(21).Visible = True
'Else
''txtParaNormal(21).Visible = False
'Frame1.Visible = False
'End If
'End Sub



Private Sub tmrFinishedBeep_Timer()
    tmrFinishedBeep.Enabled = False
    SetTower 2, True
End Sub

Private Sub tmrFinishedLight_Timer()
    If Kernel.sngTC(0) <= gbintFinishedLight Then
        tmrFinishedLight.Enabled = False
        mdifrmRTP.tbrRTP.Buttons("iRun").Enabled = True
        frmPlotProcess.fraProcessChart.BackColor = &H8000000F
        If gbintFinishedBeep > 0 Then
            tmrFinishedBeep.Interval = gbintFinishedBeep * 1000
            tmrFinishedBeep.Enabled = True
            SetTower 1, True
        End If
    End If
End Sub





Private Sub tmrTest_Timer()
    ReadTC
End Sub

Private Sub tmrWatchDog_Timer()
    Dim lngValue As Long
    If gblngCheckPC = gblngDO_PC_Check1 Then
        SetDO gblngDO_PC_Check1, False
        SetDO gblngDO_PC_Check2, True
        gblngCheckPC = gblngDO_PC_Check2
    Else
        SetDO gblngDO_PC_Check2, False
        SetDO gblngDO_PC_Check1, True
        gblngCheckPC = gblngDO_PC_Check1
    End If
End Sub

Public Sub StartWatchDog()
    tmrWatchDog.Enabled = True
    gblngCheckPC = gblngDO_PC_Check1
End Sub

Public Sub StopWatchDog()
    tmrWatchDog.Enabled = False
    SetDO gblngDO_PC_Check1, False
    SetDO gblngDO_PC_Check2, False
End Sub

Public Sub SetCTSlope()
    Dim i As Integer
    
    For i = 0 To GB_CT_MAX - 1
        gbsngCTGateSlope(i) = (gbsngIntensityRef(1) - gbsngIntensityRef(0)) / (gbsngCTGate2(i) - gbsngCTGate1(i))
    Next i
End Sub

Private Sub txtAlarm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        fgAlarm.TextMatrix(fgAlarm.row, fgAlarm.col) = txtAlarm.text
        txtAlarm.Visible = False
    End If
End Sub

Private Sub txtAO_KeyPress(KeyAscii As Integer)
    Dim sng As Single
    
    If KeyAscii = 13 Then
        sng = Val(txtAO.text)
        If sng > 5 Then sng = 5
        fgAO.TextMatrix(fgAO.row, fgAO.col) = txtAO.text
        txtAO.Visible = False
        SetAO CLng(fgAO.row + 1), sng
    End If
End Sub

Private Sub txtAO_LostFocus()
    txtAO.text = "0"
    txtAO.Visible = False
End Sub

Private Sub txtAO_Type_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    
    If KeyAscii = 13 Then
        i = Val(txtAO_Type.text)
        If i > 1 Then i = 0
        fgAO.TextMatrix(fgAO.row, fgAO.col) = txtAO_Type.text
        txtAO_Type.Visible = False
        
    End If
End Sub

Private Sub txtErrorMap_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        fgAdvDaqAI.TextMatrix(fgAdvDaqAI.row, fgAdvDaqAI.col) = txtErrorMap.text
        txtErrorMap.Visible = False
               
    End If
End Sub

Private Sub tmrDIO_Timer()
    Dim i As Integer
    
    Call ReadDI
        
    If Kernel.IsRun = 1 Then
        If SysDI.IsReady = 0 Then
            ShowAlarmFlash 10
        End If
        
        If SysDI.IsDoorClose = 0 And Kernel.IsDoorMoving = 0 Then
            ShowAlarmFlash 11
        End If
        
        If SysDI.IsCDA = 0 Then
            ShowAlarmFlash 12
        End If
        
        If SysDI.IsWater = 0 Then
            ShowAlarmFlash 13
        End If
        
        For i = 1 To 4
            If SysDI.IsLampError(i) = 1 Then
                SysDI.LampErrorCount = SysDI.LampErrorCount + 1
                If SysDI.LampErrorCount > 50 Then
                    SysDI.LampErrorCount = 0
                    ShowAlarmFlash 14
                    Exit For
                End If
            End If
        Next i
        If SysDI.IsLampError(1) = 0 And SysDI.IsLampError(2) = 0 And SysDI.IsLampError(3) = 0 And SysDI.IsLampError(4) = 0 Then
            SysDI.LampErrorCount = 0
        End If
        
    Else
        If SysDI.IsDoorClose = 0 And Kernel.IsNeedTestRun = 0 Then
            SetLampCooling True
        Else
            SetLampCooling False
        End If
    End If
    
    If SysDI.IsOverHeat = 1 Then
        SysDI.OverHeatCount = SysDI.OverHeatCount + 1
        If SysDI.OverHeatCount > 50 Then
            SysDI.OverHeatCount = 0
            ShowAlarmFlash 2
        End If
    Else
        SysDI.OverHeatCount = 0
    End If
    If SysDI.IsEMO = 1 Then
        
        ShowAlarmFlash 18
    End If
        
    frmPlotProcess.ShowStatus
    frmDiagnosis.ShowStatus
    mdifrmRTP.ShowStatus
    Me.ShowStatus
End Sub

Private Sub tmrAIO_Timer()
    Dim i As Integer
    Dim j As Long
    Dim iRet As Integer
    Dim bRet As Boolean
    Dim iValue As Integer
    Dim lngValue As Long
    Dim sngValue As Single
    
    On Error GoTo ERRLINE
    
    Call ReadAI
    
    If Kernel.IsRun = 0 Then
        ReadTC
        If frmPlotProcess.tmrIdleCheck.Enabled = False And gbsngIdleWarning > 0 Then
            frmPlotProcess.tmrIdleCheck.Enabled = True
        End If
    Else
        frmPlotProcess.tmrIdleCheck.Enabled = False
    End If
 
    If gbsngGaugeValue > 0 Then
'        If Kernel.sngPressure < gbsngGaugeValue / 1000 Then
       If Kernel.sngPressure < gbsngGaugeValue / 1000 And SysAI.AvgValue(gblngAI_Vacuum_Gauge) > 0 Then
            SetDO gblngDO_APCGaugeValve, True
        Else
            SetDO gblngDO_APCGaugeValve, False
        End If
    End If
    
      If gbsngGaugeValue2 > 0 Then
'        If Kernel.sngPressure < gbsngGaugeValue / 1000 Then
       If Kernel.sngPressure2 < gbsngGaugeValue2 / 1000 And SysAI.AvgValue(gblngAI_Vacuum_Gauge2) > 0 Then
            SetDO gblngDO_APCGaugeValve, True
        Else
            SetDO gblngDO_APCGaugeValve, False
        End If
    End If
    
    
    
    If Para.sngGaugeAngle > 0 Then
        If Kernel.sngPressure < Para.sngGaugeAngle / 1000 Then
            SetDO gblngDO_APCGaugeAngle, True
        Else
            SetDO gblngDO_APCGaugeAngle, False
        End If
    End If
    
    
    frmPlotProcess.ShowStatus
    frmDiagnosis.ShowStatus
    Me.ShowStatus
    Exit Sub
ERRLINE:
    gbstrAlarmHint = " DIO Timer error"
    ShowAlarmFlash 1
End Sub

Public Sub ShowStatus()
    Dim i As Integer
            
    For i = 0 To 31
        fgAI.TextMatrix(i + 1, 2) = Format(SysAI.value(i), "0.00")
        imgDI(i) = IIf(SysDI.value(i) = 0, imgOff, imgOn)
        imgDO(i) = IIf(SysDO.value(i) = 0, imgOff, imgOn)
    Next i
    For i = 0 To 23
        fgAO.TextMatrix(i + 1, 2) = Format(SysAO.value(i), "0.00")
        fgAdvDaqAI.TextMatrix(i + 1, 5) = Format(Kernel.sngOrigTC(i), "0.00")
    Next i
   
End Sub

Public Sub RefreshFunctionRight()
    fraDatabase(0).Visible = IIf(gbintActiveModule_Database = 1, True, False)
    txtServerPath.Visible = chkBarcodeServer.Enabled
    fraDatabase(1).Visible = IIf(gbintActiveModule_CIM = 1, True, False)
    fraAutoMode(4).Visible = IIf(gbintActiveModule_Auto = 1, True, False)
    
End Sub

Private Sub txtErrorV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        fgAI.TextMatrix(fgAI.row, fgAI.col) = txtErrorV.text
        txtErrorV.Visible = False
               
    End If
End Sub

Private Sub txtErrorV_LostFocus()
    txtErrorV.text = "0"
    txtErrorV.Visible = False
End Sub




Private Sub TxthdOffset_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If TxthdOffset.text <> "" Then
Call WritePrivateProfileString("Special_Setting", "Hold_Offset", TxthdOffset.text, App.Path + ProcDict_Path)
MsgBox "設定成功"
Else
MsgBox "請輸入Hold Offset值"
End If
End If
End Sub



Private Sub TxthdTimes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If TxthdTimes.text <> "" Then
Call WritePrivateProfileString("Special_Setting", "Hold_Times", TxthdTimes.text, App.Path + ProcDict_Path)
MsgBox "設定成功"
Else
MsgBox "請輸入Hold Times值"
End If
End If
End Sub

Private Sub txtWriteTeach_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim i As Integer
    Dim ii As Integer
    
    If KeyAscii = 13 Then
        ii = Val(txtTeachIndex.text) + 1
        For i = 0 To 3
            fgTeach.TextMatrix(ii, i + 1) = txtWriteTeach(i)
        Next i
    End If
End Sub

Private Sub tmrSendCT_Timer()
    Dim i As Integer
    
    tmrSendCT.Enabled = False
    tmrTimeoutCT.Enabled = True
    
    

    Kernel.IsPM = 0
    '0~7,9~16,17-22
    
    'gbintPMcmdID = 2
    Call SendCmdPM(gbintPMcmdID)
'    If Kernel.IsPM = 0 Then
'        If intSendPMCount < 5 Then
'            intSendPMCount = intSendPMCount + 1
'            'gbintPMcmdID = 8
'            Call SendCmdPM(gbintPMcmdID)
'
'
'        Else
'            intSendPMCount = 0
'            ShowAlarmFlash 20
'            tmrSendCT.Enabled = False
'        End If
'    Else
'        gbintPMcmdID = gbintPMcmdID + 1
'        If gbintPMcmdID > 10 Then
'            gbintPMcmdID = 0
'        End If
'
'        Kernel.IsPM = 0
'        intSendPMCount = 0
'        'gbintPMcmdID = 8
'        Call SendCmdPM(gbintPMcmdID)
       
        
'    End If
  mdifrmRTP.SaveToHistory Kernel.dblCT
    
End Sub

Private Sub tmrTimeoutCT_Timer()
    
    tmrTimeoutCT.Enabled = False
    If intSendPMCount < 5 Then
        intSendPMCount = intSendPMCount + 1
        tmrSendCT.Enabled = True
        'Call SendCmdPM(gbintPMcmdID)
    Else
        intSendPMCount = 0
        ShowAlarmFlash 20
        tmrSendCT.Enabled = False
    End If
End Sub

Private Sub MSComm1_OnComm()
    Dim InByte() As Byte, InData_B() As Byte
    Dim RawData() As Byte, inData() As Byte
    Dim i As Integer, j As Integer
    Dim tmpTime As Long
    Dim tmpStr As String
    Dim DataLength As Integer
    Dim HD_Ptr As Integer, bFind_HD As Boolean
    Dim sng As Single
    Dim a(4) As Byte
    Dim imaxByte, imaxCard, imaxPort, ii As Integer
    
    If MSComm1.CommEvent = comEvReceive Then
        tmrTimeoutCT.Enabled = False
        
        InByte = MSComm1.Input

        inData = InByte
        
        InData_B = InByte
        DataLength = UBound(InData_B) + 1

        tmpStr = ""
        imaxByte = UBound(InByte)
        If imaxByte > 148 Then imaxByte = 148
        For i = 0 To imaxByte
            ReadData(i) = InByte(i)
        Next
        
        imaxPort = 2
        If gbintPMcmdID >= Para.intPMbig * 8 Then imaxPort = 3
        
'        imax = 2
'        If gbintPMcmdID > 7 Then
'            imax = 3
'        End If
        For i = 0 To imaxPort
            a(3) = ReadData(3 + i * 36 + 4 + 2)
            a(2) = ReadData(3 + i * 36 + 4 + 3)
            a(1) = ReadData(3 + i * 36 + 4)
            a(0) = ReadData(3 + i * 36 + 4 + 1)

            CopyMemory sng, a(0), 4
            
            If gbintPMcmdID < Para.intPMbig * 8 Then
                ii = gbintPMcmdID * 3 + i
                Kernel.dblCT(ii) = CDbl(sng)
            Else
                ii = Para.intPMbig * 24 + (gbintPMcmdID - Para.intPMbig * 8) * 4 + i
                Kernel.dblCT(ii) = CDbl(sng)
                
            End If
            
'            If gbintPMcmdID < 9 Then
'                ii = gbintPMcmdID * 3 + i
'                Kernel.dblCT(ii) = CDbl(sng)
'            Else
'                If gbintPMcmdID = 9 Then Kernel.dblCT(28 + i) = CDbl(sng)
'                If gbintPMcmdID = 10 Then Kernel.dblCT(32 + i) = CDbl(sng)
'            End If
                  
        Next i
        intSendPMCount = 0
        gbintPMcmdID = gbintPMcmdID + 1
        
        imaxCard = Para.intPMbig * 8 + Para.intPMsmall - 1
'        Para.intPMbig = Val(txtPMbig.Text)
'        Para.intPMsmall = Val(txtPMsmall.Text)
        
        
'        If gbintNumOfBanks = 6 Then
'            imax = 7
'        Else
'            If gbintNumOfBanks = 13 Then
'                imax = 8
'            Else
'                imax = gbintNumOfBanks - 2
'            End If
'        End If
        If gbintPMcmdID > imaxCard Then
            gbintPMcmdID = 0
        End If
        frmConfiguration.tmrSendCT.Enabled = True
        Kernel.IsPM = 1
    End If
    
    
End Sub


'將offset寫入TCM方法

Public Sub WriteOffsetToTCM()
    Dim StrFileName As String
    Dim OffsetValue          As Long
    Dim address As Long
    Dim r1, r2 As Boolean
    Dim i As Integer
     Dim StrData(50)    As String * 4
     
    address = 316
    
    '取system�堶悸榣ffset值
'    On Error GoTo ERR_PARAMETER_OPEN
    StrFileName = gbSystemPath & "\Config\Procdict.ini"
'    If dir(strFileName) = "" Then GoTo ERR_PARAMETER_OPEN
    
    
    For i = 1 To GB_MAX_LOOPS - 1
        OffsetValue = GetPrivateProfileString("Offset", "OffsetEX" & CStr(i), "0", StrData(0), 20, StrFileName)
        If i = 5 Then
            r2 = frmAz2.WritePara(address, -CInt(StrData(0)))
        Else
            r1 = frmAz1.WritePara(address, -CInt(StrData(0)))
        End If
        address = address + 100
    Next i
'    Exit Sub
'ERR_PARAMETER_OPEN:
'    Call AlertShow("Open Parameter Failed!!", ERRORTYPE)
End Sub

Private Sub SaveOffsetValue()
    Dim lngRet As Long
    Dim StrFileName As String
    Dim StrData(50)    As String * 30
    Dim i As Integer
    
    
    StrFileName = gbSystemPath & "\Config\Procdict.ini"
    
    For i = 1 To GB_MAX_LOOPS - 1
    
        If Val(txtRatioEX(i).text) >= -100 And Val(txtRatioEX(i).text) <= 100 Then
        
'            gbsngRatioEX(i) = Val(txtRatioEX(i).text)
'            lngRet = WritePrivateProfileString("PARAMETER", "RatioEX" & CStr(i), txtRatioEX(i).text, strFileName)
            
            lngRet = WritePrivateProfileString("Offset", "OffsetEX" & CStr(i), txtRatioEX(i).text, StrFileName)
            
'            If i <> GB_MAX_LOOPS - 1 Then
'            gbsngErrorTC(i) = Val(txtRatioEX(i).text)
'            lngRet = WritePrivateProfileString("TCModuleErr", "Error" & CStr(i), txtRatioEX(i).text, strFileName)
'            End If
        Else
            MsgBox "TC" + CStr(i) + "對應offset請輸入-100到100之間的數值！"
        End If
    Next i

End Sub

Private Sub OpenOffsetValue()
      Dim lngRet As Long
      Dim StrFileName As String
      Dim StrData(50)    As String * 30
      Dim StrData2(50)    As String * 30
      Dim i As Integer
    
      StrFileName = gbSystemPath & "\Config\Procdict.ini"
      
      For i = 1 To GB_MAX_LOOPS - 1
        lngRet = GetPrivateProfileString("Offset", "OffsetEX" & CStr(i), "1", StrData(0), 20, StrFileName)
        txtRatioEX(i).text = StrData(0)
'        gbsngRatioEX(i) = Val(txtRatioEX(i).text)
        
'        lngRet = GetPrivateProfileString("TCModuleErr", "Error" & CStr(i), "0", StrData2(i), 20, strFileName)
'        fgAdvDaqAI.TextMatrix(i, 3) = StrData(0)
    Next i

End Sub


Private Sub OpenFunctionIni()
      Dim StrFileName As String
      Dim Switch As String
      On Error GoTo ERR_OpenFuctionIni
      StrFileName = gbSystemPath & "\Config\Function.ini"
      Switch = CommnonReadini("FuncSwitch", "TcOffset", StrFileName)
      ChkTcOffset.value = CInt(Switch)
      GbTcoffset_Switch = ChkTcOffset.value
      Switch = CommnonReadini("FuncSwitch", "ShowToolBar14", StrFileName)
      ChkShowChamberNo.value = CInt(Switch)
      GbChamberNo_Switch = ChkShowChamberNo.value
      Switch = CommnonReadini("FuncSwitch", "ShowToolBar15", StrFileName)
      GbShowDebugButton = CInt(Switch)
      If GbShowDebugButton = 1 Then
        mdifrmRTP.tbrRTP.Buttons(15).Visible = True
        Else
        mdifrmRTP.tbrRTP.Buttons(15).Visible = False
      End If
      Switch = CommnonReadini("FuncSwitch", "TestMode", StrFileName)
      GbTestMode_Switch = CInt(Switch)
      ChkTestMode.value = GbTestMode_Switch
      frmPlotProcess.FraVac2.Visible = IIf(GbTestMode_Switch = 1, True, False)
      Exit Sub
ERR_OpenFuctionIni:
    Call AlertShow("Open Function.ini Failed!!", ERRORTYPE)
End Sub

Private Sub SaveFunctionIni()
      Dim StrFileName As String
      Dim lngRet As Long
      On Error GoTo ERR_SaveFunctionIni
      StrFileName = gbSystemPath & "\Config\Function.ini"
      lngRet = WritePrivateProfileString("FuncSwitch", "TcOffset", CStr(ChkTcOffset.value), StrFileName)
      lngRet = WritePrivateProfileString("FuncSwitch", "ShowToolBar14", CStr(ChkShowChamberNo.value), StrFileName)
      lngRet = WritePrivateProfileString("FuncSwitch", "StopTCM", CStr(ckeStopTCM.value), StrFileName)
      lngRet = WritePrivateProfileString("FuncSwitch", "ProcStepInUse", CStr(ChkDefineProcStep.value), StrFileName)
      lngRet = WritePrivateProfileString("FuncSwitch", "OffsetWriteToTcm", CStr(ChkWriteOffsetToTCM.value), StrFileName)
      lngRet = WritePrivateProfileString("FuncSwitch", "ForcePreheat", CStr(ckeForcePreheat.value), StrFileName)
      lngRet = WritePrivateProfileString("FuncSwitch", "CTDisplay", CStr(ckeCTDisplay.value), StrFileName)
      lngRet = WritePrivateProfileString("FuncSwitch", "CheckTcWafer", CStr(chkAlarmEnable(14).value), StrFileName)
      lngRet = WritePrivateProfileString("FuncSwitch", "CheckRValue", CStr(chkAlarmEnable(15).value), StrFileName)
      lngRet = WritePrivateProfileString("FuncSwitch", "TestMode", CStr(ChkTestMode.value), StrFileName)
      gbintActiveAlarm_TcWafer = chkAlarmEnable(14).value
      gbintActiveAlarm_RValue = chkAlarmEnable(15).value
      lngRet = WritePrivateProfileString("Device2", "IsUsed", CStr(ckSCREable.value), App.Path + ModbusRtu_Path)
      lngRet = WritePrivateProfileString("Device2", "Port", "COM" + txtCommSCR.text, App.Path + ModbusRtu_Path)
      Exit Sub
ERR_SaveFunctionIni:
    Call AlertShow("Save Function.ini Failed!!", ERRORTYPE)
End Sub

