VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{D731BCEE-1E13-419B-AB0D-5AD3E9FD841B}#1.0#0"; "G_AngleValve.ocx"
Object = "{CAB6E897-4690-4558-BB71-FA8192D14898}#1.0#0"; "G_TankA.ocx"
Object = "{EE4D4E3B-DE0B-40F8-A1FA-B6C4FCC2CEB9}#3.0#0"; "G_ReleaseValve.ocx"
Object = "{51107E11-7FE1-4465-A58A-ECF978ABC595}#1.0#0"; "G_TankB.ocx"
Object = "{FC17F15C-7BAC-467F-A861-FDA894BE46F6}#1.0#0"; "G_Pump.ocx"
Begin VB.Form frmDiagnosis 
   Caption         =   "Diagnosis"
   ClientHeight    =   11910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17340
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
   ScaleHeight     =   11910
   ScaleWidth      =   17340
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab tabMain 
      Height          =   11895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   17175
      _ExtentX        =   30295
      _ExtentY        =   20981
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   88
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmDiagnosis.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraIntensity"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraCTCheck"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraVacFunc"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraBankH"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraAZ"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fraTurbo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fraCover"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmDiagnosis.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.Frame fraCover 
         Caption         =   "蓋版控制"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1095
         Left            =   11760
         TabIndex        =   174
         Top             =   8640
         Visible         =   0   'False
         Width           =   2175
         Begin VB.CommandButton cmdCoverOrig 
            Caption         =   "歸零"
            Height          =   495
            Left            =   1320
            TabIndex        =   181
            Top             =   360
            Width           =   735
         End
         Begin VB.CommandButton cmdCoverDown 
            Caption         =   "下"
            Height          =   495
            Left            =   3000
            TabIndex        =   178
            Top             =   360
            Width           =   735
         End
         Begin VB.CommandButton cmdCoverUp 
            Caption         =   "上"
            Height          =   495
            Left            =   2160
            TabIndex        =   177
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lbCoverDown 
            Alignment       =   2  'Center
            BackColor       =   &H0000FF00&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   3120
            TabIndex        =   180
            Top             =   150
            Width           =   540
         End
         Begin VB.Label lbCoverUp 
            Alignment       =   2  'Center
            BackColor       =   &H0000FF00&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   2280
            TabIndex        =   179
            Top             =   150
            Width           =   540
         End
         Begin VB.Label lbCoverStatus 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "停止"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   176
            Top             =   720
            Width           =   600
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "狀態:"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   16
            Left            =   240
            TabIndex        =   175
            Top             =   360
            Width           =   750
         End
      End
      Begin VB.Frame fraTurbo 
         Caption         =   "Turbo Pump"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1095
         Left            =   11760
         TabIndex        =   172
         Top             =   7200
         Visible         =   0   'False
         Width           =   1935
         Begin VB.CheckBox chkTurbo 
            Caption         =   "開啟Turbo"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   173
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame fraAZ 
         Caption         =   "能量輸出控制"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1935
         Left            =   11760
         TabIndex        =   162
         Top             =   120
         Width           =   3855
         Begin VB.TextBox txtBLK_Keep 
            Height          =   390
            Left            =   1080
            TabIndex        =   171
            Text            =   "5"
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox txtBLK_Output 
            Height          =   390
            Left            =   1080
            TabIndex        =   170
            Text            =   "10"
            Top             =   840
            Width           =   975
         End
         Begin VB.CheckBox chkBlkTest 
            Caption         =   "開啟測試"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   164
            Top             =   360
            Width           =   975
         End
         Begin VB.ComboBox cmbBlkList 
            Height          =   390
            ItemData        =   "frmDiagnosis.frx":0038
            Left            =   1080
            List            =   "frmDiagnosis.frx":003A
            TabIndex        =   163
            Text            =   "BLK-01"
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lbName 
            Caption         =   "Sec"
            Height          =   255
            Index           =   15
            Left            =   2040
            TabIndex        =   169
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Keep"
            Height          =   270
            Index           =   12
            Left            =   240
            TabIndex        =   168
            Top             =   1320
            Width           =   570
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Power"
            Height          =   270
            Index           =   9
            Left            =   240
            TabIndex        =   167
            Top             =   840
            Width           =   675
         End
         Begin VB.Label lbName 
            Caption         =   "%"
            Height          =   255
            Index           =   8
            Left            =   2160
            TabIndex        =   166
            Top             =   840
            Width           =   255
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Block"
            Height          =   270
            Index           =   5
            Left            =   360
            TabIndex        =   165
            Top             =   360
            Width           =   585
         End
      End
      Begin VB.Frame fraBankH 
         Caption         =   "Lamp Current"
         Height          =   2175
         Left            =   120
         TabIndex        =   101
         Top             =   9600
         Width           =   16935
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgCTConfigProcess 
            Height          =   1815
            Left            =   0
            TabIndex        =   187
            Top             =   240
            Visible         =   0   'False
            Width           =   16695
            _ExtentX        =   29448
            _ExtentY        =   3201
            _Version        =   393216
            BorderStyle     =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   41
            Left            =   600
            TabIndex        =   120
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   42
            Left            =   1080
            TabIndex        =   119
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   43
            Left            =   1560
            TabIndex        =   118
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   44
            Left            =   2040
            TabIndex        =   117
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   45
            Left            =   3240
            TabIndex        =   116
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   46
            Left            =   3720
            TabIndex        =   115
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   47
            Left            =   4200
            TabIndex        =   114
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   48
            Left            =   4680
            TabIndex        =   113
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   49
            Left            =   5160
            TabIndex        =   112
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   50
            Left            =   6120
            TabIndex        =   111
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   51
            Left            =   6600
            TabIndex        =   110
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   52
            Left            =   7080
            TabIndex        =   109
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   53
            Left            =   7560
            TabIndex        =   108
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   54
            Left            =   8040
            TabIndex        =   107
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   55
            Left            =   9000
            TabIndex        =   106
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   56
            Left            =   9480
            TabIndex        =   105
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   57
            Left            =   9960
            TabIndex        =   104
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   58
            Left            =   10440
            TabIndex        =   103
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   59
            Left            =   10920
            TabIndex        =   102
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   40
            Left            =   120
            TabIndex        =   121
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   20
            Left            =   120
            TabIndex        =   141
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   33
            Left            =   7560
            TabIndex        =   140
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   34
            Left            =   8040
            TabIndex        =   139
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   35
            Left            =   9000
            TabIndex        =   138
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   29
            Left            =   5160
            TabIndex        =   137
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   30
            Left            =   6120
            TabIndex        =   136
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   31
            Left            =   6600
            TabIndex        =   135
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   32
            Left            =   7080
            TabIndex        =   134
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   25
            Left            =   3240
            TabIndex        =   133
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   26
            Left            =   3720
            TabIndex        =   132
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   27
            Left            =   4200
            TabIndex        =   131
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   28
            Left            =   4680
            TabIndex        =   130
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   21
            Left            =   600
            TabIndex        =   129
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   22
            Left            =   1080
            TabIndex        =   128
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   23
            Left            =   1560
            TabIndex        =   127
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   24
            Left            =   2040
            TabIndex        =   126
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   36
            Left            =   9480
            TabIndex        =   125
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   37
            Left            =   9960
            TabIndex        =   124
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   38
            Left            =   10440
            TabIndex        =   123
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   39
            Left            =   10920
            TabIndex        =   122
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   161
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   1
            Left            =   600
            TabIndex        =   160
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   2
            Left            =   1080
            TabIndex        =   159
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   5
            Left            =   3240
            TabIndex        =   158
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   4
            Left            =   2040
            TabIndex        =   157
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   3
            Left            =   1560
            TabIndex        =   156
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   7
            Left            =   4200
            TabIndex        =   155
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   6
            Left            =   3720
            TabIndex        =   154
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   9
            Left            =   5160
            TabIndex        =   153
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   8
            Left            =   4680
            TabIndex        =   152
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   11
            Left            =   6600
            TabIndex        =   151
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   10
            Left            =   6120
            TabIndex        =   150
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   14
            Left            =   8040
            TabIndex        =   149
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   13
            Left            =   7560
            TabIndex        =   148
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   12
            Left            =   7080
            TabIndex        =   147
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   17
            Left            =   9960
            TabIndex        =   146
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   16
            Left            =   9480
            TabIndex        =   145
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   15
            Left            =   9000
            TabIndex        =   144
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   18
            Left            =   10440
            TabIndex        =   143
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lbCT 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   375
            Index           =   19
            Left            =   10920
            TabIndex        =   142
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame fraVacFunc 
         Caption         =   "腔體壓力控制"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1095
         Left            =   11760
         TabIndex        =   29
         Top             =   6120
         Visible         =   0   'False
         Width           =   3855
         Begin VB.CheckBox chkPurge 
            Caption         =   "破真空"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2040
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox chkPumping 
            Caption         =   "開啟泵"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Height          =   9375
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   11535
         Begin VB.TextBox TxtChamberNo 
            Alignment       =   2  'Center
            Height          =   495
            Left            =   1320
            TabIndex        =   185
            Top             =   360
            Width           =   1335
         End
         Begin VB.Timer tmrSetReleaseOFF 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   1560
            Top             =   1560
         End
         Begin VB.Timer tmrPreHeat 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   1080
            Top             =   1560
         End
         Begin VB.Timer tmrCoverOrig 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   600
            Top             =   1560
         End
         Begin VB.Frame fraOxygen 
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1695
            Left            =   8160
            TabIndex        =   98
            Top             =   7560
            Width           =   1815
            Begin VB.Label lbOxygen 
               Alignment       =   2  'Center
               Caption         =   "999"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   150
               TabIndex        =   100
               Top             =   530
               Width           =   1500
            End
            Begin VB.Shape Shape3 
               BackColor       =   &H00004000&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00808080&
               BorderWidth     =   5
               Height          =   495
               Left            =   120
               Shape           =   4  'Rounded Rectangle
               Top             =   480
               Width           =   1575
            End
            Begin VB.Label lbName 
               Alignment       =   2  'Center
               Caption         =   "氧氣(ppm)"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   7
               Left            =   120
               TabIndex        =   99
               Top             =   1200
               Width           =   1575
            End
         End
         Begin VB.Timer tmrCheckMFC 
            Enabled         =   0   'False
            Index           =   5
            Interval        =   1000
            Left            =   2520
            Top             =   960
         End
         Begin VB.Timer tmrCheckMFC 
            Enabled         =   0   'False
            Index           =   4
            Interval        =   1000
            Left            =   2040
            Top             =   960
         End
         Begin VB.Timer tmrHoldSafeON 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   120
            Top             =   1560
         End
         Begin VB.Timer tmrCheckMFC 
            Enabled         =   0   'False
            Index           =   3
            Interval        =   1000
            Left            =   1560
            Top             =   960
         End
         Begin VB.Timer tmrCheckMFC 
            Enabled         =   0   'False
            Index           =   2
            Interval        =   1000
            Left            =   1080
            Top             =   960
         End
         Begin VB.Timer tmrCheckMFC 
            Enabled         =   0   'False
            Index           =   1
            Interval        =   1000
            Left            =   600
            Top             =   960
         End
         Begin VB.Timer tmrCheckMFC 
            Enabled         =   0   'False
            Index           =   0
            Interval        =   1000
            Left            =   120
            Top             =   960
         End
         Begin VB.Frame Frame5 
            Caption         =   "目前狀態"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1695
            Left            =   1440
            TabIndex        =   48
            Top             =   7560
            Width           =   6615
            Begin VB.Label lbName 
               Alignment       =   2  'Center
               Caption         =   "燈管檢知"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   3
               Left            =   3360
               TabIndex        =   83
               Top             =   1200
               Width           =   1215
            End
            Begin VB.Shape shpCT 
               BackColor       =   &H00000080&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00004000&
               BorderWidth     =   3
               Height          =   375
               Index           =   4
               Left            =   3840
               Shape           =   4  'Rounded Rectangle
               Top             =   600
               Width           =   375
            End
            Begin VB.Shape shpCT 
               BackColor       =   &H00000080&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00004000&
               BorderWidth     =   3
               Height          =   375
               Index           =   3
               Left            =   3480
               Shape           =   4  'Rounded Rectangle
               Top             =   600
               Width           =   375
            End
            Begin VB.Shape shpCT 
               BackColor       =   &H00000080&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00004000&
               BorderWidth     =   3
               Height          =   375
               Index           =   2
               Left            =   3840
               Shape           =   4  'Rounded Rectangle
               Top             =   240
               Width           =   375
            End
            Begin VB.Shape shpCT 
               BackColor       =   &H00000080&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00004000&
               BorderWidth     =   3
               Height          =   375
               Index           =   1
               Left            =   3480
               Shape           =   4  'Rounded Rectangle
               Top             =   240
               Width           =   375
            End
            Begin VB.Label lbName 
               Alignment       =   2  'Center
               Caption         =   "壓力(Torr)"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   4800
               TabIndex        =   82
               Top             =   1200
               Width           =   1575
            End
            Begin VB.Label lbVacuum 
               Alignment       =   2  'Center
               BackColor       =   &H0000FF00&
               Caption         =   "760"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   4500
               TabIndex        =   81
               Top             =   510
               Width           =   1980
            End
            Begin VB.Shape shpOverheat 
               BackColor       =   &H00000080&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00004000&
               BorderWidth     =   5
               Height          =   495
               Left            =   2520
               Shape           =   4  'Rounded Rectangle
               Top             =   480
               Width           =   495
            End
            Begin VB.Shape shpReady 
               BackColor       =   &H00004000&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00808080&
               BorderWidth     =   5
               Height          =   495
               Left            =   600
               Shape           =   4  'Rounded Rectangle
               Top             =   480
               Width           =   495
            End
            Begin VB.Shape shpAlarm 
               BackColor       =   &H00000080&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00004000&
               BorderWidth     =   5
               Height          =   495
               Left            =   1560
               Shape           =   4  'Rounded Rectangle
               Top             =   480
               Width           =   495
            End
            Begin VB.Label lbName 
               Alignment       =   2  'Center
               Caption         =   "備妥"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   28
               Left            =   240
               TabIndex        =   51
               Top             =   1200
               Width           =   1215
            End
            Begin VB.Label lbName 
               Alignment       =   2  'Center
               Caption         =   "警報"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   36
               Left            =   1440
               TabIndex        =   50
               Top             =   1200
               Width           =   735
            End
            Begin VB.Label lbName 
               Alignment       =   2  'Center
               Caption         =   "過熱"
               BeginProperty Font 
                  Name            =   "標楷體"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   38
               Left            =   2400
               TabIndex        =   49
               Top             =   1200
               Width           =   735
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H00004000&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00808080&
               BorderWidth     =   5
               Height          =   495
               Left            =   4470
               Shape           =   4  'Rounded Rectangle
               Top             =   480
               Width           =   2055
            End
         End
         Begin VB.Frame fraMonitorTC 
            Caption         =   "監控點溫度(℃)"
            Height          =   3375
            Left            =   9120
            TabIndex        =   35
            Top             =   240
            Width           =   2175
            Begin VB.Label lbTC 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               ForeColor       =   &H00800000&
               Height          =   330
               Index           =   0
               Left            =   1080
               TabIndex        =   183
               Top             =   360
               Width           =   915
            End
            Begin VB.Label lbNameMonitorTC 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " MTC 0"
               ForeColor       =   &H00800000&
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   182
               Top             =   360
               Width           =   975
            End
            Begin VB.Label lbTC 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               ForeColor       =   &H00800000&
               Height          =   330
               Index           =   7
               Left            =   1080
               TabIndex        =   97
               Top             =   2880
               Width           =   915
            End
            Begin VB.Label lbNameMonitorTC 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " MTC 7"
               ForeColor       =   &H00800000&
               Height          =   330
               Index           =   7
               Left            =   120
               TabIndex        =   96
               Top             =   2880
               Width           =   975
            End
            Begin VB.Label lbNameMonitorTC 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " MTC 3"
               ForeColor       =   &H00800000&
               Height          =   330
               Index           =   3
               Left            =   120
               TabIndex        =   47
               Top             =   1440
               Width           =   975
            End
            Begin VB.Label lbTC 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               ForeColor       =   &H00800000&
               Height          =   330
               Index           =   3
               Left            =   1080
               TabIndex        =   46
               Top             =   1440
               Width           =   915
            End
            Begin VB.Label lbNameMonitorTC 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " MTC 2"
               ForeColor       =   &H00800000&
               Height          =   330
               Index           =   2
               Left            =   120
               TabIndex        =   45
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label lbTC 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               ForeColor       =   &H00800000&
               Height          =   330
               Index           =   2
               Left            =   1080
               TabIndex        =   44
               Top             =   1080
               Width           =   915
            End
            Begin VB.Label lbNameMonitorTC 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " MTC 1"
               ForeColor       =   &H00800000&
               Height          =   330
               Index           =   1
               Left            =   120
               TabIndex        =   43
               Top             =   720
               Width           =   975
            End
            Begin VB.Label lbTC 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               ForeColor       =   &H00800000&
               Height          =   330
               Index           =   1
               Left            =   1080
               TabIndex        =   42
               Top             =   720
               Width           =   915
            End
            Begin VB.Label lbTC 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               ForeColor       =   &H00800000&
               Height          =   330
               Index           =   4
               Left            =   1080
               TabIndex        =   41
               Top             =   1800
               Width           =   915
            End
            Begin VB.Label lbNameMonitorTC 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " MTC 4"
               ForeColor       =   &H00800000&
               Height          =   330
               Index           =   4
               Left            =   120
               TabIndex        =   40
               Top             =   1800
               Width           =   975
            End
            Begin VB.Label lbNameMonitorTC 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " MTC 5"
               ForeColor       =   &H00800000&
               Height          =   330
               Index           =   5
               Left            =   120
               TabIndex        =   39
               Top             =   2160
               Width           =   975
            End
            Begin VB.Label lbTC 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               ForeColor       =   &H00800000&
               Height          =   330
               Index           =   5
               Left            =   1080
               TabIndex        =   38
               Top             =   2160
               Width           =   915
            End
            Begin VB.Label lbNameMonitorTC 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " MTC 6"
               ForeColor       =   &H00800000&
               Height          =   330
               Index           =   6
               Left            =   120
               TabIndex        =   37
               Top             =   2520
               Width           =   975
            End
            Begin VB.Label lbTC 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               ForeColor       =   &H00800000&
               Height          =   330
               Index           =   6
               Left            =   1080
               TabIndex        =   36
               Top             =   2520
               Width           =   915
            End
         End
         Begin VB.Timer tmrPurge 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   1560
            Top             =   360
         End
         Begin VB.Timer tmrPumpON 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   120
            Top             =   360
         End
         Begin VB.Timer tmrPumpOFF 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   600
            Top             =   360
         End
         Begin VB.Timer tmrPowerTest 
            Enabled         =   0   'False
            Interval        =   3000
            Left            =   1080
            Top             =   360
         End
         Begin G_TankB.GTankB gtcAirTank 
            Height          =   735
            Left            =   4200
            TabIndex        =   52
            Top             =   1080
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   1296
            ShowMode        =   1
         End
         Begin G_TankA.GTankA gtcTankWater 
            Height          =   855
            Left            =   5520
            TabIndex        =   53
            Top             =   1080
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   1508
         End
         Begin VB.Frame fraVacuum 
            BorderStyle     =   0  'None
            Height          =   2415
            Left            =   660
            TabIndex        =   54
            Top             =   4125
            Visible         =   0   'False
            Width           =   1575
            Begin G_ReleaseValve.GReleaseValve gtcRV 
               Height          =   735
               Left            =   -360
               TabIndex        =   55
               Top             =   2280
               Visible         =   0   'False
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   1296
            End
            Begin G_AngleValve.GAngleValve gtcAV 
               Height          =   495
               Left            =   0
               TabIndex        =   56
               Top             =   360
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   873
            End
            Begin G_Pump.GPump gtcPump 
               Height          =   615
               Left            =   0
               TabIndex        =   57
               Top             =   1320
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   1085
            End
            Begin VB.Label lbRV_Status 
               Caption         =   "OFF"
               Height          =   255
               Left            =   -240
               TabIndex        =   63
               Top             =   3360
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.Label lbAV_Status 
               Caption         =   "Close"
               Height          =   255
               Left            =   480
               TabIndex        =   62
               Top             =   240
               Width           =   855
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "RV"
               Height          =   270
               Index           =   26
               Left            =   -240
               TabIndex        =   61
               Top             =   3120
               Visible         =   0   'False
               Width           =   330
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "AV"
               Height          =   270
               Index           =   10
               Left            =   0
               TabIndex        =   60
               Top             =   0
               Width           =   330
            End
            Begin VB.Label lbName 
               Caption         =   "Pump"
               Height          =   375
               Index           =   11
               Left            =   0
               TabIndex        =   59
               Top             =   1920
               Width           =   615
            End
            Begin VB.Shape shpPumpLine 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00000000&
               FillColor       =   &H00008000&
               FillStyle       =   2  'Horizontal Line
               Height          =   135
               Index           =   1
               Left            =   495
               Top             =   585
               Width           =   1095
            End
            Begin VB.Shape shpPumpLine 
               BackColor       =   &H00FFFFFF&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00000000&
               FillColor       =   &H00008000&
               FillStyle       =   3  'Vertical Line
               Height          =   615
               Index           =   0
               Left            =   120
               Top             =   840
               Width           =   135
            End
            Begin VB.Shape shpVGH 
               BackStyle       =   1  'Opaque
               Height          =   135
               Left            =   480
               Top             =   1080
               Width           =   135
            End
            Begin VB.Shape shpVGL 
               BackColor       =   &H0000FF00&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00000000&
               Height          =   135
               Left            =   360
               Top             =   1080
               Width           =   135
            End
            Begin VB.Label lbName 
               AutoSize        =   -1  'True
               Caption         =   "VG"
               Height          =   270
               Index           =   50
               Left            =   720
               TabIndex        =   58
               Top             =   960
               Width           =   345
            End
         End
         Begin VB.Label lbChamberNo 
            Caption         =   "腔體編號："
            Height          =   615
            Left            =   120
            TabIndex        =   184
            Top             =   360
            Width           =   1335
         End
         Begin VB.Shape shpGasPipe 
            BackStyle       =   1  'Opaque
            FillColor       =   &H00800000&
            FillStyle       =   2  'Horizontal Line
            Height          =   135
            Index           =   5
            Left            =   6960
            Top             =   7200
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lbMFC 
            Caption         =   "NA"
            Height          =   255
            Index           =   5
            Left            =   7560
            TabIndex        =   87
            Top             =   6600
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lbGasValue 
            Caption         =   "0"
            Height          =   255
            Index           =   5
            Left            =   7440
            TabIndex        =   86
            Top             =   7080
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lbGasValue 
            Caption         =   "0"
            Height          =   255
            Index           =   4
            Left            =   7440
            TabIndex        =   85
            Top             =   6120
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Shape shpGasPipe 
            BackStyle       =   1  'Opaque
            FillColor       =   &H00800000&
            FillStyle       =   2  'Horizontal Line
            Height          =   135
            Index           =   4
            Left            =   6960
            Top             =   6240
            Width           =   375
         End
         Begin VB.Label lbMFC 
            Caption         =   "NA"
            Height          =   255
            Index           =   4
            Left            =   7560
            TabIndex        =   84
            Top             =   5640
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Shape shpLampT 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   0
            Left            =   2880
            Shape           =   4  'Rounded Rectangle
            Top             =   2280
            Width           =   750
         End
         Begin VB.Shape shpLampB 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   0
            Left            =   2880
            Shape           =   4  'Rounded Rectangle
            Top             =   6480
            Width           =   750
         End
         Begin VB.Shape shpCDAPipe 
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   4680
            Shape           =   4  'Rounded Rectangle
            Top             =   1800
            Width           =   135
         End
         Begin VB.Shape shpWaterPipe 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            FillColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   5880
            Top             =   1680
            Width           =   135
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   825
            Index           =   0
            Left            =   240
            Shape           =   4  'Rounded Rectangle
            Top             =   7080
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   1
            Left            =   1080
            Shape           =   4  'Rounded Rectangle
            Top             =   6960
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   9
            Left            =   360
            Shape           =   4  'Rounded Rectangle
            Top             =   6600
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   1065
            Index           =   10
            Left            =   1080
            Shape           =   4  'Rounded Rectangle
            Top             =   7200
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpMainPipe 
            BackStyle       =   1  'Opaque
            FillColor       =   &H00800000&
            FillStyle       =   2  'Horizontal Line
            Height          =   135
            Index           =   1
            Left            =   6720
            Top             =   4440
            Width           =   135
         End
         Begin VB.Shape shpMainPipe 
            BackStyle       =   1  'Opaque
            FillColor       =   &H00800000&
            FillStyle       =   3  'Vertical Line
            Height          =   5295
            Index           =   0
            Left            =   6840
            Top             =   2040
            Width           =   135
         End
         Begin VB.Shape shpGasPipe 
            BackStyle       =   1  'Opaque
            FillColor       =   &H00800000&
            FillStyle       =   2  'Horizontal Line
            Height          =   135
            Index           =   0
            Left            =   6960
            Top             =   2040
            Width           =   375
         End
         Begin VB.Shape shpGasPipe 
            BackStyle       =   1  'Opaque
            FillColor       =   &H00800000&
            FillStyle       =   2  'Horizontal Line
            Height          =   135
            Index           =   1
            Left            =   6960
            Top             =   3120
            Width           =   375
         End
         Begin VB.Label lbMFC 
            Caption         =   "NA"
            Height          =   255
            Index           =   1
            Left            =   7560
            TabIndex        =   80
            Top             =   2520
            Width           =   735
         End
         Begin VB.Label lbMFC 
            Alignment       =   2  'Center
            Caption         =   "NA"
            Height          =   255
            Index           =   0
            Left            =   7320
            TabIndex        =   79
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "CDA"
            Height          =   270
            Index           =   14
            Left            =   4440
            TabIndex        =   78
            Top             =   720
            Width           =   525
         End
         Begin VB.Label lbGasValue 
            Caption         =   "0"
            Height          =   255
            Index           =   0
            Left            =   7440
            TabIndex        =   77
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label lbGasValue 
            Caption         =   "0"
            Height          =   255
            Index           =   1
            Left            =   7440
            TabIndex        =   76
            Top             =   3000
            Width           =   735
         End
         Begin VB.Label lbMFC 
            Caption         =   "NA"
            Height          =   255
            Index           =   3
            Left            =   7560
            TabIndex        =   75
            Top             =   4680
            Width           =   735
         End
         Begin VB.Shape shpGasPipe 
            BackStyle       =   1  'Opaque
            FillColor       =   &H00800000&
            FillStyle       =   2  'Horizontal Line
            Height          =   135
            Index           =   2
            Left            =   6960
            Top             =   4080
            Width           =   375
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   16
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   8880
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   17
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   9240
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   18
            Left            =   480
            Shape           =   4  'Rounded Rectangle
            Top             =   9240
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   19
            Left            =   360
            Shape           =   4  'Rounded Rectangle
            Top             =   6960
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   20
            Left            =   0
            Shape           =   3  'Circle
            Top             =   7680
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   21
            Left            =   0
            Shape           =   3  'Circle
            Top             =   8040
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   22
            Left            =   0
            Shape           =   3  'Circle
            Top             =   8400
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   23
            Left            =   600
            Shape           =   3  'Circle
            Top             =   8880
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   24
            Left            =   840
            Shape           =   3  'Circle
            Top             =   8040
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   25
            Left            =   600
            Shape           =   3  'Circle
            Top             =   8160
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   26
            Left            =   600
            Shape           =   3  'Circle
            Top             =   7800
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   27
            Left            =   600
            Shape           =   3  'Circle
            Top             =   7440
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpDoor 
            BackColor       =   &H00808080&
            BackStyle       =   1  'Opaque
            FillColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3600
            Top             =   6120
            Width           =   1935
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Water"
            Height          =   270
            Index           =   45
            Left            =   5640
            TabIndex        =   74
            Top             =   720
            Width           =   630
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   28
            Left            =   1200
            Shape           =   3  'Circle
            Top             =   7320
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   29
            Left            =   360
            Shape           =   3  'Circle
            Top             =   7680
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   30
            Left            =   600
            Shape           =   3  'Circle
            Top             =   8760
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   31
            Left            =   600
            Shape           =   3  'Circle
            Top             =   8280
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   32
            Left            =   720
            Shape           =   3  'Circle
            Top             =   7800
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   33
            Left            =   480
            Shape           =   3  'Circle
            Top             =   8880
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   34
            Left            =   480
            Shape           =   3  'Circle
            Top             =   8520
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   35
            Left            =   480
            Shape           =   3  'Circle
            Top             =   7920
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   36
            Left            =   480
            Shape           =   3  'Circle
            Top             =   7560
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   37
            Left            =   480
            Shape           =   3  'Circle
            Top             =   7200
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   38
            Left            =   480
            Shape           =   3  'Circle
            Top             =   7080
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   39
            Left            =   360
            Shape           =   3  'Circle
            Top             =   8040
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpGasPipe 
            BackStyle       =   1  'Opaque
            FillColor       =   &H00800000&
            FillStyle       =   2  'Horizontal Line
            Height          =   135
            Index           =   3
            Left            =   6960
            Top             =   5280
            Width           =   375
         End
         Begin VB.Label lbMFC 
            Caption         =   "NA"
            Height          =   255
            Index           =   2
            Left            =   7560
            TabIndex        =   73
            Top             =   3600
            Width           =   735
         End
         Begin VB.Shape shpWafer 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   735
            Index           =   1
            Left            =   4080
            Shape           =   3  'Circle
            Top             =   3480
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Shape shpWafer 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   735
            Index           =   2
            Left            =   4920
            Shape           =   3  'Circle
            Top             =   3480
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Shape shpWafer 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   735
            Index           =   3
            Left            =   3240
            Shape           =   3  'Circle
            Top             =   4320
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Shape shpWafer 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   735
            Index           =   4
            Left            =   4080
            Shape           =   3  'Circle
            Top             =   4320
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Shape shpWafer 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   735
            Index           =   5
            Left            =   4920
            Shape           =   3  'Circle
            Top             =   4320
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Shape shpWafer 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   735
            Index           =   6
            Left            =   3240
            Shape           =   3  'Circle
            Top             =   5160
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Shape shpWafer 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   735
            Index           =   7
            Left            =   4080
            Shape           =   3  'Circle
            Top             =   5160
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Shape shpWafer 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   735
            Index           =   8
            Left            =   4920
            Shape           =   3  'Circle
            Top             =   5160
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lbUSBThermal 
            Alignment       =   2  'Center
            BackColor       =   &H0000FF00&
            Caption         =   "20"
            Height          =   375
            Index           =   0
            Left            =   3960
            TabIndex        =   72
            Top             =   4440
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lbUSBThermal 
            Alignment       =   2  'Center
            BackColor       =   &H0000FF00&
            Caption         =   "20"
            Height          =   375
            Index           =   1
            Left            =   4080
            TabIndex        =   71
            Top             =   4800
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lbUSBThermal 
            Alignment       =   2  'Center
            BackColor       =   &H0000FF00&
            Caption         =   "20"
            Height          =   375
            Index           =   2
            Left            =   4080
            TabIndex        =   70
            Top             =   3600
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lbUSBThermal 
            Alignment       =   2  'Center
            BackColor       =   &H0000FF00&
            Caption         =   "20"
            Height          =   375
            Index           =   3
            Left            =   4680
            TabIndex        =   69
            Top             =   4440
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lbUSBThermal 
            Alignment       =   2  'Center
            BackColor       =   &H0000FF00&
            Caption         =   "20"
            Height          =   375
            Index           =   4
            Left            =   3600
            TabIndex        =   68
            Top             =   4440
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lbTCValue 
            Alignment       =   2  'Center
            Caption         =   "0"
            Height          =   375
            Left            =   4080
            TabIndex        =   67
            Top             =   4800
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lbGasValue 
            Caption         =   "0"
            Height          =   255
            Index           =   3
            Left            =   7440
            TabIndex        =   66
            Top             =   5160
            Width           =   735
         End
         Begin VB.Shape shpWafer 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   2415
            Index           =   0
            Left            =   3240
            Shape           =   3  'Circle
            Top             =   3360
            Width           =   2535
         End
         Begin VB.Line lineExhaust 
            BorderColor     =   &H00808080&
            BorderWidth     =   5
            Index           =   0
            X1              =   3480
            X2              =   3480
            Y1              =   1920
            Y2              =   1320
         End
         Begin VB.Line lineExhaust 
            BorderColor     =   &H00808080&
            BorderWidth     =   5
            Index           =   1
            X1              =   3360
            X2              =   3480
            Y1              =   1560
            Y2              =   1320
         End
         Begin VB.Line lineExhaust 
            BorderColor     =   &H00808080&
            BorderWidth     =   5
            Index           =   2
            X1              =   3600
            X2              =   3480
            Y1              =   1560
            Y2              =   1320
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Exhaust"
            Height          =   270
            Index           =   57
            Left            =   3120
            TabIndex        =   65
            Top             =   720
            Width           =   825
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   3
            Left            =   1080
            Shape           =   4  'Rounded Rectangle
            Top             =   7680
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   4
            Left            =   1080
            Shape           =   4  'Rounded Rectangle
            Top             =   8040
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   5
            Left            =   1080
            Shape           =   4  'Rounded Rectangle
            Top             =   8400
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   2
            Left            =   1080
            Shape           =   4  'Rounded Rectangle
            Top             =   7320
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   825
            Index           =   11
            Left            =   360
            Shape           =   4  'Rounded Rectangle
            Top             =   10200
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   1065
            Index           =   13
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   7800
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   1065
            Index           =   12
            Left            =   600
            Shape           =   4  'Rounded Rectangle
            Top             =   7200
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   6
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   7320
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   7
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   6600
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   8
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   6960
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   1065
            Index           =   14
            Left            =   720
            Shape           =   4  'Rounded Rectangle
            Top             =   7080
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   1065
            Index           =   15
            Left            =   1680
            Shape           =   4  'Rounded Rectangle
            Top             =   7440
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label lbGasValue 
            Caption         =   "0"
            Height          =   255
            Index           =   2
            Left            =   7440
            TabIndex        =   64
            Top             =   4080
            Width           =   735
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   825
            Index           =   40
            Left            =   360
            Shape           =   4  'Rounded Rectangle
            Top             =   9360
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   825
            Index           =   41
            Left            =   240
            Shape           =   4  'Rounded Rectangle
            Top             =   8160
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   825
            Index           =   42
            Left            =   360
            Shape           =   4  'Rounded Rectangle
            Top             =   7680
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   825
            Index           =   43
            Left            =   360
            Shape           =   4  'Rounded Rectangle
            Top             =   6840
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   585
            Index           =   44
            Left            =   480
            Shape           =   4  'Rounded Rectangle
            Top             =   6600
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   45
            Left            =   1920
            Shape           =   4  'Rounded Rectangle
            Top             =   7800
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   46
            Left            =   1920
            Shape           =   4  'Rounded Rectangle
            Top             =   7440
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   825
            Index           =   47
            Left            =   360
            Shape           =   4  'Rounded Rectangle
            Top             =   7440
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   825
            Index           =   48
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   7320
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   825
            Index           =   49
            Left            =   600
            Shape           =   4  'Rounded Rectangle
            Top             =   6840
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   825
            Index           =   50
            Left            =   1200
            Shape           =   4  'Rounded Rectangle
            Top             =   7200
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   825
            Index           =   51
            Left            =   600
            Shape           =   4  'Rounded Rectangle
            Top             =   6720
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLamp 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   585
            Index           =   52
            Left            =   1320
            Shape           =   4  'Rounded Rectangle
            Top             =   6960
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLampL 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   555
            Index           =   0
            Left            =   2400
            Shape           =   4  'Rounded Rectangle
            Top             =   5760
            Width           =   345
         End
         Begin VB.Shape shpLampR 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   555
            Index           =   0
            Left            =   6240
            Shape           =   4  'Rounded Rectangle
            Top             =   5760
            Width           =   345
         End
         Begin VB.Shape shpLampT 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   1
            Left            =   3720
            Shape           =   4  'Rounded Rectangle
            Top             =   2280
            Width           =   750
         End
         Begin VB.Shape shpLampT 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   2
            Left            =   4560
            Shape           =   4  'Rounded Rectangle
            Top             =   2280
            Width           =   750
         End
         Begin VB.Shape shpLampT 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   3
            Left            =   5400
            Shape           =   4  'Rounded Rectangle
            Top             =   2280
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Shape shpLampT 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   4
            Left            =   240
            Shape           =   4  'Rounded Rectangle
            Top             =   3720
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Shape shpLampB 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   1
            Left            =   3720
            Shape           =   4  'Rounded Rectangle
            Top             =   6480
            Width           =   750
         End
         Begin VB.Shape shpLampL 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   555
            Index           =   1
            Left            =   2400
            Shape           =   4  'Rounded Rectangle
            Top             =   5160
            Width           =   345
         End
         Begin VB.Shape shpLampL 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   555
            Index           =   2
            Left            =   2400
            Shape           =   4  'Rounded Rectangle
            Top             =   4560
            Width           =   345
         End
         Begin VB.Shape shpLampL 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   555
            Index           =   3
            Left            =   2400
            Shape           =   4  'Rounded Rectangle
            Top             =   3960
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLampL 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   555
            Index           =   4
            Left            =   2400
            Shape           =   4  'Rounded Rectangle
            Top             =   3360
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLampL 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   555
            Index           =   5
            Left            =   2400
            Shape           =   4  'Rounded Rectangle
            Top             =   2760
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLampR 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   555
            Index           =   1
            Left            =   6240
            Shape           =   4  'Rounded Rectangle
            Top             =   5160
            Width           =   345
         End
         Begin VB.Shape shpLampR 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   555
            Index           =   2
            Left            =   6240
            Shape           =   4  'Rounded Rectangle
            Top             =   4560
            Width           =   345
         End
         Begin VB.Shape shpLampR 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   555
            Index           =   3
            Left            =   6240
            Shape           =   4  'Rounded Rectangle
            Top             =   3960
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLampR 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   555
            Index           =   4
            Left            =   6240
            Shape           =   4  'Rounded Rectangle
            Top             =   3360
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLampR 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   555
            Index           =   5
            Left            =   6240
            Shape           =   4  'Rounded Rectangle
            Top             =   2760
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Shape shpLampB 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   2
            Left            =   4560
            Shape           =   4  'Rounded Rectangle
            Top             =   6480
            Width           =   750
         End
         Begin VB.Shape shpLampB 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   3
            Left            =   5400
            Shape           =   4  'Rounded Rectangle
            Top             =   6480
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Shape shpLampB 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   4
            Left            =   240
            Shape           =   4  'Rounded Rectangle
            Top             =   3600
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Shape shpLampT 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   5
            Left            =   600
            Shape           =   4  'Rounded Rectangle
            Top             =   3360
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Shape shpLampB 
            BorderColor     =   &H000000FF&
            FillColor       =   &H0000C0C0&
            FillStyle       =   0  'Solid
            Height          =   345
            Index           =   5
            Left            =   600
            Shape           =   4  'Rounded Rectangle
            Top             =   3600
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Shape shpTray 
            BackColor       =   &H00C0FFFF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H0000C0C0&
            BorderWidth     =   3
            Height          =   2895
            Left            =   3120
            Top             =   3120
            Width           =   2775
         End
         Begin VB.Shape shpPipeExhaust 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            FillColor       =   &H00808080&
            FillStyle       =   7  'Diagonal Cross
            Height          =   1095
            Left            =   3240
            Top             =   1080
            Width           =   495
         End
         Begin VB.Shape shpChamber 
            BackColor       =   &H00C0C0C0&
            BorderWidth     =   3
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            Height          =   3615
            Left            =   2880
            Shape           =   4  'Rounded Rectangle
            Top             =   2760
            Width           =   3255
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00C0C0C0&
            BorderWidth     =   3
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   4815
            Left            =   2280
            Shape           =   4  'Rounded Rectangle
            Top             =   2160
            Width           =   4455
         End
         Begin VB.Shape shpGas 
            BackColor       =   &H00C0C000&
            BackStyle       =   1  'Opaque
            BorderWidth     =   3
            Height          =   495
            Index           =   0
            Left            =   7320
            Top             =   1800
            Width           =   975
         End
         Begin VB.Shape shpGas 
            BackColor       =   &H00C0C000&
            BackStyle       =   1  'Opaque
            BorderWidth     =   3
            Height          =   495
            Index           =   1
            Left            =   7320
            Top             =   2880
            Width           =   975
         End
         Begin VB.Shape shpGas 
            BackColor       =   &H00C0C000&
            BackStyle       =   1  'Opaque
            BorderWidth     =   3
            Height          =   495
            Index           =   2
            Left            =   7320
            Top             =   3960
            Width           =   975
         End
         Begin VB.Shape shpGas 
            BackColor       =   &H00C0C000&
            BackStyle       =   1  'Opaque
            BorderWidth     =   3
            Height          =   495
            Index           =   3
            Left            =   7320
            Top             =   5040
            Width           =   975
         End
         Begin VB.Shape shpGas 
            BackColor       =   &H00C0C000&
            BackStyle       =   1  'Opaque
            BorderWidth     =   3
            Height          =   495
            Index           =   4
            Left            =   7320
            Top             =   6000
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Shape shpGas 
            BackColor       =   &H00C0C000&
            BackStyle       =   1  'Opaque
            BorderWidth     =   3
            Height          =   495
            Index           =   5
            Left            =   7320
            Top             =   6960
            Visible         =   0   'False
            Width           =   975
         End
      End
      Begin VB.Frame fraCTCheck 
         Caption         =   "燈管檢查"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1095
         Left            =   13680
         TabIndex        =   32
         Top             =   7200
         Visible         =   0   'False
         Width           =   1935
         Begin VB.CommandButton cmdCheckSCR 
            Caption         =   "Go"
            Height          =   615
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Visible         =   0   'False
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "流量控制"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   3975
         Index           =   1
         Left            =   11760
         TabIndex        =   11
         Top             =   2160
         Width           =   3855
         Begin VB.TextBox txtSetGasValue 
            Height          =   390
            Index           =   5
            Left            =   1200
            TabIndex        =   92
            Text            =   "0"
            Top             =   3360
            Width           =   855
         End
         Begin VB.TextBox txtSetGasValue 
            Height          =   390
            Index           =   4
            Left            =   1200
            TabIndex        =   88
            Text            =   "0"
            Top             =   2760
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtSetGasValue 
            Height          =   390
            Index           =   0
            Left            =   1200
            TabIndex        =   16
            Text            =   "0"
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtSetGasValue 
            Height          =   390
            Index           =   1
            Left            =   1200
            TabIndex        =   15
            Text            =   "0"
            Top             =   960
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtSetGasValue 
            Height          =   390
            Index           =   2
            Left            =   1200
            TabIndex        =   14
            Text            =   "0"
            Top             =   1560
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtSetGasValue 
            Height          =   390
            Index           =   3
            Left            =   1200
            TabIndex        =   13
            Text            =   "0"
            Top             =   2160
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox chkGasN2 
            BackColor       =   &H00C0C000&
            Caption         =   "OFF"
            Height          =   400
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   3960
            Value           =   2  'Grayed
            Visible         =   0   'False
            Width           =   2100
         End
         Begin VB.HScrollBar sclSetGasValue 
            Height          =   375
            Index           =   0
            Left            =   960
            Max             =   30
            TabIndex        =   17
            Top             =   360
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.HScrollBar sclSetGasValue 
            Height          =   375
            Index           =   1
            Left            =   960
            Max             =   1000
            TabIndex        =   18
            Top             =   960
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.HScrollBar sclSetGasValue 
            Height          =   375
            Index           =   2
            Left            =   960
            Max             =   1000
            TabIndex        =   19
            Top             =   1560
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.HScrollBar sclSetGasValue 
            Height          =   375
            Index           =   3
            Left            =   960
            Max             =   30
            TabIndex        =   20
            Top             =   2160
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.HScrollBar sclSetGasValue 
            Height          =   375
            Index           =   4
            Left            =   960
            Max             =   30
            TabIndex        =   89
            Top             =   2760
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.HScrollBar sclSetGasValue 
            Height          =   375
            Index           =   5
            Left            =   960
            Max             =   30
            TabIndex        =   93
            Top             =   3360
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lbGasName 
            AutoSize        =   -1  'True
            Caption         =   "NA"
            Height          =   270
            Index           =   5
            Left            =   240
            TabIndex        =   95
            Top             =   3360
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Label lbGasUnit 
            AutoSize        =   -1  'True
            Caption         =   "SLPM"
            Height          =   270
            Index           =   5
            Left            =   2640
            TabIndex        =   94
            Top             =   3360
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.Label lbGasName 
            AutoSize        =   -1  'True
            Caption         =   "NA"
            Height          =   270
            Index           =   4
            Left            =   240
            TabIndex        =   91
            Top             =   2760
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Label lbGasUnit 
            AutoSize        =   -1  'True
            Caption         =   "SLPM"
            Height          =   270
            Index           =   4
            Left            =   2640
            TabIndex        =   90
            Top             =   2760
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.Label lbGasName 
            AutoSize        =   -1  'True
            Caption         =   "NA"
            Height          =   270
            Index           =   0
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Label lbGasName 
            AutoSize        =   -1  'True
            Caption         =   "NA"
            Height          =   270
            Index           =   1
            Left            =   240
            TabIndex        =   27
            Top             =   960
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Label lbGasUnit 
            AutoSize        =   -1  'True
            Caption         =   "SLPM"
            Height          =   270
            Index           =   0
            Left            =   2640
            TabIndex        =   26
            Top             =   360
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.Label lbGasUnit 
            AutoSize        =   -1  'True
            Caption         =   "SLPM"
            Height          =   270
            Index           =   1
            Left            =   2640
            TabIndex        =   25
            Top             =   960
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.Label lbGasUnit 
            AutoSize        =   -1  'True
            Caption         =   "SLPM"
            Height          =   270
            Index           =   3
            Left            =   2640
            TabIndex        =   24
            Top             =   2160
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.Label lbGasUnit 
            AutoSize        =   -1  'True
            Caption         =   "SLPM"
            Height          =   270
            Index           =   2
            Left            =   2640
            TabIndex        =   23
            Top             =   1560
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.Label lbGasName 
            AutoSize        =   -1  'True
            Caption         =   "NA"
            Height          =   270
            Index           =   2
            Left            =   240
            TabIndex        =   22
            Top             =   1560
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Label lbGasName 
            AutoSize        =   -1  'True
            Caption         =   "NA"
            Height          =   270
            Index           =   3
            Left            =   240
            TabIndex        =   21
            Top             =   2160
            Visible         =   0   'False
            Width           =   330
         End
      End
      Begin VB.Frame fraIntensity 
         Caption         =   "能量輸出控制"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1935
         Left            =   11760
         TabIndex        =   1
         Top             =   120
         Width           =   3855
         Begin VB.TextBox txtSCR_Output 
            Height          =   390
            Index           =   0
            Left            =   1080
            TabIndex        =   5
            Text            =   "20"
            Top             =   840
            Width           =   975
         End
         Begin VB.ComboBox cmbScrList 
            Height          =   390
            ItemData        =   "frmDiagnosis.frx":003C
            Left            =   1080
            List            =   "frmDiagnosis.frx":003E
            TabIndex        =   4
            Text            =   "SCR-01"
            Top             =   360
            Width           =   1335
         End
         Begin VB.CheckBox chkScrTest 
            Caption         =   "開啟測試"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtSCR_Output 
            Height          =   390
            Index           =   10
            Left            =   1080
            TabIndex        =   2
            Text            =   "5"
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "SCR"
            Height          =   270
            Index           =   2
            Left            =   360
            TabIndex        =   10
            Top             =   360
            Width           =   510
         End
         Begin VB.Label lbName 
            Caption         =   "%"
            Height          =   255
            Index           =   29
            Left            =   2160
            TabIndex        =   9
            Top             =   840
            Width           =   255
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Power"
            Height          =   270
            Index           =   30
            Left            =   240
            TabIndex        =   8
            Top             =   840
            Width           =   675
         End
         Begin VB.Label lbName 
            AutoSize        =   -1  'True
            Caption         =   "Keep"
            Height          =   270
            Index           =   6
            Left            =   240
            TabIndex        =   7
            Top             =   1320
            Width           =   570
         End
         Begin VB.Label lbName 
            Caption         =   "Sec"
            Height          =   255
            Index           =   13
            Left            =   2040
            TabIndex        =   6
            Top             =   1320
            Width           =   495
         End
      End
   End
   Begin VB.Label lbMFC 
      Caption         =   "NA"
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   186
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmDiagnosis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public intPreHeatCount As Integer


Private Sub Command1_Click()
    
    
End Sub

Private Sub Command2_Click()
    
End Sub




Private Sub Form_Activate()
    Dim i As Integer
    
    If gbintActiveModule_Vacuum = 1 Then
        fraVacFunc.Visible = True
        fraVacuum.Visible = True
        
    Else
        fraVacFunc.Visible = False
        fraVacuum.Visible = False
        
    End If
    fraOxygen.Visible = IIf(gbintActiveModule_Oxygen = 1, True, False)
    
    
    For i = 0 To 7
        If gbintMonitorTCActive(i) = 1 Then
           lbNameMonitorTC(i).Caption = gbstrNameTC(i)
            lbNameMonitorTC(i).Visible = True
            lbTC(i).Visible = True
        Else
            lbNameMonitorTC(i).Visible = False
            lbTC(i).Visible = False
        End If
    Next i
     
         
          
          
'    If gbintCTCheck = 1 Then
'        fraCTCheck.Visible = True
'    Else
'        fraCTCheck.Visible = False
'    End If
    
    
    For i = 0 To 5
        shpLampL(i).Left = 2400
        shpLampR(i).Left = 6240
        shpLampT(i).Top = 2280
        shpLampB(i).Top = 6480
    Next i
    
    If gbintNumOfBanks = 10 Then
        
            
        
        For i = 3 To 5
            shpLampL(i).Visible = True
            shpLampR(i).Visible = True
        Next i
        shpLampT(3).Visible = True
        shpLampB(3).Visible = True
        
        shpLampL(0).Top = 5760
        shpLampR(0).Top = 5760
                
        shpLampL(1).Top = 5160
        shpLampR(1).Top = 5160
                
        shpLampL(2).Top = 4560
        shpLampR(2).Top = 4560
        
        shpLampL(3).Top = 3960
        shpLampR(3).Top = 3960
        
        shpLampL(4).Top = 3360
        shpLampR(4).Top = 3360
        
        shpLampL(5).Top = 2760
        shpLampR(5).Top = 2760
        
        shpLampT(0).Left = 2880
        shpLampB(0).Left = 2880
        
        shpLampT(1).Left = 3720
        shpLampB(1).Left = 3720
                
        shpLampT(2).Left = 4560
        shpLampB(2).Left = 4560
        
        shpLampT(3).Left = 5400
        shpLampB(3).Left = 5400
    Else
        
        
        
        
        For i = 3 To 5
            shpLampL(i).Visible = False
            shpLampR(i).Visible = False
            shpLampT(i).Visible = False
            shpLampB(i).Visible = False
        Next i
        
        
        shpLampL(0).Top = 4200
        shpLampR(0).Top = 4200
                
        shpLampL(1).Top = 5280
        shpLampR(1).Top = 5280
                
        shpLampL(2).Top = 3240
        shpLampR(2).Top = 3240
        
        shpLampT(0).Left = 4080
        shpLampB(0).Left = 4080
        
        shpLampT(1).Left = 5040
        shpLampB(1).Left = 5040
                
        shpLampT(2).Left = 3120
        shpLampB(2).Left = 3120
    End If
    tabMain.TabVisible(0) = True
    tabMain.TabVisible(1) = False
'    If gbintRtaType = 3 Then
'        tabMain.TabVisible(0) = False
'        tabMain.TabVisible(1) = True
'    Else
'        tabMain.TabVisible(0) = True
'        tabMain.TabVisible(1) = False
'    End If
    
    If Para.UseCT = 0 Then
        fraBankH.Visible = False
        
    Else
        fraBankH.Visible = True
        
    End If
    
    fraTurbo.Visible = IIf(Para.useTPump = 1, True, False)
    fraCover.Visible = IIf(Para.UseCover = 1, True, False)
    
End Sub

Private Sub Form_Load()
    Dim j As Integer
    Dim i As Integer
    
'    For j = 0 To gbintMaxGasEnable - 1
      For j = 0 To UBound(gbstrGasAlias) - 1
        If gbstrGasAlias(j) <> "NA" And gbstrGasAlias(j) <> "Pump" Then
            lbMFC(j).Visible = True
            shpGas(j).Visible = True
            lbGasValue(j).Visible = True
            shpGasPipe(j).Visible = True
            lbGasName(j).Visible = True
            txtSetGasValue(j).Visible = True
            sclSetGasValue(j).Visible = True
            lbGasUnit(j).Visible = True
            
            lbGasName(j).Caption = gbstrGasAlias(j)
            lbGasUnit(j).Caption = gbstrGasUnit(j)
            gbsngGasFlowScale(j) = CheckValueScale(CStr(gbsngMaxGasSLMP(j)))
            sclSetGasValue(j).Max = gbsngMaxGasSLMP(j) * gbsngGasFlowScale(j)
        Else
            shpGas(j).Visible = False
            lbGasValue(j).Visible = False
            shpGasPipe(j).Visible = False
            txtSetGasValue(j).Visible = False
            sclSetGasValue(j).Visible = False
            lbGasName(j).Visible = False
            lbGasUnit(j).Visible = False
        End If

    Next j
    'lbGasValue(2).Visible = False
    Call RefreshGasDefiniation
    
    If gbintActiveModule_Vacuum = 1 Then
        fraVacFunc.Visible = True
    Else
        fraVacFunc.Visible = False
    End If
    
    cmbScrList.AddItem "SCR-01"
    Select Case gbintNumOfBanks
        Case 5
            cmbScrList.AddItem "SCR-02"
            cmbScrList.AddItem "SCR-03"
            cmbScrList.AddItem "SCR-04"
            cmbScrList.AddItem "SCR-05"
        Case 6
            cmbScrList.AddItem "SCR-02"
            cmbScrList.AddItem "SCR-03"
            cmbScrList.AddItem "SCR-04"
            cmbScrList.AddItem "SCR-05"
            cmbScrList.AddItem "SCR-06"
        Case 10
            cmbScrList.AddItem "SCR-02"
            cmbScrList.AddItem "SCR-03"
            cmbScrList.AddItem "SCR-04"
            cmbScrList.AddItem "SCR-05"
            cmbScrList.AddItem "SCR-06"
            cmbScrList.AddItem "SCR-07"
            cmbScrList.AddItem "SCR-08"
            cmbScrList.AddItem "SCR-09"
            cmbScrList.AddItem "SCR-10"
        Case 12
            cmbScrList.AddItem "SCR-02"
            cmbScrList.AddItem "SCR-03"
            cmbScrList.AddItem "SCR-04"
            cmbScrList.AddItem "SCR-05"
            cmbScrList.AddItem "SCR-06"
            cmbScrList.AddItem "SCR-07"
            cmbScrList.AddItem "SCR-08"
            cmbScrList.AddItem "SCR-09"
            cmbScrList.AddItem "SCR-10"
            cmbScrList.AddItem "SCR-11"
            cmbScrList.AddItem "SCR-12"
        Case 13
            cmbScrList.AddItem "SCR-02"
            cmbScrList.AddItem "SCR-03"
            cmbScrList.AddItem "SCR-04"
            cmbScrList.AddItem "SCR-05"
            cmbScrList.AddItem "SCR-06"
            cmbScrList.AddItem "SCR-07"
            cmbScrList.AddItem "SCR-08"
            cmbScrList.AddItem "SCR-09"
            cmbScrList.AddItem "SCR-10"
            cmbScrList.AddItem "SCR-11"
            cmbScrList.AddItem "SCR-12"
            cmbScrList.AddItem "SCR-13"
            cmbScrList.AddItem "SCR-14"
        Case 14
            cmbScrList.AddItem "SCR-02"
            cmbScrList.AddItem "SCR-03"
            cmbScrList.AddItem "SCR-04"
            cmbScrList.AddItem "SCR-05"
            cmbScrList.AddItem "SCR-06"
            cmbScrList.AddItem "SCR-07"
            cmbScrList.AddItem "SCR-08"
            cmbScrList.AddItem "SCR-09"
            cmbScrList.AddItem "SCR-10"
            cmbScrList.AddItem "SCR-11"
            cmbScrList.AddItem "SCR-12"
            cmbScrList.AddItem "SCR-13"
            cmbScrList.AddItem "SCR-14"
        Case 17
            cmbScrList.AddItem "SCR-02"
            cmbScrList.AddItem "SCR-03"
            cmbScrList.AddItem "SCR-04"
            cmbScrList.AddItem "SCR-05"
            cmbScrList.AddItem "SCR-06"
            cmbScrList.AddItem "SCR-07"
            cmbScrList.AddItem "SCR-08"
            cmbScrList.AddItem "SCR-09"
            cmbScrList.AddItem "SCR-10"
            cmbScrList.AddItem "SCR-11"
            cmbScrList.AddItem "SCR-12"
            cmbScrList.AddItem "SCR-13"
            cmbScrList.AddItem "SCR-14"
            cmbScrList.AddItem "SCR-15"
            cmbScrList.AddItem "SCR-16"
            cmbScrList.AddItem "SCR-17"
    End Select
    cmbScrList.AddItem "SCR-ALL"
    cmbScrList.ListIndex = 0
    
    If Para.UseAz1 = False Then
        fraIntensity.Visible = True
        fraAZ.Visible = False
    Else
        fraIntensity.Visible = False
        fraAZ.Visible = True
        
        If Para.UseAz1 Then
            cmbBlkList.AddItem "BLK-01"
            cmbBlkList.AddItem "BLK-02"
            cmbBlkList.AddItem "BLK-03"
            cmbBlkList.AddItem "BLK-04"
        End If
        If Para.UseAz2 Then
            cmbBlkList.AddItem "BLK-05"
            cmbBlkList.AddItem "BLK-06"
            cmbBlkList.AddItem "BLK-07"
            cmbBlkList.AddItem "BLK-08"
        End If
        cmbBlkList.ListIndex = 0
        
    End If
  
   TxtChamberNo.text = CommnonReadini("DeviceInfo", "ChamberNo", App.Path + DeviceInfo_Path)
   gbChamberNo = TxtChamberNo.text
   If GbChamberNo_Switch = 1 Then
   mdifrmRTP.tbrRTP.Buttons(14).Caption = "腔體編號:" & gbChamberNo
   End If
   
 
   If CTDisplay = 1 Then
        hfgCTConfigProcess.Visible = True
        ShowCTTable
        
        For i = 0 To 59
            lbCT(i).Visible = False
        Next i
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmConfiguration.StopWatchDog
End Sub

Private Sub chkTurbo_Click()
    If chkTurbo.value = 1 Then
        SetAngle True
        Call frmHistory.AppendLogAlert(1, "Manual", 1015, "手動開啟Turbo", 1)
    Else
        SetAngle False
        Call frmHistory.AppendLogAlert(1, "Manual", 1016, "手動關閉Turbo", 1)
    End If
End Sub


Public Sub chkPumping_Click()
'    If Para.useTPump = 0 Then
'        If chkPumping.value = 1 Then
'            If SysDI.IsChamberGaugeL = 0 Then
'                chkPumping.value = 0
'                ShowMessageOK "腔體處於真空狀態,請先釋放壓力!"
'                Exit Sub
'            End If
'
'            SetAngle True
'            frmDiagnosis.tmrPumpON.Enabled = True
'            Call frmHistory.AppendLogAlert(1, "Manual", 1012, "手動開啟泵浦", 1)
'        Else
'            SetAngle False
'            frmDiagnosis.tmrPumpON.Enabled = False
'            frmDiagnosis.tmrPumpOFF.Enabled = True
'            Call frmHistory.AppendLogAlert(1, "Manual", 1013, "手動關閉泵浦", 1)
'        End If
'    Else
'        If chkPumping.value = 1 Then
'            If SysDI.IsChamberGaugeL = 0 Then
'                chkPumping.value = 0
'                ShowMessageOK "腔體處於真空狀態,請先釋放壓力!"
'                Exit Sub
'            End If
'
'            SetPump True
'            Call frmHistory.AppendLogAlert(1, "Manual", 1012, "手動開啟泵浦", 1)
'        Else
'            SetPump False
'            Call frmHistory.AppendLogAlert(1, "Manual", 1013, "手動關閉泵浦", 1)
'            If gbintReleaseOpenDelay > 0 Then tmrSetReleaseOFF.Enabled = True
'        End If
'    End If
    Dim countdown As Long
    Dim isMessageShown As Boolean
  
    isMessageShown = False
     If gbintReleaseOpenDelay > 0 Then
        Do
            DoEvents
                If tmrSetReleaseOFF.Enabled Then
                    If Not isMessageShown Then
                        countdown = tmrSetReleaseOFF.Interval / 1000
                        ShowMessageOK "下一次操作將在" & countdown & "秒后進行！"
                        isMessageShown = True
                        chkPumping.value = IIf(PumpState, 1, 0)
                    End If
                    Sleep (100)
                Else
                    Exit Do
                    
                End If
            Loop
    End If
    If Not tmrSetReleaseOFF.Enabled Then
        Call TogglePumping(chkPumping, Para.useTPump = 1, frmDiagnosis.tmrSetReleaseOFF)
        frmPlotProcess.chkPumping.value = IIf(PumpState, 1, 0)
    End If
    
End Sub

Public Sub chkPurge_Click()

    
    Dim sngGas(10) As Single
    Dim i As Integer
    
    'On Error GoTo ERRLINE
    
    If chkPurge.value = 1 Then
        If gbsngKeepPurge > 0 Then
            txtSetGasValue(0).text = CStr(gbsngKeepPurge)
            For i = 0 To 4
                sngGas(i) = Val(txtSetGasValue(i).text)
            Next i
            SetAO_MFC sngGas
            
            tmrPurge.Enabled = True
            Kernel.IsPurge = 1
            Call frmHistory.AppendLogAlert(1, "Manual", 1015, "手動開啟破真空", 1)
            If gbintLoginRight > 1 Then Call mdifrmRTP.ShowTitleBar(False)
            frmPlotProcess.chkPurge.value = 1
        End If
        
    Else
        tmrPurge.Enabled = False
        Kernel.IsPurge = 0
        txtSetGasValue(0).text = 0
        For i = 0 To 4
                sngGas(i) = 0
        Next i
        SetAO_MFC sngGas
        
        Call mdifrmRTP.ShowTitleBar(True)
        frmPlotProcess.chkPurge.value = 0
    End If
    

           
End Sub

Private Sub chkBlkTest_Click()
    Dim Index As Integer
    Dim iValue As Integer
    
    Index = cmbBlkList.ListIndex + 1
    iValue = Val(txtBLK_Output.text)
        
    If iValue >= 0 And Index > 0 And chkBlkTest.value = 1 Then
                        
        If Index > 0 And Index <= 4 Then
            Call frmAz1.WritePara(112 + Index, 1)
            Call frmAz1.WritePara(116 + Index, iValue)
        ElseIf Index >= 5 And Index <= 8 Then
            Call frmAz2.WritePara(112 + Index - 4, 1)
            Call frmAz2.WritePara(116 + Index - 4, iValue)
        End If
        Call SetLampCooling(True)
                
        tmrPowerTest.Interval = Val(txtBLK_Keep.text) * 1000
        tmrPowerTest.Enabled = True
    Else
        tmrPowerTest.Enabled = False
        If Index <= 4 Then
            Call frmAz1.WritePara(116 + Index, 0)
        ElseIf Index >= 5 And Index <= 8 Then
            Call frmAz2.WritePara(116 + Index - 4, 0)
        End If
        Call SetLampCooling(False)
    End If
    
End Sub

Private Sub chkScrTest_Click()
    Dim i As Integer
    Dim iRet As Integer
    Dim iValue As Integer
    Dim sngValue As Single
    Dim lngSCR_AOChannel(GB_SCR_MAX) As Long
    
    lngSCR_AOChannel(0) = gblngAO_SCR_TBC
    lngSCR_AOChannel(1) = gblngAO_SCR_TR
    lngSCR_AOChannel(2) = gblngAO_SCR_TL
    lngSCR_AOChannel(3) = gblngAO_SCR_BF
    lngSCR_AOChannel(4) = gblngAO_SCR_BR
    '120713 Josh
    lngSCR_AOChannel(5) = gblngAO_SCR_6
    lngSCR_AOChannel(6) = gblngAO_SCR_7
    lngSCR_AOChannel(7) = gblngAO_SCR_8
    lngSCR_AOChannel(8) = gblngAO_SCR_9
    lngSCR_AOChannel(9) = gblngAO_SCR_10
    lngSCR_AOChannel(10) = gblngAO_SCR_11
    lngSCR_AOChannel(11) = gblngAO_SCR_12
    lngSCR_AOChannel(12) = gblngAO_SCR_13
    lngSCR_AOChannel(13) = gblngAO_SCR_14
    lngSCR_AOChannel(14) = gblngAO_SCR_15
    lngSCR_AOChannel(15) = gblngAO_SCR_16
    lngSCR_AOChannel(16) = gblngAO_SCR_17
    iValue = cmbScrList.ListIndex
    
    
    If iValue >= 0 And chkScrTest.value = 1 Then
        sngValue = Val(txtSCR_Output(0).text)
        If txtSCR_Output(0).text = "" Then txtSCR_Output(0).text = "0"
        If Val(txtSCR_Output(0).text) > 100 Then txtSCR_Output(0).text = "10"
        If Val(txtSCR_Output(0).text) < 0 Then txtSCR_Output(0).text = "0"
        If Not CheckStringIsNumber(txtSCR_Output(0).text) Then MsgBox "Keyin Error!": Exit Sub
                
       
        If iValue < (cmbScrList.ListCount - 1) Then
            SetAO lngSCR_AOChannel(iValue), sngValue
        Else
            For i = 0 To gbintNumOfBanks - 1
                If lngSCR_AOChannel(i) >= 0 Then
                    SetAO lngSCR_AOChannel(i), sngValue
                End If
            Next i
        End If
        
        
        Call SetLampCooling(True)
                
        tmrPowerTest.Interval = Val(txtSCR_Output(10).text) * 1000
        tmrPowerTest.Enabled = True
    Else
        tmrPowerTest.Enabled = False
        For i = 0 To gbintNumOfBanks - 1
            SetAO lngSCR_AOChannel(i), 0
        Next i
        
        Call SetLampCooling(False)
    End If
      
   
End Sub

Public Sub RunPreheat(intPower As Integer)
    Dim i As Integer
    Dim lngSCR_AOChannel(GB_SCR_MAX) As Long
        
    If Kernel.IsRun = 0 And Kernel.IsAlarm = 0 Then
        If Para.RtaType = 9 Then
            If Para.UseAz1 = 1 Then
                For i = 1 To 4
                    Call frmAz1.WritePara(112 + i, 1)
                    Call frmAz1.WritePara(116 + i, intPower)
                Next i
            End If
            If Para.UseAz2 = 1 Then
                For i = 1 To 4
                    Call frmAz2.WritePara(112 + i, 1)
                    Call frmAz2.WritePara(116 + i, intPower)
                Next i
            End If
        Else
            lngSCR_AOChannel(0) = gblngAO_SCR_TBC
            lngSCR_AOChannel(1) = gblngAO_SCR_TR
            lngSCR_AOChannel(2) = gblngAO_SCR_TL
            lngSCR_AOChannel(3) = gblngAO_SCR_BF
            lngSCR_AOChannel(4) = gblngAO_SCR_BR
            '120713 Josh
            lngSCR_AOChannel(5) = gblngAO_SCR_6
            lngSCR_AOChannel(6) = gblngAO_SCR_7
            lngSCR_AOChannel(7) = gblngAO_SCR_8
            lngSCR_AOChannel(8) = gblngAO_SCR_9
            lngSCR_AOChannel(9) = gblngAO_SCR_10
            lngSCR_AOChannel(10) = gblngAO_SCR_11
            lngSCR_AOChannel(11) = gblngAO_SCR_12
            lngSCR_AOChannel(12) = gblngAO_SCR_13
            lngSCR_AOChannel(13) = gblngAO_SCR_14
            lngSCR_AOChannel(14) = gblngAO_SCR_15
            lngSCR_AOChannel(15) = gblngAO_SCR_16
            lngSCR_AOChannel(16) = gblngAO_SCR_17
            
            For i = 0 To gbintNumOfBanks - 1
                If lngSCR_AOChannel(i) >= 0 Then
                    SetAO lngSCR_AOChannel(i), Val(intPower)
                End If
            Next i
        End If
        
        If intPower > 0 Then
            Call SetLampCooling(True)
            intPreHeatCount = 0
            Kernel.IsPreHeat = 1
            tmrPreHeat.Enabled = True
        Else
            Call SetLampCooling(False)
            intPreHeatCount = 0
            Kernel.IsPreHeat = 0
            tmrPreHeat.Enabled = False
        End If
    End If
    
End Sub

Private Sub tmrCheckMFC_Timer(Index As Integer)
    Dim sngTemp(3) As Single
    Dim strTemp As String
    
    If Kernel.sngCurrOutMFC(Index) > 0 Then
        sngTemp(0) = Kernel.sngCurrOutMFC(Index) + (Kernel.sngCurrOutMFC(Index) * gbsngGasError(Index) / 100)
        sngTemp(1) = Kernel.sngCurrOutMFC(Index) - (Kernel.sngCurrOutMFC(Index) * gbsngGasError(Index) / 100)
        If (SysAI.sngMFC(Index) > sngTemp(0)) Or (SysAI.sngMFC(Index) < sngTemp(1)) Then
            gbsngGasErrorC(Index) = gbsngGasErrorC(Index) + 1
            If gbsngGasErrorC(Index) > gbsngGasErrorN(Index) Then
                tmrCheckMFC(Index).Enabled = False
                gbstrAlarmHint = gbstrGasAlias(Index)
                ShowAlarmFlash 19
                Call frmHistory.AppendLogAlert(1, "Alarm", 3034, "MFC 流量異常=" & gbstrAlarmHint, 1)
            End If
        Else
            gbsngGasErrorC(Index) = 0
            tmrCheckMFC(Index).Enabled = False
        End If
    Else
        gbsngGasErrorC(Index) = 0
        tmrCheckMFC(Index).Enabled = False
    End If
End Sub



Private Sub tmrHoldSafeON_Timer()
    tmrHoldSafeON.Enabled = False
    SetDO gblngDO_ARM_FRONT, True
End Sub

Private Sub tmrPowerTest_Timer()
    Dim i As Integer
    Dim lngSCR_AOChannel(GB_SCR_MAX) As Long
    Dim Index As Integer

    tmrPowerTest.Enabled = False
    Call SetLampCooling(False)
    
    If Para.RtaType = 9 Then
        
        Index = cmbBlkList.ListIndex + 1
        If Index <= 4 Then
            Call frmAz1.WritePara(112 + Index, 0)
            Call frmAz1.WritePara(116 + Index, 0)
        ElseIf Index >= 5 And Index <= 8 Then
            Call frmAz2.WritePara(112 + Index - 4, 0)
            Call frmAz2.WritePara(116 + Index - 4, 0)
        End If
              
        chkBlkTest.value = 0
    Else
    
    
        lngSCR_AOChannel(0) = gblngAO_SCR_TBC
        lngSCR_AOChannel(1) = gblngAO_SCR_TR
        lngSCR_AOChannel(2) = gblngAO_SCR_TL
        lngSCR_AOChannel(3) = gblngAO_SCR_BF
        lngSCR_AOChannel(4) = gblngAO_SCR_BR
        '120713 Josh
        lngSCR_AOChannel(5) = gblngAO_SCR_6
        lngSCR_AOChannel(6) = gblngAO_SCR_7
        lngSCR_AOChannel(7) = gblngAO_SCR_8
        lngSCR_AOChannel(8) = gblngAO_SCR_9
        lngSCR_AOChannel(9) = gblngAO_SCR_10
        lngSCR_AOChannel(10) = gblngAO_SCR_11
        lngSCR_AOChannel(11) = gblngAO_SCR_12
        lngSCR_AOChannel(12) = gblngAO_SCR_13
        lngSCR_AOChannel(13) = gblngAO_SCR_14
        lngSCR_AOChannel(14) = gblngAO_SCR_15
        lngSCR_AOChannel(15) = gblngAO_SCR_16
        lngSCR_AOChannel(16) = gblngAO_SCR_17
    
        For i = 0 To gbintNumOfBanks - 1
            If lngSCR_AOChannel(i) >= 0 Then
                SetAO lngSCR_AOChannel(i), 0
                'Call ShowLampStatus(i, 0)
            End If
        Next i
        
        
        chkScrTest.value = 0
    End If
        
End Sub

Private Sub tmrPreHeat_Timer()
    Dim i As Integer
    Dim lngSCR_AOChannel(GB_SCR_MAX) As Long
    
    intPreHeatCount = intPreHeatCount + 1
    If intPreHeatCount > gbintPreheatTime Then
        tmrPreHeat.Enabled = False
        Call SetLampCooling(False)
        
        If Para.RtaType = 9 Then
            If Para.UseAz1 = 1 Then
                For i = 1 To 4
                    Call frmAz1.WritePara(112 + i, 0)
                    Call frmAz1.WritePara(116 + i, 0)
                Next i
            End If
            If Para.UseAz2 = 1 Then
                For i = 1 To 4
                    Call frmAz2.WritePara(112 + i, 0)
                    Call frmAz2.WritePara(116 + i, 0)
                Next i
            End If
        Else
            lngSCR_AOChannel(0) = gblngAO_SCR_TBC
            lngSCR_AOChannel(1) = gblngAO_SCR_TR
            lngSCR_AOChannel(2) = gblngAO_SCR_TL
            lngSCR_AOChannel(3) = gblngAO_SCR_BF
            lngSCR_AOChannel(4) = gblngAO_SCR_BR
            '120713 Josh
            lngSCR_AOChannel(5) = gblngAO_SCR_6
            lngSCR_AOChannel(6) = gblngAO_SCR_7
            lngSCR_AOChannel(7) = gblngAO_SCR_8
            lngSCR_AOChannel(8) = gblngAO_SCR_9
            lngSCR_AOChannel(9) = gblngAO_SCR_10
            lngSCR_AOChannel(10) = gblngAO_SCR_11
            lngSCR_AOChannel(11) = gblngAO_SCR_12
            lngSCR_AOChannel(12) = gblngAO_SCR_13
            lngSCR_AOChannel(13) = gblngAO_SCR_14
            lngSCR_AOChannel(14) = gblngAO_SCR_15
            lngSCR_AOChannel(15) = gblngAO_SCR_16
            lngSCR_AOChannel(16) = gblngAO_SCR_17
        
            For i = 0 To gbintNumOfBanks - 1
                If lngSCR_AOChannel(i) >= 0 Then
                    SetAO lngSCR_AOChannel(i), 0
                End If
            Next i
        End If
    End If
End Sub

Private Sub tmrPumpON_Timer()
    tmrPumpON.Enabled = False
    SetPump True
    If gbintReleaseOpenDelay > 0 Then
'        Call frmHistory.AppendLogAlert(1, "Manual", 1013, "自動關閉洩壓閥", 1)
         Call frmHistory.AppendLogAlert(1, "Process", 1013, "自動關閉洩壓閥", 1)
        SetRelease True
    End If
    gtcPump.SHOWMODE = 2
End Sub

Private Sub tmrPumpOFF_Timer()
    tmrPumpOFF.Enabled = False
    SetPump False
    If gbintReleaseOpenDelay > 0 Then
        tmrSetReleaseOFF.Interval = gbintReleaseOpenDelay
        tmrSetReleaseOFF.Enabled = True
    End If
    gtcPump.SHOWMODE = 0
End Sub

Private Sub tmrPurge_Timer()
    If SysDI.IsChamberGaugeL = 1 Then
        tmrPurge.Enabled = False
        DelayTime (3)
        chkPurge.value = 0
        Kernel.IsPurge = 0
        frmPlotProcess.chkPurge.value = 0
        txtSetGasValue(0).text = 0
        sclSetGasValue(0).value = 0
        Call frmHistory.AppendLogAlert(1, "Manual", 1016, "腔體已回壓至ATM", 1)
        If Para.UseAutoMode = 0 Then
            If gbintActiveModule_Door > 0 Then
                QuestionAns = vbNo
                ShowMessageYN "是否要開啟腔門?"
                Call mdifrmRTP.ShowTitleBar(True)
                If QuestionAns = vbYes Then
                    SetDoor 1
                    SetDoor 2
                End If
            End If
        Else
            frmUDP.wsServer.SendData "#PU=0,"
        End If
    End If
End Sub

Private Sub tmrSetReleaseOFF_Timer()
    tmrSetReleaseOFF.Enabled = False
'    Call frmHistory.AppendLogAlert(1, "Manual", 1013, "自動開啟洩壓閥", 1)
     Call frmHistory.AppendLogAlert(1, "Process", 1013, "自動開啟洩壓閥", 1)
    SetRelease False
End Sub

Private Sub cmdCheckSCR_Click()
    cmdCheckSCR.Enabled = False
    Call ControlSCR(2, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100)
    DelayTime (2)
    Call ControlSCR(0, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 100)
    cmdCheckSCR.Enabled = True
End Sub

Private Sub sclSetGasValue_Change(Index As Integer)
'
'    Dim i As Integer
'    Dim j As Integer
'    Dim iCnt As Integer
'
'    Dim sngGas(GB_GAS_MAX) As Single
'
'    iCnt = 0
'    For i = 0 To gbintMaxGasEnable
'        If (gbstrGasAlias(i) = "O2" And sclSetGasValue(i).value > 0) Then
'            iCnt = iCnt + 1
'        End If
'    Next i
'    If iCnt > 0 Then
'        For i = 0 To gbintMaxGasEnable
'            If (gbstrGasAlias(i) = "H2") Then
'                txtSetGasValue(i).Text = "0"
'                txtSetGasValue(i).Enabled = False
'                sclSetGasValue(i).Enabled = False
'            End If
'        Next i
'    Else
'        For i = 0 To gbintMaxGasEnable
'            If (gbstrGasAlias(i) = "H2") Then
'                txtSetGasValue(i).Enabled = True
'                sclSetGasValue(i).Enabled = True
'            End If
'        Next i
'    End If
'
'    iCnt = 0
'    For i = 0 To gbintMaxGasEnable
'        If (gbstrGasAlias(i) = "H2" And sclSetGasValue(i).value > 0) Then
'            iCnt = iCnt + 1
'        End If
'    Next i
'    If iCnt > 0 Then
'        For i = 0 To gbintMaxGasEnable
'            If (gbstrGasAlias(i) = "O2") Then
'                txtSetGasValue(i).Text = "0"
'                txtSetGasValue(i).Enabled = False
'                sclSetGasValue(i).Enabled = False
'            End If
'        Next i
'    Else
'        For i = 0 To gbintMaxGasEnable
'            If (gbstrGasAlias(i) = "O2") Then
'                txtSetGasValue(i).Enabled = True
'                sclSetGasValue(i).Enabled = True
'            End If
'        Next i
'    End If
'
'
'    txtSetGasValue(Index).Text = CStr(CSng(sclSetGasValue(Index).value / gbsngGasFlowScale(Index)))
'
'    For i = 0 To gbintMaxGasEnable
'        sngGas(i) = CSng(Val(txtSetGasValue(i).Text))
'    Next i
'
'    If (UCase(gbstrGasAlias(Index)) <> "PUMP") Then
'        ShowMFCStatus Index, CSng(txtSetGasValue(Index).Text)
'    End If
'
'
'    SetAO_MFC sngGas
    
End Sub

Public Sub ShowLampStatus(intIndex As Integer, sngValue As Single)
    Dim i As Integer
    
    
    
    
    If Val(sngValue) <> 0 Then
        If gbintNumOfBanks = 10 Then
            If intIndex < 6 Then
                shpLampL(intIndex).FillColor = RGB(255, 255, CInt(sngValue) * 2) 'vbYellow
                shpLampR(intIndex).FillColor = RGB(255, 255, CInt(sngValue) * 2) 'vbYellow
            Else
                shpLampT(intIndex - 6).FillColor = RGB(255, 255, CInt(sngValue) * 2) 'vbYellow
                shpLampB(intIndex - 6).FillColor = RGB(255, 255, CInt(sngValue) * 2) 'vbYellow
            End If
            
            
        Else
            Select Case intIndex
                Case 0  'SCR 1
                    shpLampL(intIndex).FillColor = RGB(255, 255, CInt(sngValue) * 2) 'vbYellow
                    shpLampR(intIndex).FillColor = RGB(255, 255, CInt(sngValue) * 2) 'vbYellow
                    shpLampT(intIndex).FillColor = RGB(255, 255, CInt(sngValue) * 2) 'vbYellow
                    shpLampB(intIndex).FillColor = RGB(255, 255, CInt(sngValue) * 2) 'vbYellow
                Case 1
                    shpLampT(intIndex).FillColor = RGB(255, 255, CInt(sngValue) * 2) 'vbYellow
                    shpLampB(intIndex).FillColor = RGB(255, 255, CInt(sngValue) * 2) 'vbYellow
                Case 2
                    shpLampT(intIndex).FillColor = RGB(255, 255, CInt(sngValue) * 2) 'vbYellow
                    shpLampB(intIndex).FillColor = RGB(255, 255, CInt(sngValue) * 2) 'vbYellow
                Case 3
                    shpLampL(1).FillColor = RGB(255, 255, CInt(sngValue) * 2) 'vbYellow
                    shpLampR(1).FillColor = RGB(255, 255, CInt(sngValue) * 2) 'vbYellow
                Case 4
                    shpLampL(2).FillColor = RGB(255, 255, CInt(sngValue) * 2) 'vbYellow
                    shpLampR(2).FillColor = RGB(255, 255, CInt(sngValue) * 2) 'vbYellow
            End Select
        
        End If
    Else
        For i = 0 To 5
            shpLampL(i).FillColor = &HC0C0&
            shpLampR(i).FillColor = &HC0C0&
            shpLampT(i).FillColor = &HC0C0&
            shpLampB(i).FillColor = &HC0C0&
        Next i
        
        
    End If
        
    
    

End Sub

Public Sub ShowMFCStatus(intIndex As Integer, sngValue As Single)
    Dim i As Integer
    Dim iCnt As Integer
    
    iCnt = 0
    If sngValue > 0 Then
        shpGas(intIndex).BackColor = GB_ColorLightCyan
        shpGasPipe(intIndex).BackColor = GB_ColorVeryLightCyan
    ElseIf sngValue <= 0 Then
        shpGas(intIndex).BackColor = GB_ColorCyan
        shpGasPipe(intIndex).BackColor = vbWhite
    End If
    For i = 0 To 1
    
        'If Val(sclSetGasValue(i).Value) > 0 Then iCnt = iCnt + 1
        If sngValue > 0 Then iCnt = iCnt + 1
    Next i
    If iCnt > 0 Then
        shpMainPipe(0).BackColor = GB_ColorVeryLightCyan
        shpMainPipe(1).BackColor = shpMainPipe(0).BackColor
        
    Else
        shpMainPipe(0).BackColor = vbWhite
        shpMainPipe(1).BackColor = shpMainPipe(0).BackColor
        
    End If
End Sub

Public Sub ShowStatus()
    Dim i As Integer
    
    shpReady.BackColor = IIf(SysDI.IsReady = 0, &H80&, &HFF00&)
    shpOverheat.BackColor = IIf(SysDI.IsOverHeat = 0, &H80&, &HFF00&)
    lbVacuum.BackColor = IIf(SysDI.IsChamberGaugeL = 0, &HFF00&, &H8000000F)
    lbVacuum.Caption = IIf(Para.useTPump = 1, Format(Kernel.sngPressure, "0.000000"), Format(Kernel.sngPressure, "0.000"))
    
    lbOxygen.Caption = Format(Kernel.sngOxygen, "0.00")
    gtcAirTank.SHOWMODE = IIf(SysDI.IsCDA = 0, 2, 1)
    gtcTankWater.SHOWMODE = IIf(SysDI.IsWater = 0, 2, 1)
    shpDoor.BackColor = IIf(SysDI.IsDoorClose = 0, &HFFFFFF, &H808080)
    If Kernel.IsAlarm = 0 Then
        shpAlarm.BackColor = &H80&
    End If
    For i = 1 To 4
        shpCT(i).BackColor = IIf(SysDI.IsLampError(i) = 0, &H80&, &HFF00&)
    Next i
    For i = 0 To 7
'        lbTC(i).Caption = Format(Kernel.sngTC(i), "0.0")
          lbTC(i).Caption = frmPlotProcess.SetFormat(Kernel.sngTC(i), gbintPrecisionDigit(i))
    Next i
    
    For i = 0 To 5
        shpGas(i).BackColor = IIf(Kernel.sngCurrOutMFC(i) > 0, &HFFFF00, &HC0C000)
'        If gblngAI_MFC_Read(i) >= 0 Then lbGasValue(i).Caption = Format(SysAI.sngMFC(i), "0.0")
        If lbMFC(i).Caption <> "APC" Then
        If gblngAI_MFC_Read(i) >= 0 Then lbGasValue(i).Caption = Format(SysAI.sngMFC(i), "0.00")
        Else
        Dim j As Integer
        For j = 5 To 7
        If gbstrNameTC(j) = "PS" Then
         lbGasValue(i).Caption = Format(Kernel.sngTC(j), "0.0")
        Exit For
        End If
        Next j
        End If
        
    Next i
    
    If Para.UseCT = 1 Then
        If CTDisplay = 0 Then
            For i = 0 To 59
                lbCT(i).Caption = Format(Kernel.dblCT(i), "0.0")
            Next i
        Else
            ShowCTTable
        End If
    End If
       
    If Para.UseCover = 1 Then
        If SysDI.IsCoverAlarm = 0 And SysDI.IsCoverServoRdy = 1 And SysDI.IsCoverOrigRdy = 1 Then
            lbCoverStatus.Caption = IIf(SysDI.IsCoverMoving = 0, "正常", "運轉中")
            lbCoverStatus.BackColor = IIf(SysDI.IsCoverMoving = 0, &H8000000F, &HFF00&)
        Else
            lbCoverStatus.Caption = "錯誤"
            lbCoverStatus.BackColor = &HFF&
        End If
        'cmdCoverUp.Enabled = IIf(gblngDI_CoverUpInpos = 0, True, False)
        lbCoverUp.BackColor = IIf(SysDI.IsCoverUp = 0, &H8000000F, &HFF00&)
        'cmdCoverDown.Enabled = IIf(gblngDI_CoverDownInpos = 0, True, False)
        lbCoverDown.BackColor = IIf(SysDI.IsCoverDown = 0, &H8000000F, &HFF00&)
    End If
        
End Sub

Public Sub RefreshGasDefiniation()
    Dim j As Integer
    Dim k As Integer
    Dim strTemp As String
    
        For j = 0 To 5
        lbGasName(j).Caption = gbstrGasAlias(j)
        lbMFC(j).Caption = Trim(gbstrGasAlias(j))
        lbGasUnit(j).Caption = gbstrGasUnit(j)
        frmPlotProcess.lbNameMFC(j).Caption = Trim(gbstrGasAlias(j))
        frmPlotProcess.lbGasUnit(j).Caption = Trim(gbstrGasUnit(j))
        gbsngGasFlowScale(j) = CheckValueScale(CStr(gbsngMaxGasSLMP(j)))
        sclSetGasValue(j).Max = gbsngMaxGasSLMP(j) * gbsngGasFlowScale(j)
        strTemp = "0."
        For k = 0 To (gbsngGasFlowScale(j) ^ 0.1)
            strTemp = strTemp & "0"
        Next k
        gbstrGasPrecision(j) = strTemp
        lbGasValue(j).Caption = Format(0, gbstrGasPrecision(j))
    Next j
'    For j = 0 To gbintMaxGasEnable
'        If gbstrGasAlias(j) <> "NA" And gbstrGasAlias(j) <> "Pump" Then
'            lbMFC(j).Visible = True
'            shpGas(j).Visible = True
'            lbGasValue(j).Visible = True
'            shpGasPipe(j).Visible = True
'            txtSetGasValue(j).Enabled = True
'            sclSetGasValue(j).Enabled = True
'        Else
'            lbMFC(j).Visible = False
'            shpGas(j).Visible = False
'            shpGas(j).Visible = False
'            lbGasValue(j).Visible = False
'            shpGasPipe(j).Visible = False
'            txtSetGasValue(j).Enabled = False
'            sclSetGasValue(j).Enabled = False
'        End If
'        If gbstrGasAlias(j) = "Pump" Then
'            txtSetGasValue(j).Enabled = True
'            sclSetGasValue(j).Enabled = True
'            lbGasValue(j).Visible = False
'        End If
'
'
'
'        lbGasName(j).Caption = gbstrGasAlias(j)
'        lbMFC(j).Caption = Trim(gbstrGasAlias(j))
'        lbGasUnit(j).Caption = gbstrGasUnit(j)
'        frmPlotProcess.lbNameMFC(j).Caption = Trim(gbstrGasAlias(j))
'        frmPlotProcess.lbGasUnit(j).Caption = Trim(gbstrGasUnit(j))
'        gbsngGasFlowScale(j) = CheckValueScale(CStr(gbsngMaxGasSLMP(j)))
'        sclSetGasValue(j).Max = gbsngMaxGasSLMP(j) * gbsngGasFlowScale(j)
'        strTemp = "0."
'        For K = 0 To (gbsngGasFlowScale(j) ^ 0.1)
'            strTemp = strTemp & "0"
'        Next K
'        gbstrGasPrecision(j) = strTemp
'        lbGasValue(j).Caption = Format(0, gbstrGasPrecision(j))
'    Next j
    

End Sub





Private Sub TxtChamberNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If TxtChamberNo.text <> "" Then
Select Case MsgBox("是否要更新此腔體編號", vbYesNo + vbQuestion, "確認")
        Case vbYes
           Call WritePrivateProfileString("DeviceInfo", "ChamberNo", TxtChamberNo.text, App.Path + DeviceInfo_Path)
            mdifrmRTP.tbrRTP.Buttons(14).Caption = "腔體編號:" & TxtChamberNo.text
           chkBlkTest.SetFocus
    End Select
    Else
MsgBox "請先輸入腔體編號!!!"
End If
End If

End Sub

Private Sub txtSetGasValue_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    Dim sngGas(10) As Single
    Dim i As Integer
    
    On Error GoTo ERRLINE
    
    If KeyCode = 13 Then
'            sclSetGasValue(Index).Enabled = False
'            If CSng(Val(txtSetGasValue(Index).Text)) > CSng(sclSetGasValue(Index).Max / gbsngGasFlowScale(Index)) Then
'                txtSetGasValue(Index).Text = CStr(CSng(sclSetGasValue(Index).Max / gbsngGasFlowScale(Index)))
'            ElseIf CSng(Val(txtSetGasValue(Index).Text)) < 0 Then
'                txtSetGasValue(Index).Text = "0"
'            End If
'            sclSetGasValue(Index).value = CLng(Val(txtSetGasValue(Index).Text) * gbsngGasFlowScale(Index))
'            sclSetGasValue(Index).Enabled = True
'            sclSetGasValue_Change (Index)
'            If sclSetGasValue(Index).value > 0 Then
'                Call frmHistory.AppendLogAlert(1, "Manual", 1010, "手動開啟MFC", 1)
'            Else
'                Call frmHistory.AppendLogAlert(1, "Manual", 1010, "手動關閉MFC", 1)
'            End If
        For i = 0 To 5
            sngGas(i) = Val(txtSetGasValue(i).text)
        Next i
        SetAO_MFC sngGas
    End If

    Exit Sub
ERRLINE:
    ShowMessageOK "輸入數據錯誤"
    
End Sub
Private Sub tmrCoverOrig_Timer()
    If SysDI.IsCoverServoRdy = 1 And SysDI.IsCoverOrigRdy = 1 Then
        tmrCoverOrig.Enabled = False
        SetCover False
    Else
        If gbintCoverOrigCount > 10 Then
            tmrCoverOrig.Enabled = False
            gbstrAlarmHint = " Cover Servo error"
            ShowAlarmFlash 28
        Else
            gbintCoverOrigCount = gbintCoverOrigCount + 1
        End If
    End If
End Sub

Private Sub cmdCoverOrig_Click()
    Call frmHistory.AppendLogAlert(1, "Manual", 1050, "Cover Home", 1)
    InitCover
End Sub
Private Sub cmdCoverDown_Click()
    Call frmHistory.AppendLogAlert(1, "Manual", 1051, "Cover Down", 1)
    SetCover True
End Sub

Private Sub cmdCoverUp_Click()
    Call frmHistory.AppendLogAlert(1, "Manual", 1052, "Cover Up", 1)
    SetCover False
End Sub

Public Sub ShowCTTable()
        Dim i As Integer, j As Integer, k As Integer
    Dim sngTotalGridWidth As Single
    Dim S As String
    Dim s1 As String
    Dim startIdx As Integer
    Dim values(3) As Integer
    Dim m As Integer
    Dim maxVal As Integer
    Dim names(3) As String
    Dim rowCount As Integer
    Dim nameOrder(3) As Integer
    Dim rowIndex As Integer
    
    On Error Resume Next
        values(0) = CTNumber1
        values(1) = CTNumber2
        values(2) = CTNumber3
        values(3) = CTNumber4
        names(0) = CTName1
        names(1) = CTName2
        names(2) = CTName3
        names(3) = CTName4
        nameOrder(0) = CInt(CTOrder1)
        nameOrder(1) = CInt(CTOrder2)
        nameOrder(2) = CInt(CTOrder3)
        nameOrder(3) = CInt(CTOrder4)
    On Error GoTo 0
    
    maxVal = values(0)
    For m = 1 To 3
        If values(m) > maxVal Then
            maxVal = values(m)
        End If
    Next m
    rowCount = 0
    For i = LBound(names) To UBound(names)
        If Trim(names(i)) <> "" Then
            rowCount = rowCount + 1
        End If
    Next i
    With hfgCTConfigProcess
        .Left = 200
        .FixedCols = 1
        .FixedRows = 1
        .Rows = rowCount + 1
        .Cols = maxVal + 1
        For i = 0 To .Cols - 1
            .ColWidth(i) = 800
        Next i
        For i = 0 To .Cols - 1
            sngTotalGridWidth = sngTotalGridWidth + .ColWidth(i)
            .ColAlignmentFixed = flexAlignCenterCenter
            .ColAlignment(i) = flexAlignCenterCenter
        Next i
        .Width = sngTotalGridWidth + 150
        
        For k = 1 To .Cols - 1
            .TextMatrix(0, k) = CStr(k)
        Next k
        
        For i = LBound(names) To UBound(names)
            If Trim(names(i)) <> "" Then
                .TextMatrix(nameOrder(i), 0) = names(i)
            End If
        Next i
        
        startIdx = 0
'        rowIndex = 1
        For i = 0 To UBound(values)
            Dim rowLength As Integer
            rowLength = values(i)
            If rowLength > 0 Then
                For k = 1 To rowLength
                    If startIdx + k - 1 <= UBound(Kernel.dblCT) Then
                        .TextMatrix(nameOrder(i), k) = Format(Kernel.dblCT(startIdx + k - 1), "#0.0")
                    End If
                Next k
                startIdx = startIdx + rowLength
'                rowIndex = rowIndex + 1
            End If
        Next i
        
        
        .Refresh
        .AllowUserResizing = flexResizeNone
    End With
End Sub
