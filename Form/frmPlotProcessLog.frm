VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPlotProcessLog 
   Caption         =   "製程曲線"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18885
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   15
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   18885
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtRecipeInfo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   6960
      TabIndex        =   134
      Top             =   720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox TxtRecipeInfo 
      Height          =   495
      Index           =   1
      Left            =   3600
      TabIndex        =   132
      Top             =   720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox TxtRecipeInfo 
      Height          =   495
      Index           =   0
      Left            =   960
      TabIndex        =   130
      Top             =   720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame FraPower 
      Caption         =   "能量"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   12000
      TabIndex        =   102
      Top             =   8400
      Width           =   6855
      Begin VB.TextBox IntValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   4
         Left            =   3360
         TabIndex        =   126
         Top             =   840
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H00008080&
         Caption         =   "Int5"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   34
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   125
         Top             =   840
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.TextBox IntValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   3
         Left            =   1080
         TabIndex        =   124
         Top             =   840
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H0000C0C0&
         Caption         =   "Int4"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   33
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   840
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.TextBox IntValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   2
         Left            =   5760
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H0000FFFF&
         Caption         =   "Int3"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   32
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   360
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.TextBox IntValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   3360
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H0080FFFF&
         Caption         =   "Int2"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   31
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   121
         Top             =   360
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.TextBox IntValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   1080
         TabIndex        =   104
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Int1"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   360
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin VB.Frame FraGas 
      Caption         =   "氣體"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   12000
      TabIndex        =   89
      Top             =   6480
      Width           =   6855
      Begin VB.TextBox GasValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   6
         Left            =   5760
         TabIndex        =   128
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox ChkGasColor 
         BackColor       =   &H00800080&
         Caption         =   "NA"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   127
         Top             =   1320
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.TextBox GasValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   5
         Left            =   3360
         TabIndex        =   101
         Top             =   1320
         Width           =   855
      End
      Begin VB.CheckBox ChkGasColor 
         BackColor       =   &H000080FF&
         Caption         =   "NA"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   360
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.TextBox GasValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   4
         Left            =   3360
         TabIndex        =   99
         Top             =   840
         Width           =   855
      End
      Begin VB.CheckBox ChkGasColor 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NA"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   1320
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.TextBox GasValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   3
         Left            =   3360
         TabIndex        =   97
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox ChkGasColor 
         BackColor       =   &H00FF8080&
         Caption         =   "NA"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   840
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.TextBox GasValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   1080
         TabIndex        =   95
         Top             =   840
         Width           =   855
      End
      Begin VB.CheckBox ChkGasColor 
         BackColor       =   &H00FFFF00&
         Caption         =   "NA"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   360
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.TextBox GasValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   2
         Left            =   1080
         TabIndex        =   93
         Top             =   1320
         Width           =   855
      End
      Begin VB.CheckBox ChkGasColor 
         BackColor       =   &H0000C000&
         Caption         =   "NA"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         HelpContextID   =   1
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   840
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.TextBox GasValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   1080
         TabIndex        =   91
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox ChkGasColor 
         BackColor       =   &H00FF00FF&
         Caption         =   "NA"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         HelpContextID   =   1
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   360
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin VB.Frame fraMTCB 
      Caption         =   "MTC2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   16560
      TabIndex        =   70
      Top             =   2280
      Width           =   2175
      Begin VB.TextBox TempValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   23
         Left            =   1080
         TabIndex        =   120
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox TempValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   22
         Left            =   1080
         TabIndex        =   119
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox TempValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   21
         Left            =   1080
         TabIndex        =   118
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox TempValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   20
         Left            =   1080
         TabIndex        =   117
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox TempValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   19
         Left            =   1080
         TabIndex        =   116
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox TempValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   18
         Left            =   1080
         TabIndex        =   115
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox TempValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   17
         Left            =   1080
         TabIndex        =   114
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox TempValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   16
         Left            =   1080
         TabIndex        =   113
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H008080FF&
         Caption         =   "TC24"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         HelpContextID   =   1
         Index           =   30
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   3720
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H00FF8080&
         Caption         =   "TC23"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         HelpContextID   =   1
         Index           =   29
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   3240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H00FF00FF&
         Caption         =   "TC22"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         HelpContextID   =   1
         Index           =   28
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   2760
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H000080FF&
         Caption         =   "TC21"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         HelpContextID   =   1
         Index           =   27
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   2280
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H00FFFF00&
         Caption         =   "TC19"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         HelpContextID   =   1
         Index           =   25
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   1320
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H0000FF00&
         Caption         =   "TC18"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         HelpContextID   =   1
         Index           =   24
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   840
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H000000FF&
         Caption         =   "TC17"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   23
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   360
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H0000FFFF&
         Caption         =   "TC20"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   26
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   1800
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdSaveCSV 
      Caption         =   "另存CSV"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15000
      MaskColor       =   &H8000000F&
      TabIndex        =   69
      Top             =   600
      Width           =   1575
   End
   Begin VB.Frame fraMTC 
      Caption         =   "MTC1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   14280
      TabIndex        =   60
      Top             =   2280
      Width           =   2175
      Begin VB.TextBox TempValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   15
         Left            =   1080
         TabIndex        =   112
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox TempValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   14
         Left            =   1080
         TabIndex        =   111
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox TempValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   13
         Left            =   1080
         TabIndex        =   110
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox TempValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   12
         Left            =   1080
         TabIndex        =   109
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox TempValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   11
         Left            =   1080
         TabIndex        =   108
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox TempValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   10
         Left            =   1080
         TabIndex        =   107
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox TempValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   9
         Left            =   1080
         TabIndex        =   106
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox TempValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   8
         Left            =   1080
         TabIndex        =   105
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H0000FFFF&
         Caption         =   "TC12"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   1800
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H000000FF&
         Caption         =   "TC9"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   360
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H0000FF00&
         Caption         =   "TC10"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         HelpContextID   =   1
         Index           =   16
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   840
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H00FFFF00&
         Caption         =   "TC11"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         HelpContextID   =   1
         Index           =   17
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   1320
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H000080FF&
         Caption         =   "TC13"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         HelpContextID   =   1
         Index           =   19
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   2280
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H00FF00FF&
         Caption         =   "TC14"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         HelpContextID   =   1
         Index           =   20
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   2760
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H00FF8080&
         Caption         =   "TC15"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         HelpContextID   =   1
         Index           =   21
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   3240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H008080FF&
         Caption         =   "TC16"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         HelpContextID   =   1
         Index           =   22
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   3720
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   1215
      Left            =   10560
      TabIndex        =   35
      Top             =   9600
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   2143
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "PID 參數"
      TabPicture(0)   =   "frmPlotProcessLog.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbPIDValue(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbName(5)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbPIDValue(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbName(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbName(10)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbName(11)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbPIDValue(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbPIDValue(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Barcode ID"
      TabPicture(1)   =   "frmPlotProcessLog.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbName(6)"
      Tab(1).Control(1)=   "lbName(1)"
      Tab(1).Control(2)=   "lbName(3)"
      Tab(1).Control(3)=   "lbName(4)"
      Tab(1).Control(4)=   "lbPN"
      Tab(1).Control(5)=   "lbBN"
      Tab(1).Control(6)=   "lbID1"
      Tab(1).Control(7)=   "lbID2"
      Tab(1).ControlCount=   8
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "PN:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   6
         Left            =   -74880
         TabIndex        =   58
         Top             =   480
         Width           =   390
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "BN:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   -74880
         TabIndex        =   57
         Top             =   840
         Width           =   390
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "EN:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   -74880
         TabIndex        =   56
         Top             =   1200
         Width           =   390
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "ID:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   4
         Left            =   -74880
         TabIndex        =   55
         Top             =   1560
         Width           =   285
      End
      Begin VB.Label lbPN 
         Caption         =   "NA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74400
         TabIndex        =   54
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lbBN 
         Caption         =   "NA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74400
         TabIndex        =   53
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label lbID1 
         Caption         =   "NA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74400
         TabIndex        =   52
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label lbID2 
         Caption         =   "NA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74400
         TabIndex        =   51
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label lbPIDValue 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   50
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lbPIDValue 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   49
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "P1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   11
         Left            =   240
         TabIndex        =   48
         Top             =   480
         Width           =   300
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "Inte1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   10
         Left            =   240
         TabIndex        =   47
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "Inte2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   46
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lbPIDValue 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   45
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "P2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   5
         Left            =   240
         TabIndex        =   44
         Top             =   840
         Width           =   300
      End
      Begin VB.Label lbPIDValue 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   43
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "Proportional2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   20
         Left            =   -74880
         TabIndex        =   42
         Top             =   960
         Width           =   1410
      End
      Begin VB.Label lbPIDValue 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   12
         Left            =   -73320
         TabIndex        =   41
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lbPIDValue 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   10
         Left            =   -73320
         TabIndex        =   40
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lbPIDValue 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   9
         Left            =   -73320
         TabIndex        =   39
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lbPIDValue 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   8
         Left            =   -73320
         TabIndex        =   38
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "Proportional"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   18
         Left            =   -74880
         TabIndex        =   37
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "Integral"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   17
         Left            =   -74880
         TabIndex        =   36
         Top             =   1320
         Width           =   750
      End
   End
   Begin VB.Frame fraColor 
      Caption         =   "溫度"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   12000
      TabIndex        =   8
      Top             =   2280
      Width           =   2175
      Begin VB.TextBox TempValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   7
         Left            =   1080
         TabIndex        =   88
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox TempValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   6
         Left            =   1080
         TabIndex        =   87
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox TempValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   5
         Left            =   1080
         TabIndex        =   86
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox TempValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   4
         Left            =   1080
         TabIndex        =   85
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox TempValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   3
         Left            =   1080
         TabIndex        =   84
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox TempValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   2
         Left            =   1080
         TabIndex        =   83
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox TempValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   390
         Index           =   1
         Left            =   1080
         TabIndex        =   82
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox TempValue 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   390
         Index           =   0
         Left            =   1080
         TabIndex        =   81
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H008080FF&
         Caption         =   "TC8"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         HelpContextID   =   1
         Index           =   14
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   3720
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H00FF8080&
         Caption         =   "TC7"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         HelpContextID   =   1
         Index           =   13
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H00FF00FF&
         Caption         =   "TC6"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         HelpContextID   =   1
         Index           =   12
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   2760
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H000080FF&
         Caption         =   "TC5"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         HelpContextID   =   1
         Index           =   11
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2280
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H00FFFF00&
         Caption         =   "TC3"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         HelpContextID   =   1
         Index           =   9
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1320
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H0000FF00&
         Caption         =   "TC2"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         HelpContextID   =   1
         Index           =   8
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   840
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H000000FF&
         Caption         =   "TC1"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   360
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H0000FFFF&
         Caption         =   "TC4"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1800
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Vacuum (Sec)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   12000
      TabIndex        =   22
      Top             =   9840
      Width           =   2535
      Begin VB.Label lbPumpDownTime 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   375
         Left            =   1440
         TabIndex        =   26
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "Factor2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   16
         Left            =   120
         TabIndex        =   25
         Top             =   2040
         Width           =   810
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "真空延遲:"
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
         Index           =   15
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1350
      End
      Begin VB.Label lbPIDValue 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   1560
         TabIndex        =   23
         Top             =   2040
         Width           =   1095
      End
   End
   Begin VB.TextBox txtArray 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   10
      Left            =   5880
      TabIndex        =   19
      Text            =   "TextArray"
      Top             =   9600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtArray 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   9
      Left            =   3720
      TabIndex        =   18
      Text            =   "TextArray"
      Top             =   9600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtArray 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   8
      Left            =   2400
      TabIndex        =   17
      Text            =   "TextArray"
      Top             =   9600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtArray 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   7
      Left            =   1080
      TabIndex        =   16
      Text            =   "TextArray"
      Top             =   9600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtArray 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   6
      Left            =   4920
      TabIndex        =   15
      Text            =   "TextArray"
      Top             =   9600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtArray 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   5
      Left            =   6960
      TabIndex        =   14
      Text            =   "TextArray"
      Top             =   9600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtArray 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   4
      Left            =   8040
      TabIndex        =   13
      Text            =   "TextArray"
      Top             =   9600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtArray 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   3
      Left            =   9240
      TabIndex        =   12
      Text            =   "TextArray"
      Top             =   9600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtArray 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   2
      Left            =   6960
      TabIndex        =   11
      Text            =   "TextArray"
      Top             =   9840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtArray 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   6480
      TabIndex        =   10
      Text            =   "TextArray"
      Top             =   9720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtArray 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   6000
      TabIndex        =   9
      Text            =   "TextArray"
      Top             =   9600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame fraProcessLogChart 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   11655
      Begin VB.PictureBox picProcessLog 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8295
         Index           =   0
         Left            =   0
         ScaleHeight     =   8235
         ScaleWidth      =   11835
         TabIndex        =   3
         Top             =   120
         Width           =   11895
         Begin VB.Shape shpRect 
            BorderColor     =   &H80000003&
            BorderStyle     =   3  'Dot
            FillStyle       =   4  'Upward Diagonal
            Height          =   1335
            Left            =   5520
            Top             =   3120
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Label lbCurrTime 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5640
            TabIndex        =   21
            Top             =   0
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Line linMove 
            BorderColor     =   &H000000FF&
            BorderStyle     =   3  'Dot
            Visible         =   0   'False
            X1              =   3960
            X2              =   3960
            Y1              =   0
            Y2              =   8160
         End
         Begin VB.Label lbSec 
            AutoSize        =   -1  'True
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1920
            TabIndex        =   5
            Top             =   120
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lbTemp 
            AutoSize        =   -1  'True
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   4
            Top             =   2520
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Line linCourseVer 
            BorderColor     =   &H000000FF&
            Visible         =   0   'False
            X1              =   1920
            X2              =   1920
            Y1              =   360
            Y2              =   2880
         End
         Begin VB.Line linCourseHor 
            BorderColor     =   &H000000FF&
            Visible         =   0   'False
            X1              =   0
            X2              =   1920
            Y1              =   2880
            Y2              =   2880
         End
      End
   End
   Begin VB.Frame fraProcessHistory 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   12000
      TabIndex        =   6
      Top             =   360
      Width           =   2895
      Begin VB.CommandButton cmdZoomBack 
         Caption         =   "重置"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1560
         MaskColor       =   &H8000000F&
         TabIndex        =   28
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox chkZoom 
         Caption         =   "放大"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1080
         Width           =   1095
      End
      Begin MSComDlg.CommonDialog cdFile 
         Left            =   240
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdOpenLog 
         BackColor       =   &H80000002&
         Caption         =   "開啟檔案"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         MaskColor       =   &H8000000F&
         TabIndex        =   7
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Label lblRecipeName 
      Caption         =   "RecipeName:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   5400
      TabIndex        =   133
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lbl_LotID 
      Caption         =   "LotID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2880
      TabIndex        =   131
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lbl_OPID 
      Caption         =   "OPID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   240
      TabIndex        =   129
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lbData 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   80
      Top             =   120
      Width           =   8295
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbValue 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   79
      Top             =   120
      Width           =   60
   End
End
Attribute VB_Name = "frmPlotProcessLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lngPlot_X_Point(100) As Long
Dim lngPlot_Y_Point(100) As Long

Dim lngAssignX As Long
Dim lngAssignY As Long
Dim intColumnPitch As Long
Dim intRowPitch As Long

Dim sngXPitch As Single
Dim sngYPitch As Single
Dim lngTable_Xsize As Long
Dim lngTable_Ysize As Long
Dim SaveFileName As String

Dim lngMaxTime As Long
Dim lngMaxTemp As Long

Dim m_PlotXY(100) As PlotRecipePoint   'get recipe from table
Dim m_SecPoint(100) As Single    'Because Process Step limit 20
Dim m_TempPoint(100)  As Single  'declear Tempurature point max 20 point
Dim tmpX As Single
Dim tmpY As Single

Public CurrRatio As Single

Dim blnIntensityVisible As Boolean
Dim blnTCVisible As Boolean
Dim blnGas1Visible As Boolean
Dim blnIGas2Visible As Boolean
Dim blnGasVisible(GB_GAS_MAX) As Boolean
Dim blnTCArrayVisible(10) As Boolean
Dim blnPowerArrayVisible(10) As Boolean
Dim blnOxygenVisible As Boolean
Dim blnVacuumVisible As Boolean

Dim lngGasProfileColor(GB_GAS_MAX) As Long

Dim blnOpened   As Boolean
Dim lnGridW As Long
Dim lnGridH As Long
Dim TimeMap(65535) As Long
Dim iniLog As New cInifile

Dim blnZoom   As Boolean
Dim ZoomRect As RectType




Private Sub cmdSaveCSV_Click()
    Dim StrFileName As String
    Dim strNewFileName As String
    Dim TextLine As String
    Dim NewLine As String
    
    gbblnNoModalForm = True
    blnOpened = False
    blnZoom = False
    On Error GoTo ERRHNADLE
    Call frmConfiguration.StopWatchDog
    cdFile.InitDir = gbSystemPath & "\Log"
    cdFile.Filter = "*.log|*.log"
    cdFile.FilterIndex = 1
    cdFile.CancelError = True
    cdFile.ShowOpen
    gbblnNoModalForm = False
    
'    If cdFile.FileName <> "" Then
'        strFileName = cdFile.FileName
'        strNewFileName = Mid(strFileName, 1, Len(strFileName) - 3) & "csv"
'        'FileCopy strFileName, strNewFileName
'
'       Open strFileName For Input As #1
'       Open strNewFileName For Output As #2
'
'        Do While Not EOF(1)   ' 執行迴圈直到檔尾為止。
'            Line Input #1, TextLine   ' 讀入一行資料並將之指定給變數。
'            NewLine = Replace(TextLine, "=", ",")
'            Print #2, NewLine
'        Loop
'
'        Close #1
'        Close #2
'
'    End If
    
    If cdFile.fileName <> "" Then
        StrFileName = cdFile.fileName
    cdFile.Filter = "CSV files (*.csv)|*.csv"
    cdFile.FilterIndex = 1
    cdFile.CancelError = True
    cdFile.ShowSave
    
    If cdFile.fileName <> "" Then
        strNewFileName = cdFile.fileName
        
        If Right$(strNewFileName, 4) <> ".csv" Then
            strNewFileName = Left$(strNewFileName, InStrRev(strNewFileName, ".") - 1)
            strNewFileName = strNewFileName & ".csv"
        End If
       
        Open strNewFileName For Output As #1
        If Len(StrFileName) > 0 Then
            Open StrFileName For Input As #2
            Do While Not EOF(2)
                Line Input #2, TextLine
                NewLine = Replace(TextLine, "=", ",")
                Print #1, NewLine
            Loop
            Close #2
        End If
        Close #1
        MsgBox "CSV file saved successfully!", vbInformation
    End If
End If
    frmConfiguration.StartWatchDog
    blnOpened = True
    Exit Sub
ERRHNADLE:
    Close #1
    Close #2
    frmConfiguration.StartWatchDog
    blnOpened = False
End Sub

Private Sub Form_Activate()
    Dim i As Integer
    
    For i = 0 To UBound(gbstrGasAlias)
'        If UCase(gbstrGasAlias(i)) <> "PUMP" And UCase(gbstrGasAlias(i)) <> "NA" Then ChkGasColor(i).Caption = Trim(gbstrGasAlias(i))
         If UCase(gbstrGasAlias(i)) <> "PUMP" Then
         ChkGasColor(i).Caption = Trim(gbstrGasAlias(i))
         Else
         ChkGasColor(i).Caption = "Vac"
         End If
    Next i

    
     For i = 6 To 29
        chkPlotColor(i + 1).Caption = Trim(gbstrNameTC(i - 6))
    Next i
       
    fraMTC.Visible = IIf(Para.UseMTC = 1, True, False)
'    lbMTC.Visible = IIf(Para.UseMTC = 1, True, False)
    fraMTCB.Visible = IIf(Para.UseMTCB = 1, True, False)
'    lbMTCB.Visible = IIf(Para.UseMTCB = 1, True, False)
End Sub


Private Sub Form_Load()
    Dim bRet    As Boolean
    Dim PicNO           As Long
    Dim i    As Integer

    blnIntensityVisible = True
    blnTCVisible = True
    blnGas1Visible = True
    blnIGas2Visible = True
    For i = 0 To gbintMaxGasEnable
        blnGasVisible(i) = True
    Next i
    blnOxygenVisible = True
    blnVacuumVisible = True
    blnOpened = False
    blnZoom = False

    
    
    picProcessLog(0).AutoRedraw = True
    
    
    ZoomRect.Left = -7500
    ZoomRect.Top = 11000
    ZoomRect.Right = 67500
    ZoomRect.Bottom = -500
    picProcessLog(0).Width = 11000
    picProcessLog(0).Height = 9000
    fraProcessLogChart.Height = picProcessLog(0).Height + 500
    fraProcessLogChart.Width = picProcessLog(0).Width + 500
    
    'picProcessLog(0).Left = fraProcessLogChart.Left + 50
    'picProcessLog(0).Top = fraProcessLogChart.Top + 100
    Call Plot_ProcessLogTable(60, 1000)
'    For i = 0 To 3
'        chkPlotColor(i + 3).Caption = Trim(gbstrGasAlias(i))
'    Next i
For i = 0 To 23
'TempValue(i).Enabled = False
TempValue(i).Locked = True
Next i
For i = 0 To 2
TxtRecipeInfo(i).Locked = True
Next i
For i = 0 To 6
'GasValue(i).Enabled = False
GasValue(i).Locked = True
Next i
'VacValue.Locked = True

For i = 0 To 4
'IntValue(i).Enabled = False
IntValue(i).Locked = True
Next i
Frame1.Visible = False
lbData.Visible = False
chkPlotColor(0).value = 1
chkPlotColor(31).value = 0
chkPlotColor(32).value = 0
chkPlotColor(33).value = 0
chkPlotColor(34).value = 0
End Sub


Private Sub chkPlotColor_Click(Index As Integer)
    Select Case Index
        Case 0
            blnIntensityVisible = IIf(chkPlotColor(Index).value = 1, True, False)
        Case 1
            blnTCVisible = IIf(chkPlotColor(Index).value = 1, True, False)
        Case 2
            'blnIntensityVisible = IIf(chkPlotColor(Index).Value = 0, True, False)
        Case 3
            blnGasVisible(0) = IIf(chkPlotColor(Index).value = 1, True, False)
            blnGas1Visible = IIf(chkPlotColor(Index).value = 1, True, False)
        Case 4
            blnGasVisible(1) = IIf(chkPlotColor(Index).value = 1, True, False)
            blnIGas2Visible = IIf(chkPlotColor(Index).value = 1, True, False)
        Case 5
            blnGasVisible(2) = IIf(chkPlotColor(Index).value = 1, True, False)
        Case 6
            blnGasVisible(3) = IIf(chkPlotColor(Index).value = 1, True, False)
        Case 7
            blnOxygenVisible = IIf(chkPlotColor(Index).value = 1, True, False)
        Case 8
            blnVacuumVisible = IIf(chkPlotColor(Index).value = 1, True, False)

    End Select
End Sub







Private Sub cmdOpenLog_Click()
    Dim StrFileName As String
    gbblnNoModalForm = True
    blnOpened = False
    blnZoom = False
    On Error GoTo ERRHNADLE
    Call frmConfiguration.StopWatchDog
    cdFile.InitDir = gbSystemPath & "\Log"
    cdFile.Filter = "*.log|*.log"
    cdFile.FilterIndex = 1
    cdFile.CancelError = True
    cdFile.ShowOpen
    gbblnNoModalForm = False
    If cdFile.fileName <> "" Then
        StrFileName = cdFile.fileName
        frmProgress.Show
        Call OpenProcessLog(StrFileName)
        Me.Caption = StrFileName
        iniLog.Path = cdFile.fileName
    End If
    If Me.PlotProcessLogChart = False Then
        ShowMessageOK "記錄檔開啟失敗!"
        Else
        Unload frmProgress
       TxtRecipeInfo(2).text = Mid(StrFileName, InStrRev(StrFileName, "\") + 1)
    End If

    frmConfiguration.StartWatchDog
    blnOpened = True
        
    cmdOpenLog.Enabled = True
    chkZoom.Enabled = True
    cmdZoomBack.Enabled = True
    CurrRatio = 1
    fraColor.Visible = False
    fraColor.Visible = True
    
    picProcessLog(0).Visible = False
    picProcessLog(0).Visible = True
    
    
    If Para.UseMTC = 1 Then
        fraMTC.Visible = False
        fraMTC.Visible = True
    End If
    If Para.UseMTCB = 1 Then
        fraMTCB.Visible = False
        fraMTCB.Visible = True
    End If
    Exit Sub
ERRHNADLE:
    frmConfiguration.StartWatchDog
    blnOpened = False
    
    'Call Plot_ProcessLogTable(60, 1000)
End Sub

Private Sub chkZoom_Click()
    blnZoom = chkZoom.value
End Sub

Private Sub cmdZoomBack_Click()
    '    Dim I, OldFontSize   ' 宣告變數。
'   Width = 8640: Height = 5760   ' 以 twip 設定表單大小。
'   Move 100, 100  ' 移動表單至起點。
'   AutoRedraw = -1   ' 開啟 AutoRedraw。
'   OldFontSize = FontSize   ' 維持原來字型的大小。
'   BackColor = QBColor(7)   ' 將背景設定為灰色。
'   Scale (0, 110)-(130, 0)   ' 設定自訂座標系統。
'   For I = 100 To 10 Step -10
'      Line (0, I)-(2, I)   ' 每隔 10 個單位劃刻度。
'      CurrentY = CurrentY + 1.5   ' 移動指標位置。
'      Print I   ' Print scale mark value on left.
'      Line (ScaleWidth - 2, I)-(ScaleWidth, I)
'      CurrentY = CurrentY + 1.5   ' 移動指標位置。
'      CurrentX = ScaleWidth - 9
'      Print I   ' 將刻度值列印在右邊。
'   Next I
'   ' 畫長條圖。
'   Line (10, 0)-(20, 45), RGB(0, 0, 255), BF   ' 第一條為藍色。
'   Line (20, 0)-(30, 55), RGB(255, 0, 0), BF   ' 第一條為紅色。
'   Line (40, 0)-(50, 40), RGB(0, 0, 255), BF
'   Line (50, 0)-(60, 25), RGB(255, 0, 0), BF
'   Line (70, 0)-(80, 35), RGB(0, 0, 255), BF
'   Line (80, 0)-(90, 60), RGB(255, 0, 0), BF
'   Line (100, 0)-(110, 75), RGB(0, 0, 255), BF
'   Line (110, 0)-(120, 90), RGB(255, 0, 0), BF
'   CurrentX = 18: CurrentY = 100   ' 移動指標位置。
'   FontSize = 14   ' 放大標題尺寸。
'   Print Widget; Quarterly; Sales   ' 列印標題。
'   FontSize = OldFontSize   ' 回存字型大小。
'   CurrentX = 27: CurrentY = 93   ' 移動指標位置。
'   Print Planned; Vs.Actual     ' 列印子標題。
'   Line (29, 86)-(34, 88), RGB(0, 0, 255), BF   ' 列印圖例。
'   Line (43, 86)-(49, 88), RGB(255, 0, 0), BF
    
    
'    picProcessLog(0).Scale (-intColumnPitch * 2, gbProcessLogTable.Yscale_size + intRowPitch) _
'                        -(gbProcessLogTable.Xscale_size + intColumnPitch, -intRowPitch)
    frmProgress.Show
    blnZoom = False
    PlotProcessLogChart
    cmdOpenLog.Enabled = True
    chkZoom.Enabled = True
    cmdZoomBack.Enabled = True
    Unload frmProgress
    'SetCursor picProcessLog(0), blnZoom
End Sub

Public Function Zoom_ProcessLogTable(lngProcessTime As Long, lngMaxTemp As Long) As Boolean
    Dim lngX As Long
    Dim lngY As Long

    Dim iCount As Integer
    Dim PercentPitch  As Integer
    Dim PercentStep As Integer
    Dim TempPitch As Integer
    Dim TempStep As Integer
    Dim TimePitch As Integer
    Dim TimeStep As Integer
    Dim m_Min As Integer
    Dim m_Sec As Integer
    Dim iScaleX As Integer
    Dim iScaleY As Integer

    Dim i As Integer
    
    With gbProcessLogTable
        .Xsize = 10000
        .Ysize = 8000
        .ExtendXY_Size = 500
        .row = 40
        .Column = 15
        lngAssignX = lngProcessTime * 1000 '60    '120 sec 120 * 60 = 7200
        lngAssignY = lngMaxTemp * 10    '1200 degree 1200*10= 12000
        intColumnPitch = lngAssignX / (gbProcessLogTable.Column + 1)
        intRowPitch = lngAssignY / gbProcessLogTable.row * 2
        .Xscale_size = lngAssignX + intColumnPitch
        .Yscale_size = lngAssignY + intRowPitch
    End With
    lnGridW = lngAssignX
    
    
    picProcessLog(0).Cls
    picProcessLog(0).DrawWidth = 1
    lngTable_Xsize = gbProcessLogTable.Xscale_size - intColumnPitch 'Table's Weight
    lngTable_Ysize = gbProcessLogTable.Yscale_size - intRowPitch 'Table's Height
    sngXPitch = (lngTable_Xsize) / gbProcessLogTable.Column
    sngYPitch = (lngTable_Ysize) / gbProcessLogTable.row
    iCount = 0
    'This loop  plot the lines
    For lngY = 0 To (lngTable_Ysize) Step sngYPitch
        For lngX = 0 To (lngTable_Xsize) Step sngXPitch
            picProcessLog(0).Line (lngX, 0)-(lngX, lngTable_Ysize)
        Next lngX
        
        If 0 = (iCount Mod 2) Then
            Color_Choice = vbBlack
        Else
            Color_Choice = GB_ColorLightBlue 'GB_ColorLightGray
        End If
        picProcessLog(0).Line (0, lngY)-(lngTable_Xsize, lngY), Color_Choice
        iCount = iCount + 1
    Next lngY

    Zoom_ProcessLogTable = True
End Function

Public Function Plot_ProcessLogTable(lngProcessTime As Long, lngMaxTemp As Long) As Boolean
    Dim lngX As Long
    Dim lngY As Long

    Dim iCount As Integer
    Dim PercentPitch  As Integer
    Dim PercentStep As Integer
    Dim TempPitch As Integer
    Dim TempStep As Integer
    Dim TimePitch As Integer
    Dim TimeStep As Integer
    Dim m_Min As Integer
    Dim m_Sec As Integer
    Dim iScaleX As Integer
    Dim iScaleY As Integer

    Dim i As Integer
    
    With gbProcessLogTable
        .Xsize = 10000
        .Ysize = 8000
        .ExtendXY_Size = 500
        .row = 40
        .Column = 15
        lngAssignX = lngProcessTime * 1000 '60    '120 sec 120 * 60 = 7200
'        lngAssignY = lngMaxTemp * 10    '1200 degree 1200*10= 12000
        If lngMaxTemp < 760 Then
        lngAssignY = 760 * 10
        Else
        lngAssignY = lngMaxTemp * 10
        End If
        intColumnPitch = lngAssignX / (gbProcessLogTable.Column + 1)
        intRowPitch = lngAssignY / gbProcessLogTable.row * 2
        .Xscale_size = lngAssignX + intColumnPitch
        .Yscale_size = lngAssignY + intRowPitch
    End With
    lnGridW = lngAssignX
    lnGridH = lngAssignY
'    picProcessLog(0).FontSize = 12
'    picProcessLog(0).Width = gbProcessLogTable.Xsize + gbProcessLogTable.ExtendXY_Size * 2
'    picProcessLog(0).Height = gbProcessLogTable.Ysize + gbProcessLogTable.ExtendXY_Size * 2
'    picProcessLog(0).Scale (-intColumnPitch * 2, gbProcessLogTable.Yscale_size + intRowPitch) _
'                        -(gbProcessLogTable.Xscale_size + intColumnPitch, -intRowPitch)
    
    If blnZoom = False Then
        picProcessLog(0).Scale (-intColumnPitch * 2, gbProcessLogTable.Yscale_size + intRowPitch) _
                            -(gbProcessLogTable.Xscale_size + intColumnPitch, -intRowPitch)
    Else
        picProcessLog(0).Scale (ZoomRect.Left, ZoomRect.Top) _
                            -(ZoomRect.Right, ZoomRect.Bottom)
    End If
'    fraProcessLogChart.Height = picProcessLog(0).Height + 500
'    fraProcessLogChart.Width = picProcessLog(0).Width + 500
    
    picProcessLog(0).Cls
    picProcessLog(0).DrawWidth = 1
    lngTable_Xsize = gbProcessLogTable.Xscale_size - intColumnPitch 'Table's Weight
    lngTable_Ysize = gbProcessLogTable.Yscale_size - intRowPitch 'Table's Height
    sngXPitch = (lngTable_Xsize) / gbProcessLogTable.Column
    sngYPitch = (lngTable_Ysize) / gbProcessLogTable.row
    iCount = 0
    'This loop  plot the lines
    For lngY = 0 To (lngTable_Ysize) Step sngYPitch
        For lngX = 0 To (lngTable_Xsize) Step sngXPitch
            picProcessLog(0).Line (lngX, 0)-(lngX, lngTable_Ysize)
        Next lngX
        
        If 0 = (iCount Mod 2) Then
            Color_Choice = vbBlack
        Else
            Color_Choice = GB_ColorLightBlue 'GB_ColorLightGray
        End If
        picProcessLog(0).Line (0, lngY)-(lngTable_Xsize, lngY), Color_Choice
        iCount = iCount + 1
    Next lngY

   'Plot temperature,percent and sec. number

   PercentPitch = Val(Set_Percent) / (gbProcessLogTable.row / 2)
'   TempPitch = Val(Set_Temp) / (gbProcessLogTable.row / 2)
   If Set_Temp > 760 Then
    TempPitch = Val(Set_Temp) / (gbProcessLogTable.row / 2)
    Else
     TempPitch = 760 / (gbProcessLogTable.row / 2)
   End If
   For i = 0 To 20
        'Percent number display
        PercentStep = i * PercentPitch
        picProcessLog(0).PSet (lngTable_Xsize + 10, (i * sngYPitch * 2) + _
                                ScaleY((picProcessLog(0).FontSize / 2 * lngTable_Ysize / picProcessLog(0).Height) _
                                     , vbPoints, vbTwips)) 'sign point and get the currentX,Y
        If i = 20 Then
            picProcessLog(0).Print str(PercentStep) & "%"
        Else
            picProcessLog(0).Print str(PercentStep)
        End If
        'Temperature number Display
        TempStep = i * TempPitch
        picProcessLog(0).CurrentX = -sngXPitch '-ScaleX(picProcessLog(0).FontSize, vbPoints, vbTwips) * 3.5
        picProcessLog(0).CurrentY = (i * sngYPitch * 2) + _
                                     ScaleY((picProcessLog(0).FontSize / 2 * lngTable_Ysize / picProcessLog(0).Height) _
                                     , vbPoints, vbTwips)
         If i = 0 Then
            picProcessLog(0).Print str(TempStep) & " (C)"
        Else
            picProcessLog(0).Print str(TempStep)
        End If
    Next i
    
    'For time sec display
    TimePitch = Val(Set_Time) / gbProcessLogTable.Column
    For i = 0 To 15
        TimeStep = i * TimePitch
        m_Min = Fix(TimeStep / 60)
        m_Sec = (TimeStep Mod 60)
        picProcessLog(0).CurrentX = (i * sngXPitch) - sngXPitch / 2 'ScaleY((picProcessLog(0).FontSize), vbPoints, vbTwips)
        picProcessLog(0).CurrentY = -intRowPitch / 3 'table axis'
        picProcessLog(0).Print str(m_Min) & ":" & str(m_Sec)
    Next i

    Plot_ProcessLogTable = True
End Function



Public Function PlotProcessLogChart() As Boolean
    
    Dim picTemp(1) As PictureBox
    Dim tmpStr      As String
    Dim Ci, Cj, Ck  As Long
    Dim tmpVal, SelPic     As Long
    Dim GetPlotData(50) As Single
    Dim i As Long, j As Long, k As Long, m As Long
    Dim lngPlot_Scale As Long
    Dim TempTC(50) As Single '120606 Josh added
    Dim temp As Integer
    Dim mc(0 To 7) As Long

    ' ?置?色常量??
    mc(0) = vbRed
    mc(1) = &HFF00&
    mc(2) = &HFFFF00
    mc(3) = &HFFFF&
    mc(4) = &H80FF&
    mc(5) = &HFF00FF
    mc(6) = &HFF8080
    mc(7) = &H8080FF
    
    Erase GetPlotData
    cmdOpenLog.Enabled = False
    chkZoom.Enabled = False
    cmdZoomBack.Enabled = False
    
    lngGasProfileColor(0) = GB_ColorPurple
    lngGasProfileColor(1) = GB_ColorGreen
    lngGasProfileColor(2) = GB_ColorFewLightBlue
    lngGasProfileColor(3) = &HFF00&
    lngGasProfileColor(4) = &HFFFF80
    For i = 0 To 39
        lngPlot_X_Point(i) = 0
        lngPlot_Y_Point(i) = 0
    Next i
    Set picTemp(0) = Me.picProcessLog(0)
        
    Call Plot_RecipeLog

    linCourseHor.Visible = False
    linCourseVer.Visible = False
    lbTemp.Visible = False
    lbSec.Visible = False

        
    picProcessLog(0).DrawWidth = 3
    If Set_Temp < 760 Then
    lngPlot_Scale = CLng(lngAssignY / 760)
    Else
    lngPlot_Scale = CLng(lngAssignY / Set_Temp)
    End If
'    Dim RecordCount As Integer
'    RecordCount = UBound(gbsngLogProcessRecord)
'    lngPlot_Scale = CLng(lngAssignY / Set_Temp)
'    For i = 0 To 65535
     For i = 0 To GbLogRcdCount
        If i > 0 And gbsngLogProcessRecord(i, 0) = 0 Then
            PlotProcessLogChart = True
            Exit Function
        End If
        If i = 2096 Then
            i = i
        End If
        
        GetPlotData(0) = Val(gbsngLogProcessRecord(i, 0)) 'Time
        If GetPlotData(0) > Set_Time * 1000 Then
         Exit For
        End If
        GetPlotData(1) = Val(gbsngLogProcessRecord(i, 1)) * (lngAssignY / 100) * 10 'INTENSITY1
        GetPlotData(17) = Val(gbsngLogProcessRecord(i, 17)) * (lngAssignY / 100) * 10 'INTENSITY2
        GetPlotData(18) = Val(gbsngLogProcessRecord(i, 18)) * (lngAssignY / 100) * 10 'INTENSITY3
        GetPlotData(19) = Val(gbsngLogProcessRecord(i, 19)) * (lngAssignY / 100) * 10 'INTENSITY4
        GetPlotData(20) = Val(gbsngLogProcessRecord(i, 20)) * (lngAssignY / 100) * 10 'INTENSITY5
        GetPlotData(2) = Val(gbsngLogProcessRecord(i, 2)) * lngPlot_Scale          'TC
        GetPlotData(3) = Val(gbsngLogProcessRecord(i, 3)) * lngPlot_Scale           'PM
        
     
        
        For j = 4 To 8
         If gbintGasEnable(j - 4) > 0 And gbsngMaxGasSLMP(j - 4) > 0 Then GetPlotData(j) = Val(gbsngLogProcessRecord(i, j)) * (5 / gbsngMaxGasSLMP(j - 4)) * (lngAssignY / 50) * gbintMFC_Ratio 'Gas
'           If gbintGasEnable(j - 4) > 0 And gbsngMaxGasSLMP(j - 4) > 0 Then GetPlotData(j) = Val(gbsngLogProcessRecord(i, j)) * lngPlot_Scale * (lngAssignY / 100 / 10) * gbintMFC_Ratio
            If GetPlotData(j) > lngAssignY Then
                GetPlotData(j) = lngAssignY
            End If
        Next j
         If gbintGasEnable(5) > 0 And gbsngMaxGasSLMP(5) > 0 Then GetPlotData(10) = Val(gbsngLogProcessRecord(i, 10)) * (5 / gbsngMaxGasSLMP(5)) * (lngAssignY / 50) * gbintMFC_Ratio 'Gas6
'            If gbintGasEnable(5) > 0 And gbsngMaxGasSLMP(5) > 0 Then GetPlotData(10) = Val(gbsngLogProcessRecord(i, 10)) * lngPlot_Scale * (lngAssignY / 100 / 10) * gbintMFC_Ratio
            If GetPlotData(10) > lngAssignY Then
                GetPlotData(10) = lngAssignY
            End If
        
        GetPlotData(9) = Val(gbsngLogProcessRecord(i, 9))
      
        
        
'        If gbsngGaugeZoomIn > 0 And GetPlotData(9) < gbsngGaugeZoomIn Then
'            GetPlotData(9) = GetPlotData(9) * 1000
'        End If
        
'        If GetPlotData(9) > 760 Then
'            GetPlotData(9) = 760
'        End If
        GetPlotData(9) = GetPlotData(9) * lngPlot_Scale          'Vacuum
        If GetPlotData(9) > lngAssignY Then
            GetPlotData(9) = lngAssignY
        End If
        
        
        
        '120606 Josh Modified
        For j = 11 To 15
            TempTC(j - 11) = 0
            If Val(gbsngLogProcessRecord(i, j)) > 0 Then
                TempTC(j - 11) = Val(gbsngLogProcessRecord(i, j))
            End If
            GetPlotData(j) = TempTC(j - 11) * lngPlot_Scale
        Next j
        
         '170201 Josh Modified
        For j = 22 To 23
            TempTC(j - 22) = 0
            If Val(gbsngLogProcessRecord(i, j)) > 0 Then
                TempTC(j - 22) = Val(gbsngLogProcessRecord(i, j))
            End If
            GetPlotData(j) = TempTC(j - 22) * lngPlot_Scale
        Next j
        
        If Para.UseMTC = 1 Then
            For j = 24 To 31
                TempTC(0) = 0
                If Val(gbsngLogProcessRecord(i, j)) > 0 Then
                    TempTC(0) = Val(gbsngLogProcessRecord(i, j))
                End If
                GetPlotData(j) = TempTC(0) * lngPlot_Scale
            Next j
        End If
        If Para.UseMTCB = 1 Then
            For j = 32 To 39
                TempTC(0) = 0
                If Val(gbsngLogProcessRecord(i, j)) > 0 Then
                    TempTC(0) = Val(gbsngLogProcessRecord(i, j))
                End If
                GetPlotData(j) = TempTC(0) * lngPlot_Scale
            Next j
        End If
        
        For m = 0 To 39
            If GetPlotData(m) < -10 Or GetPlotData(m) > 6553500 Then
                GetPlotData(m) = 0
            End If
        Next m
        
        TimeMap(i) = GetPlotData(0)
        
        tmpStr = GetPlotData(0)
        SelPic = 0
        If Val(tmpStr) >= 0 Then
            picProcessLog(0).DrawWidth = 3
            If chkPlotColor(0).value = 1 Then _
                picProcessLog(0).Line (lngPlot_X_Point(1), lngPlot_Y_Point(1))-(GetPlotData(0), GetPlotData(1)), &HC0FFFF       'Intensity1
            If chkPlotColor(31).value = 1 Then _
                picProcessLog(0).Line (lngPlot_X_Point(1), lngPlot_Y_Point(1))-(GetPlotData(0), GetPlotData(17)), &H80FFFF    'Intensity2
                If chkPlotColor(32).value = 1 Then _
                picProcessLog(0).Line (lngPlot_X_Point(1), lngPlot_Y_Point(1))-(GetPlotData(0), GetPlotData(18)), &HFFFF&     'Intensity3
                If chkPlotColor(33).value = 1 Then _
                picProcessLog(0).Line (lngPlot_X_Point(1), lngPlot_Y_Point(1))-(GetPlotData(0), GetPlotData(19)), &HC0C0&     'Intensity4
                If chkPlotColor(34).value = 1 Then _
                picProcessLog(0).Line (lngPlot_X_Point(1), lngPlot_Y_Point(1))-(GetPlotData(0), GetPlotData(20)), &H8080&     'Intensity5
                
            If ChkGasColor(0).value = 1 Then _
                picProcessLog(0).Line (lngPlot_X_Point(4), lngPlot_Y_Point(4))-(GetPlotData(0), GetPlotData(4)), ChkGasColor(0).BackColor 'GAS1
            If ChkGasColor(1).value = 1 Then _
                picProcessLog(0).Line (lngPlot_X_Point(5), lngPlot_Y_Point(5))-(GetPlotData(0), GetPlotData(5)), ChkGasColor(1).BackColor 'GAS2
            If ChkGasColor(2).value = 1 Then _
                picProcessLog(0).Line (lngPlot_X_Point(6), lngPlot_Y_Point(6))-(GetPlotData(0), GetPlotData(6)), ChkGasColor(2).BackColor 'GAS3
            If ChkGasColor(3).value = 1 Then _
                picProcessLog(0).Line (lngPlot_X_Point(7), lngPlot_Y_Point(7))-(GetPlotData(0), GetPlotData(7)), ChkGasColor(3).BackColor 'GAS4
            If ChkGasColor(4).value = 1 Then _
                picProcessLog(0).Line (lngPlot_X_Point(8), lngPlot_Y_Point(8))-(GetPlotData(0), GetPlotData(8)), ChkGasColor(4).BackColor 'GAS5
            If UCase(ChkGasColor(5).Caption) <> "NA" And ChkGasColor(5).value = 1 Then
               If UCase(ChkGasColor(5).Caption) = "VAC" Then
               picProcessLog(0).Line (lngPlot_X_Point(9), lngPlot_Y_Point(9))-(GetPlotData(0), GetPlotData(9)), ChkGasColor(5).BackColor
               Else
                picProcessLog(0).Line (lngPlot_X_Point(10), lngPlot_Y_Point(10))-(GetPlotData(0), GetPlotData(10)), ChkGasColor(5).BackColor 'GAS6
               End If
            End If
            If UCase(ChkGasColor(6).Caption) <> "NA" And ChkGasColor(6).value = 1 Then
            If UCase(ChkGasColor(6).Caption) = "VAC" Then
               picProcessLog(0).Line (lngPlot_X_Point(9), lngPlot_Y_Point(9))-(GetPlotData(0), GetPlotData(9)), ChkGasColor(6).BackColor
               Else
                picProcessLog(0).Line (lngPlot_X_Point(10), lngPlot_Y_Point(10))-(GetPlotData(0), GetPlotData(10)), ChkGasColor(6).BackColor 'GAS6
            End If
            
            End If
             
'             If ChkGasColor(5).value = 1 Then _
'                 picProcessLog(0).Line (lngPlot_X_Point(10), lngPlot_Y_Point(10))-(GetPlotData(0), GetPlotData(10)), ChkGasColor(5).BackColor 'GAS6
'            If ChkGasColor(6).value = 1 Then _
'                picProcessLog(0).Line (lngPlot_X_Point(9), lngPlot_Y_Point(9))-(GetPlotData(0), GetPlotData(9)), ChkGasColor(6).BackColor    'Vac
'
            picProcessLog(0).DrawWidth = 1
            
            If chkPlotColor(7).value = 1 Then _
                picProcessLog(0).Line (lngPlot_X_Point(2), lngPlot_Y_Point(2))-(GetPlotData(0), GetPlotData(2)), vbRed 'TC1
            If chkPlotColor(8).value = 1 Then _
                picProcessLog(0).Line (lngPlot_X_Point(11), lngPlot_Y_Point(11))-(GetPlotData(0), GetPlotData(11)), &HFF00&         'TC2
            If chkPlotColor(9).value = 1 Then _
                picProcessLog(0).Line (lngPlot_X_Point(12), lngPlot_Y_Point(12))-(GetPlotData(0), GetPlotData(12)), &HFFFF00        'TC3
            If chkPlotColor(10).value = 1 Then _
                picProcessLog(0).Line (lngPlot_X_Point(13), lngPlot_Y_Point(13))-(GetPlotData(0), GetPlotData(13)), &HFFFF&         'TC4
            If chkPlotColor(11).value = 1 Then _
                picProcessLog(0).Line (lngPlot_X_Point(14), lngPlot_Y_Point(14))-(GetPlotData(0), GetPlotData(14)), &H80FF&         'TC5
            If chkPlotColor(12).value = 1 Then _
                picProcessLog(0).Line (lngPlot_X_Point(15), lngPlot_Y_Point(15))-(GetPlotData(0), GetPlotData(15)), &HFF00FF        'TC6
            If chkPlotColor(13).value = 1 Then _
                picProcessLog(0).Line (lngPlot_X_Point(22), lngPlot_Y_Point(22))-(GetPlotData(0), GetPlotData(22)), &HFF8080        'TC7
            If chkPlotColor(14).value = 1 Then _
                picProcessLog(0).Line (lngPlot_X_Point(23), lngPlot_Y_Point(23))-(GetPlotData(0), GetPlotData(23)), &H8080FF        'TC8
            
            If Para.UseMTC = 1 Then
                If chkPlotColor(15).value = 1 Then _
                    picProcessLog(0).Line (lngPlot_X_Point(24), lngPlot_Y_Point(24))-(GetPlotData(0), GetPlotData(24)), vbRed         'MTC9
                If chkPlotColor(16).value = 1 Then _
                    picProcessLog(0).Line (lngPlot_X_Point(25), lngPlot_Y_Point(25))-(GetPlotData(0), GetPlotData(25)), &HFF00&       'MTC10
                If chkPlotColor(17).value = 1 Then _
                    picProcessLog(0).Line (lngPlot_X_Point(26), lngPlot_Y_Point(26))-(GetPlotData(0), GetPlotData(26)), &HFFFF00      'MTC11
                If chkPlotColor(18).value = 1 Then _
                    picProcessLog(0).Line (lngPlot_X_Point(27), lngPlot_Y_Point(27))-(GetPlotData(0), GetPlotData(27)), &HFFFF&       'MTC12
                If chkPlotColor(19).value = 1 Then _
                    picProcessLog(0).Line (lngPlot_X_Point(28), lngPlot_Y_Point(28))-(GetPlotData(0), GetPlotData(28)), &H80FF&       'MTC13
                If chkPlotColor(20).value = 1 Then _
                    picProcessLog(0).Line (lngPlot_X_Point(29), lngPlot_Y_Point(29))-(GetPlotData(0), GetPlotData(29)), &HFF00FF      'MTC14
                If chkPlotColor(21).value = 1 Then _
                    picProcessLog(0).Line (lngPlot_X_Point(30), lngPlot_Y_Point(30))-(GetPlotData(0), GetPlotData(30)), &HFF8080      'MTC15
                If chkPlotColor(22).value = 1 Then _
                    picProcessLog(0).Line (lngPlot_X_Point(31), lngPlot_Y_Point(31))-(GetPlotData(0), GetPlotData(31)), &H8080FF      'MTC16
            End If

            If Para.UseMTCB = 1 Then
                For j = 23 To 30
                    If GetPlotData(j + 9) < 13720 Then
                    If chkPlotColor(j).value = 1 Then _
                               picProcessLog(0).Line (lngPlot_X_Point(j + 9), lngPlot_Y_Point(j + 9))-(GetPlotData(0), GetPlotData(j + 9)), chkPlotColor(j).BackColor
'                    picProcessLog(0).Line (lngPlot_X_Point(j + 9), lngPlot_Y_Point(j + 9))-(GetPlotData(0), GetPlotData(j + 9)), mc(j - 23) 'MTC16
                    End If
                Next j
            End If
            

            For j = 0 To GB_MAX_DRAW_COL
                lngPlot_X_Point(j) = CLng(GetPlotData(0))
                lngPlot_Y_Point(j) = CLng(GetPlotData(j))
            Next j


        Else
            Erase GetPlotData
            Exit Function
        End If
    
        Erase GetPlotData
        DoEvents
    Next i
    
    PlotProcessLogChart = True
End Function

'============================================================================================================================
Public Function Plot_RecipeLog() As Boolean
    Dim bRet As Boolean
    Dim Check_Status As String
    Dim CalCounts As Integer
    Dim i As Integer
    Dim j As Integer
    Dim SysAction As String
    
    On Error GoTo ERR_PLOT_RECIPELOG
    
    bRet = False

    Check_Status = ""
    Set_Time = CHART_DEF_TIME
    Set_Temp = CHART_DEF_TEMP
    CalCounts = 0
    m_PlotXY(0).rcpX = 0
    m_PlotXY(0).rcpY = 0
      
    Check_Status = Trim(CStr(gbLogRecipe.arrayRecipe(1, GB_PROCESS_ACTION)))
    SysAction = Readini(Check_Status)
    If Readini(Check_Status) <> "" Then
     Check_Status = SysAction
    End If
    If Check_Status = "STOP" Or Check_Status = "" Then GoTo ERR_PLOT_RECIPELOG
    
    With gbLogRecipe
        lngMaxTime = 0
        lngMaxTemp = 0
        For i = 1 To 50
            lngMaxTime = lngMaxTime + CLng(gbLogRecipe.arrayRecipe(i, GB_PROCESS_TIME))
            lngMaxTemp = IIf(.arrayRecipe(i, GB_PROCESS_TEMP) >= lngMaxTemp, .arrayRecipe(i, GB_PROCESS_TEMP), lngMaxTemp)
        Next i
        'Adjust PIDTable's temperature display
        If lngMaxTemp > 760 Then
           Set_Temp = (lngMaxTemp \ 100) * 100 + 300
         Else
          Set_Temp = 760
        End If
     

        Set_Time = (lngMaxTime \ 15) * 15 + 15
        bRet = Plot_ProcessLogTable(CLng(Set_Time), CLng(Set_Temp))
        
        For i = 1 To GB_MAX_STEP_PROCESS
            CalCounts = CalCounts + 1
            Check_Status = Trim(CStr(.arrayRecipe(i, GB_PROCESS_ACTION)))
            SysAction = Readini(Check_Status)
            If Readini(Check_Status) <> "" Then
            Check_Status = SysAction
            End If
            Select Case Check_Status
                Case "Idle"
                    m_SecPoint(i) = .arrayRecipe(i, GB_PROCESS_TIME)
                    m_TempPoint(i) = .arrayRecipe(i, GB_PROCESS_TEMP)
                    If i = 1 Then
                        tmpX = m_PlotXY(CalCounts - 1).rcpX + ((m_SecPoint(i) * (lngTable_Xsize / Val(Set_Time))))
                        tmpY = 0 'm_TempPoint(i) * (lngTable_Ysize / Val(Set_Temp))
                        m_PlotXY(CalCounts).rcpX = tmpX
                        m_PlotXY(CalCounts).rcpY = tmpY
                    Else
                        tmpX = m_PlotXY(CalCounts - 1).rcpX
                        tmpY = 0 'm_TempPoint(i) * (lngTable_Ysize / Val(Set_Temp))
                        m_PlotXY(CalCounts).rcpX = tmpX
                        m_PlotXY(CalCounts).rcpY = tmpY
                        CalCounts = CalCounts + 1
                        tmpX = m_PlotXY(CalCounts - 1).rcpX + ((m_SecPoint(i) * (lngTable_Xsize / Val(Set_Time))))
                        tmpY = 0 'm_TempPoint(i) * (lngTable_Ysize / Val(Set_Temp))
                        m_PlotXY(CalCounts).rcpX = tmpX
                        m_PlotXY(CalCounts).rcpY = tmpY
                    End If
                    
                Case "PreHeat"
                    'CalCounts = CalCounts - 1
                    m_SecPoint(i) = gbLogRecipe.arrayRecipe(i, GB_PROCESS_TIME)
                    m_TempPoint(i) = gbLogRecipe.arrayRecipe(i, GB_PROCESS_TEMP)
                    tmpX = m_PlotXY(CalCounts - 1).rcpX + ((m_SecPoint(i) * (lngTable_Xsize / Val(Set_Time))))
                    tmpY = m_TempPoint(i) * (lngTable_Ysize / Val(Set_Temp))
                    m_PlotXY(CalCounts).rcpX = tmpX
                    m_PlotXY(CalCounts).rcpY = tmpY
                Case "Ramp up"
                    m_SecPoint(i) = gbLogRecipe.arrayRecipe(i, GB_PROCESS_TIME)
                    m_TempPoint(i) = gbLogRecipe.arrayRecipe(i, GB_PROCESS_TEMP)
    '                If i > 0 Then
    '                    m_Slope = (m_TempPoint(i) - m_TempPoint(i - 1)) / m_SecPoint(i)
    '                Else
    '                    m_Slope = (m_TempPoint(i) - 0) / m_SecPoint(i)
    '                End If
                    tmpX = m_PlotXY(CalCounts - 1).rcpX + m_SecPoint(i) * (lngTable_Xsize / Val(Set_Time))
                    tmpY = m_TempPoint(i) * (lngTable_Ysize / Val(Set_Temp))
                    m_PlotXY(CalCounts).rcpX = tmpX
                    m_PlotXY(CalCounts).rcpY = tmpY
                Case "Ramp Down"
                    m_SecPoint(i) = gbLogRecipe.arrayRecipe(i, GB_PROCESS_TIME)
                    m_TempPoint(i) = gbLogRecipe.arrayRecipe(i, GB_PROCESS_TEMP)
    '                If i > 0 Then
    '                    m_Slope = (m_TempPoint(i) - m_TempPoint(i - 1)) / m_SecPoint(i)
    '                Else
    '                    m_Slope = (m_TempPoint(i) - 0) / m_SecPoint(i)
    '                End If
                    tmpX = m_PlotXY(CalCounts - 1).rcpX + m_SecPoint(i) * (lngTable_Xsize / Val(Set_Time))
                    tmpY = m_TempPoint(i) * (lngTable_Ysize / Val(Set_Temp))
                    m_PlotXY(CalCounts).rcpX = tmpX
                    m_PlotXY(CalCounts).rcpY = tmpY
                                                 
                Case "Hold"
                    m_SecPoint(i) = gbLogRecipe.arrayRecipe(i, GB_PROCESS_TIME)
                    m_TempPoint(i) = gbLogRecipe.arrayRecipe(i, GB_PROCESS_TEMP)
                    tmpX = m_PlotXY(CalCounts - 1).rcpX + m_SecPoint(i) * (lngTable_Xsize / Val(Set_Time))
                    tmpY = m_TempPoint(i) * (lngTable_Ysize / Val(Set_Temp))
                    m_PlotXY(CalCounts).rcpX = tmpX
                    m_PlotXY(CalCounts).rcpY = tmpY
                    
                Case "Vent"
                    m_SecPoint(i) = .arrayRecipe(i, GB_PROCESS_TIME)
                    m_TempPoint(i) = .arrayRecipe(i, GB_PROCESS_TEMP)
                    If i = 1 Then
                        tmpX = m_PlotXY(CalCounts - 1).rcpX + ((m_SecPoint(i) * (lngTable_Xsize / Val(Set_Time))))
                        tmpY = 10 'm_TempPoint(i) * (lngTable_Ysize / Val(Set_Temp))
                        m_PlotXY(CalCounts - 1).rcpY = tmpY
                        m_PlotXY(CalCounts).rcpX = tmpX
                        m_PlotXY(CalCounts).rcpY = tmpY
                    Else
                        tmpX = m_PlotXY(CalCounts - 1).rcpX
                        tmpY = 10 'm_TempPoint(i) * (lngTable_Ysize / Val(Set_Temp))
                        m_PlotXY(CalCounts).rcpX = tmpX
                        m_PlotXY(CalCounts).rcpY = tmpY
                        CalCounts = CalCounts + 1
                        tmpX = m_PlotXY(CalCounts - 1).rcpX + ((m_SecPoint(i) * (lngTable_Xsize / Val(Set_Time))))
                        tmpY = 10 'm_TempPoint(i) * (lngTable_Ysize / Val(Set_Temp))
                        m_PlotXY(CalCounts).rcpX = tmpX
                        m_PlotXY(CalCounts).rcpY = tmpY
                    End If
                    
                Case "Purge"
                    m_SecPoint(i) = .arrayRecipe(i, GB_PROCESS_TIME)
                    m_TempPoint(i) = .arrayRecipe(i, GB_PROCESS_TEMP)
                    If i = 1 Then
                        tmpX = m_PlotXY(CalCounts - 1).rcpX + ((m_SecPoint(i) * (lngTable_Xsize / Val(Set_Time))))
                        tmpY = 10 'm_TempPoint(i) * (lngTable_Ysize / Val(Set_Temp))
                        m_PlotXY(CalCounts - 1).rcpY = tmpY
                        m_PlotXY(CalCounts).rcpX = tmpX
                        m_PlotXY(CalCounts).rcpY = tmpY
                    Else
                        tmpX = m_PlotXY(CalCounts - 1).rcpX
                        tmpY = 10 'm_TempPoint(i) * (lngTable_Ysize / Val(Set_Temp))
                        m_PlotXY(CalCounts).rcpX = tmpX
                        m_PlotXY(CalCounts).rcpY = tmpY
                        CalCounts = CalCounts + 1
                        tmpX = m_PlotXY(CalCounts - 1).rcpX + ((m_SecPoint(i) * (lngTable_Xsize / Val(Set_Time))))
                        tmpY = 10 'm_TempPoint(i) * (lngTable_Ysize / Val(Set_Temp))
                        m_PlotXY(CalCounts).rcpX = tmpX
                        m_PlotXY(CalCounts).rcpY = tmpY
                    End If
                
                Case "Stop"
                    m_SecPoint(i) = gbLogRecipe.arrayRecipe(i, GB_PROCESS_TIME)
                    m_TempPoint(i) = gbLogRecipe.arrayRecipe(i, GB_PROCESS_TEMP)
                    tmpX = m_PlotXY(CalCounts - 1).rcpX + m_SecPoint(i) * (lngTable_Xsize / Val(Set_Time))
                    tmpY = m_TempPoint(i) * (lngTable_Ysize / Val(Set_Temp))
                    m_PlotXY(CalCounts).rcpX = tmpX
                    m_PlotXY(CalCounts).rcpY = tmpY
                    Exit For
                    
                Case "Pump Down"
                    m_SecPoint(i) = .arrayRecipe(i, GB_PROCESS_TIME)
                    m_TempPoint(i) = .arrayRecipe(i, GB_PROCESS_TEMP)
                    If i = 1 Then
                        tmpX = m_PlotXY(CalCounts - 1).rcpX + ((m_SecPoint(i) * (lngTable_Xsize / Val(Set_Time))))
                    Else
                        tmpX = m_PlotXY(CalCounts - 1).rcpX '+ ((m_SecPoint(i) * (lngTable_Xsize / Val(Set_Time))))
                    End If
                    tmpY = 0 'm_TempPoint(i) * (lngTable_Ysize / Val(Set_Temp))
                    m_PlotXY(CalCounts).rcpX = tmpX
                    m_PlotXY(CalCounts).rcpY = tmpY
                
            End Select
        Next i
        'Call GetMaxTempAndTime(RcpXY)
    
        
        DoEvents        ' If coding error ..be saviful
    
        bRet = False
        'Plot Line of the Recipe
        picProcessLog(0).DrawWidth = 3
        If m_PlotXY(1).rcpX > 0 Then
                picProcessLog(0).Line (m_PlotXY(0).rcpX, m_PlotXY(0).rcpY)-(m_PlotXY(1).rcpX, m_PlotXY(1).rcpY), vbBlue
        End If
    
        For j = 1 To CalCounts - 1
            picProcessLog(0).Line (m_PlotXY(j).rcpX, m_PlotXY(j).rcpY) _
                                                    -(m_PlotXY(j + 1).rcpX, m_PlotXY(j + 1).rcpY), vbBlue
        Next j
    
    End With
    Plot_RecipeLog = True
    Exit Function
    'frmPlotProcess.tabPlotProcess.Tab = tmpTable
ERR_PLOT_RECIPELOG:
    'Call AlertShow("Plot Recipe Chart Error!!", ERRORTYPE)
    Plot_RecipeLog = False
End Function
'============================================================================================================================
Public Sub SaveCurProcessLog(IsProcessLog As Boolean)
    Dim i               As Long
    Dim j               As Integer
    Dim lngRet          As Long
    Dim StrFileName     As String
    Dim strLogFileName     As String
    Dim iInputDevice    As Integer
    Dim strFilePath     As String
    Dim strDir              As String
    Dim strTemp         As String
    
    
    On Error GoTo ERR_PROCESSLOG_SAVE
    If IsProcessLog = False Then
    strFilePath = gbSystemPath & "\Log"
    strDir = dir(strFilePath, vbDirectory)
    If strDir = "" Then MkDir strFilePath
    strFilePath = strFilePath & "\" & Year(Date)
    strDir = dir(strFilePath, vbDirectory)
    If strDir = "" Then MkDir strFilePath
    strFilePath = strFilePath & "\" & Month(Date)
    strDir = dir(strFilePath, vbDirectory)
    If strDir = "" Then MkDir strFilePath
    strFilePath = strFilePath & "\" & Day(Date)
    strDir = dir(strFilePath, vbDirectory)
    If strDir = "" Then MkDir strFilePath
    
    gbstrPN = frmPlotProcess.txtPN.text
    gbstrBN = frmPlotProcess.txtBN.text
    gbstrID1 = frmPlotProcess.txtID1.text
    gbstrID2 = frmPlotProcess.txtID2.text

    If gbintActiveModule_Barcode = 0 Or gbstrBN = "" Then
        StrFileName = strFilePath & "\" & Mid(frmPlotProcess.Caption, 1, InStr(1, frmPlotProcess.Caption, ".") - 1) & "_" & Format(Time, "hhmmss") & ".log"
    Else
        StrFileName = strFilePath & "\" & Mid(frmPlotProcess.Caption, 1, InStr(1, frmPlotProcess.Caption, ".") - 1) & "_" & Format(Time, "hhmmss") & "_" & gbstrBN & "_" & gbstrID1 & ".log"
    End If
    
       
    Kernel.strCurrLogFile = StrFileName
    SaveFileName = StrFileName
    lngRet = WritePrivateProfileString("COMMENT", "RecordTime", CStr(Now()), StrFileName)
    lngRet = WritePrivateProfileString("COMMENT", "InputDevice", CStr(frmRecipeEdit.intRecipeTempInputType), StrFileName)
    lngRet = WritePrivateProfileString("COMMENT", "InputObject", CStr(gbintPMDetectObject), StrFileName)
    lngRet = WritePrivateProfileString("COMMENT", "PN", gbstrPN, StrFileName)
    lngRet = WritePrivateProfileString("COMMENT", "BN", gbstrBN, StrFileName)
    lngRet = WritePrivateProfileString("COMMENT", "ID1", gbstrID1, StrFileName)
    lngRet = WritePrivateProfileString("COMMENT", "ID2", gbstrID2, StrFileName)
    lngRet = WritePrivateProfileString("COMMENT", "ColumnName", "Index,Time,Intensity,TC,Pump,Gas1~4,Pressure,MTC1~5,Power1~5,O2", StrFileName)
    
    lngRet = WritePrivateProfileString("COMMENT", "PumpDownTime", CStr(gblngPumpDownTime), StrFileName)
    
    For i = 1 To GB_MAX_STEP_PROCESS
        'Action
        lngRet = WritePrivateProfileString("STEP" & str(i), "ACTION", m_Recipe.arrayRecipe(i, GB_PROCESS_ACTION), StrFileName)
        'Temperature (degree)
        lngRet = WritePrivateProfileString("STEP" & str(i), "TEMP", m_Recipe.arrayRecipe(i, GB_PROCESS_TEMP), StrFileName)
        'Time (sec)
        lngRet = WritePrivateProfileString("STEP" & str(i), "TIME", m_Recipe.arrayRecipe(i, GB_PROCESS_TIME), StrFileName)
        For j = 0 To GB_GAS_MAX - 1
            lngRet = WritePrivateProfileString("STEP" & str(i), gbstrGasAlias(j), m_Recipe.arrayRecipe(i, GB_PROCESS_GAS1 + j), StrFileName)
        Next j
   Next i
    'gtcRTP 1.0.0
    lngRet = WritePrivateProfileString("COFFICIENT", "IntensityWeightD1", CStr(gbsngIntensityWeight(0)), StrFileName)
    lngRet = WritePrivateProfileString("COFFICIENT", "IntensityWeightD2", CStr(gbsngIntensityWeight(1)), StrFileName)
    lngRet = WritePrivateProfileString("COFFICIENT", "IntensityWeightD3", CStr(gbsngIntensityWeight(2)), StrFileName)
    lngRet = WritePrivateProfileString("COFFICIENT", "IntensityWeightD4", CStr(gbsngIntensityWeight(3)), StrFileName)
    lngRet = WritePrivateProfileString("COFFICIENT", "IntensityWeightD5", CStr(gbsngIntensityWeight(4)), StrFileName)
    
    lngRet = WritePrivateProfileString("COFFICIENT", "IntensityWeightS1", CStr(gbsngIntensityWeightS(0)), StrFileName)
    lngRet = WritePrivateProfileString("COFFICIENT", "IntensityWeightS2", CStr(gbsngIntensityWeightS(1)), StrFileName)
    lngRet = WritePrivateProfileString("COFFICIENT", "IntensityWeightS3", CStr(gbsngIntensityWeightS(2)), StrFileName)
    lngRet = WritePrivateProfileString("COFFICIENT", "IntensityWeightS4", CStr(gbsngIntensityWeightS(3)), StrFileName)
    lngRet = WritePrivateProfileString("COFFICIENT", "IntensityWeightS5", CStr(gbsngIntensityWeightS(4)), StrFileName)
    
    
    
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "Proportional", frmRecipeEdit.txtProportional.text, StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "Proportional2", frmRecipeEdit.txtProportional2.text, StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "Integrnal", frmRecipeEdit.txtIntegrnal.text, StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "Integral2", frmRecipeEdit.txtIntegral2.text, StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "Derivational", frmRecipeEdit.txtDerivational.text, StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "Factor1", frmRecipeEdit.txtPredit.text, StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "Factor2", frmRecipeEdit.txtFeedForward.text, StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "InputDevice", CStr(frmRecipeEdit.intRecipeTempInputType), StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "Overshoot", frmRecipeEdit.txtOvershoot.text, StrFileName)
    lngRet = WritePrivateProfileString("PROCESS_RECORD", "No", "Time,Intensity,TC1,NA,Gas1,Gas2,Gas3,Gas4,Gas5,Pressure,Gas6,TC2,TC3,TC4,TC5,TC6,Power1,Power2,Power3,Power4,Power5,O2-Gauge,TC7,O2,MTC8,MTC9,MTC10,MTC11,MTC12,MTC13,MTC14,MTC15,MTC16,MTC17,MTC18,MTC19,MTC20,MTC21,MTC22,MTC23", StrFileName)
    Else
     For i = 0 To GB_MAX_DRAW_COL + 1
       strTemp = strTemp + CStr(gbsngProcessRecorder(i)) & ","
     Next i
     strTemp = Left(strTemp, Len(strTemp) - 1)
     lngRet = WritePrivateProfileString("PROCESS_RECORD", CStr(RecordCount), strTemp, SaveFileName)
     RecordCount = RecordCount + 1
    If Not gbstrLogFilePath = "" Then
        If Not gbstrLogFilePath = "C:\Program Files\eRTA100" Then
            strFilePath = gbstrLogFilePath & "\Log"
            strDir = dir(strFilePath, vbDirectory)
            If strDir = "" Then MkDir strFilePath
            strFilePath = strFilePath & "\" & Year(Date)
            strDir = dir(strFilePath, vbDirectory)
            If strDir = "" Then MkDir strFilePath
            strFilePath = strFilePath & "\" & Month(Date)
            strDir = dir(strFilePath, vbDirectory)
            If strDir = "" Then MkDir strFilePath
            strFilePath = strFilePath & "\" & Day(Date)
            strDir = dir(strFilePath, vbDirectory)
            If strDir = "" Then MkDir strFilePath
            
            If gbintActiveModule_Barcode = 0 Or gbstrBN = "" Then
                strLogFileName = strFilePath & "\" & Mid(frmPlotProcess.Caption, 1, InStr(1, frmPlotProcess.Caption, ".") - 1) & "_" & Format(Time, "hhmmss") & ".log"
            Else
                strLogFileName = strFilePath & "\" & Mid(frmPlotProcess.Caption, 1, InStr(1, frmPlotProcess.Caption, ".") - 1) & "_" & Format(Time, "hhmmss") & "_" & gbstrBN & ".log"
            End If
            
            FileCopy SaveFileName, strLogFileName
             
             
            
        End If
    End If
    
    If Para.UseCIM = 1 And Para.intCIMPort = 1 Then
        strFilePath = gbstrLogFilePath & "\Log"
        strDir = dir(strFilePath, vbDirectory)
        If strDir = "" Then MkDir strFilePath
        strFilePath = strFilePath & "\" & Year(Date)
        strDir = dir(strFilePath, vbDirectory)
        If strDir = "" Then MkDir strFilePath
        strFilePath = strFilePath & "\" & Month(Date)
        strDir = dir(strFilePath, vbDirectory)
        If strDir = "" Then MkDir strFilePath
        strFilePath = strFilePath & "\" & Day(Date)
        strDir = dir(strFilePath, vbDirectory)
        If strDir = "" Then MkDir strFilePath
    
        CurrProc.strLogFilePath = strFilePath
        CurrProc.strLogFileName = Mid(frmPlotProcess.Caption, 1, InStr(1, frmPlotProcess.Caption, ".") - 1) & "_" & Format(Time, "hhmmss") & ".log"
    End If
    
    End If
    Exit Sub
ERR_PROCESSLOG_SAVE:
      WriteLog ("PROCESSLOG  Write Error")
End Sub


'============================================================================================================================
Public Sub SaveProcessLog()
    Dim i               As Long
    Dim j               As Integer
    Dim lngRet          As Long
    Dim StrFileName     As String
    Dim strLogFileName     As String
    Dim iInputDevice    As Integer
    Dim strFilePath     As String
    Dim strDir              As String
    Dim strTemp         As String
    
    On Error GoTo ERR_PROCESSLOG_SAVE
    
    Call frmConfiguration.StopWatchDog
    'strFilePath = gbSystemPath & "\Log" & "\" & Year(Date)
        
    strFilePath = gbSystemPath & "\Log"
    'strFilePath = gbstrLogFilePath & "\Log"
    strDir = dir(strFilePath, vbDirectory)
    If strDir = "" Then MkDir strFilePath
    strFilePath = strFilePath & "\" & Year(Date)
    strDir = dir(strFilePath, vbDirectory)
    If strDir = "" Then MkDir strFilePath
    strFilePath = strFilePath & "\" & Month(Date)
    strDir = dir(strFilePath, vbDirectory)
    If strDir = "" Then MkDir strFilePath
    strFilePath = strFilePath & "\" & Day(Date)
    strDir = dir(strFilePath, vbDirectory)
    If strDir = "" Then MkDir strFilePath
    
    gbstrPN = frmPlotProcess.txtPN.text
    gbstrBN = frmPlotProcess.txtBN.text
    gbstrID1 = frmPlotProcess.txtID1.text
    gbstrID2 = frmPlotProcess.txtID2.text
    'strFileName = strFilePath & "\" & Mid(frmPlotProcess.Caption, 1, InStr(1, frmPlotProcess.Caption, ".") - 1) & "_" & Format(Time, "hhmmss") & ".log"
    
    If gbintActiveModule_Barcode = 0 Or gbstrBN = "" Then
        StrFileName = strFilePath & "\" & Mid(frmPlotProcess.Caption, 1, InStr(1, frmPlotProcess.Caption, ".") - 1) & "_" & Format(Time, "hhmmss") & ".log"
    Else
        StrFileName = strFilePath & "\" & Mid(frmPlotProcess.Caption, 1, InStr(1, frmPlotProcess.Caption, ".") - 1) & "_" & Format(Time, "hhmmss") & "_" & gbstrBN & "_" & gbstrID1 & ".log"
    End If
    
       
    Kernel.strCurrLogFile = StrFileName
    
    lngRet = WritePrivateProfileString("COMMENT", "RecordTime", CStr(Now()), StrFileName)
    lngRet = WritePrivateProfileString("COMMENT", "InputDevice", CStr(frmRecipeEdit.intRecipeTempInputType), StrFileName)
    lngRet = WritePrivateProfileString("COMMENT", "InputObject", CStr(gbintPMDetectObject), StrFileName)
    lngRet = WritePrivateProfileString("COMMENT", "PN", gbstrPN, StrFileName)
    lngRet = WritePrivateProfileString("COMMENT", "BN", gbstrBN, StrFileName)
    lngRet = WritePrivateProfileString("COMMENT", "ID1", gbstrID1, StrFileName)
    lngRet = WritePrivateProfileString("COMMENT", "ID2", gbstrID2, StrFileName)
    lngRet = WritePrivateProfileString("COMMENT", "ColumnName", "Index,Time,Intensity,TC,Pump,Gas1~4,Pressure,MTC1~5,Power1~5,O2", StrFileName)
    
    lngRet = WritePrivateProfileString("COMMENT", "PumpDownTime", CStr(gblngPumpDownTime), StrFileName)
    
    For i = 1 To GB_MAX_STEP_PROCESS
        'Action
        lngRet = WritePrivateProfileString("STEP" & str(i), "ACTION", m_Recipe.arrayRecipe(i, GB_PROCESS_ACTION), StrFileName)
        'Temperature (degree)
        lngRet = WritePrivateProfileString("STEP" & str(i), "TEMP", m_Recipe.arrayRecipe(i, GB_PROCESS_TEMP), StrFileName)
        'Time (sec)
        lngRet = WritePrivateProfileString("STEP" & str(i), "TIME", m_Recipe.arrayRecipe(i, GB_PROCESS_TIME), StrFileName)
        For j = 0 To GB_GAS_MAX - 1
            lngRet = WritePrivateProfileString("STEP" & str(i), gbstrGasAlias(j), m_Recipe.arrayRecipe(i, GB_PROCESS_GAS1 + j), StrFileName)
        Next j
'        'Gas of N2
'        lngRet = WritePrivateProfileString("STEP" & Str(i), gbstrGasAlias(1), m_Recipe.arrayRecipe(i, GB_PROCESS_GAS1), strFileName)
'         'Gas of Ar
'        lngRet = WritePrivateProfileString("STEP" & Str(i), gbstrGasAlias(2), m_Recipe.arrayRecipe(i, GB_PROCESS_GAS2), strFileName)
'        'Gas of O2
'        lngRet = WritePrivateProfileString("STEP" & Str(i), gbstrGasAlias(3), m_Recipe.arrayRecipe(i, GB_PROCESS_GAS3), strFileName)
'        'Gas of O2
'        lngRet = WritePrivateProfileString("STEP" & Str(i), gbstrGasAlias(4), m_Recipe.arrayRecipe(i, GB_PROCESS_GAS3), strFileName)
   Next i
    'gtcRTP 1.0.0
    lngRet = WritePrivateProfileString("COFFICIENT", "IntensityWeightD1", CStr(gbsngIntensityWeight(0)), StrFileName)
    lngRet = WritePrivateProfileString("COFFICIENT", "IntensityWeightD2", CStr(gbsngIntensityWeight(1)), StrFileName)
    lngRet = WritePrivateProfileString("COFFICIENT", "IntensityWeightD3", CStr(gbsngIntensityWeight(2)), StrFileName)
    lngRet = WritePrivateProfileString("COFFICIENT", "IntensityWeightD4", CStr(gbsngIntensityWeight(3)), StrFileName)
    lngRet = WritePrivateProfileString("COFFICIENT", "IntensityWeightD5", CStr(gbsngIntensityWeight(4)), StrFileName)
    
    lngRet = WritePrivateProfileString("COFFICIENT", "IntensityWeightS1", CStr(gbsngIntensityWeightS(0)), StrFileName)
    lngRet = WritePrivateProfileString("COFFICIENT", "IntensityWeightS2", CStr(gbsngIntensityWeightS(1)), StrFileName)
    lngRet = WritePrivateProfileString("COFFICIENT", "IntensityWeightS3", CStr(gbsngIntensityWeightS(2)), StrFileName)
    lngRet = WritePrivateProfileString("COFFICIENT", "IntensityWeightS4", CStr(gbsngIntensityWeightS(3)), StrFileName)
    lngRet = WritePrivateProfileString("COFFICIENT", "IntensityWeightS5", CStr(gbsngIntensityWeightS(4)), StrFileName)
    
    
    
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "Proportional", frmRecipeEdit.txtProportional.text, StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "Proportional2", frmRecipeEdit.txtProportional2.text, StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "Integrnal", frmRecipeEdit.txtIntegrnal.text, StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "Integral2", frmRecipeEdit.txtIntegral2.text, StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "Derivational", frmRecipeEdit.txtDerivational.text, StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "Factor1", frmRecipeEdit.txtPredit.text, StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "Factor2", frmRecipeEdit.txtFeedForward.text, StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "InputDevice", CStr(frmRecipeEdit.intRecipeTempInputType), StrFileName)
    lngRet = WritePrivateProfileString("CONTROL_LOOP", "Overshoot", frmRecipeEdit.txtOvershoot.text, StrFileName)
    
   If gbProcessRecordCount = 0 Then Exit Sub
   'frmProgress.Show
   Dim iProg As Long
   iProg = gbProcessRecordCount / 100
    '0: time '1: Power '2: TC1 '3: PM '4: Gas1 '5: Gas2 '6: Gas3 '7: Gas4 '8: Vacuum
    '11:TC1 / 12:TC2 / 13:TC3 / 14:TC4 / 15:TC5 / 16:Power1 / 17:Power2 / 18:Power3 / 19:Power4 / 20:Power5
    'ColumnName=Index,Time,Intensity,TC,Pump,Gas1~4,Pressure,MTC1~5,Power1~5,O2
    
'   lngRet = WritePrivateProfileString("PROCESS_RECORD", "No", "Time,Intensity,TC1,NA,Gas1,Gas2,Gas3,Gas4,Gas5,Pressure,NA,TC2,TC3,TC4,TC5,TC6,Power1,Power2,Power3,Power4,Power5,O2-Gauge,TC7,O2,MTC8,MTC9,MTC10,MTC11,MTC12,MTC13,MTC14,MTC15,MTC16,MTC17,MTC18,MTC19,MTC20,MTC21,MTC22,MTC23", strFileName)
    lngRet = WritePrivateProfileString("PROCESS_RECORD", "No", "Time,Intensity,TC1,NA,Gas1,Gas2,Gas3,Gas4,Gas5,Pressure,Gas6,TC2,TC3,TC4,TC5,TC6,Power1,Power2,Power3,Power4,Power5,O2-Gauge,TC7,O2,MTC8,MTC9,MTC10,MTC11,MTC12,MTC13,MTC14,MTC15,MTC16,MTC17,MTC18,MTC19,MTC20,MTC21,MTC22,MTC23", StrFileName)
    
'   For i = 0 To gbProcessRecordCount - 1
'       strTemp = CStr(gbsngProcessRecorder(i, 0)) & "," & _
'                         CStr(gbsngProcessRecorder(i, 1)) & "," & _
'                         CStr(gbsngProcessRecorder(i, 2)) & "," & _
'                         CStr(gbsngProcessRecorder(i, 3)) & "," & _
'                         CStr(gbsngProcessRecorder(i, 4)) & "," & _
'                         CStr(gbsngProcessRecorder(i, 5)) & "," & _
'                         CStr(gbsngProcessRecorder(i, 6)) & "," & _
'                         CStr(gbsngProcessRecorder(i, 7)) & "," & _
'                         CStr(gbsngProcessRecorder(i, 8)) & "," & _
'                         CStr(gbsngProcessRecorder(i, 9)) & "," & _
'                         CStr(gbsngProcessRecorder(i, 10)) & "," & _
'                         CStr(gbsngProcessRecorder(i, 11)) & "," & _
'                         CStr(gbsngProcessRecorder(i, 12)) & "," & _
'                         CStr(gbsngProcessRecorder(i, 13)) & "," & _
'                         CStr(gbsngProcessRecorder(i, 14)) & "," & _
'                         CStr(gbsngProcessRecorder(i, 15)) & "," & _
'                         CStr(gbsngProcessRecorder(i, 16)) & "," & _
'                         CStr(gbsngProcessRecorder(i, 17)) & "," & _
'                         CStr(gbsngProcessRecorder(i, 18)) & "," & _
'                         CStr(gbsngProcessRecorder(i, 19)) & "," & _
'                         CStr(gbsngProcessRecorder(i, 20)) & "," & _
'                         CStr(gbsngProcessRecorder(i, 21)) & "," & _
'                         CStr(gbsngProcessRecorder(i, 22)) & "," & _
'                         CStr(gbsngProcessRecorder(i, 23))
'        If Para.UseMTC = 1 Then
'            strTemp = strTemp & "," & CStr(gbsngProcessRecorder(i, 24)) & "," & _
'                    CStr(gbsngProcessRecorder(i, 25)) & "," & _
'                    CStr(gbsngProcessRecorder(i, 26)) & "," & _
'                    CStr(gbsngProcessRecorder(i, 27)) & "," & _
'                    CStr(gbsngProcessRecorder(i, 28)) & "," & _
'                    CStr(gbsngProcessRecorder(i, 29)) & "," & _
'                    CStr(gbsngProcessRecorder(i, 30)) & "," & _
'                    CStr(gbsngProcessRecorder(i, 31))
'        End If
'        If Para.UseMTCB = 1 Then
'            strTemp = strTemp & "," & CStr(gbsngProcessRecorder(i, 32)) & "," & _
'                    CStr(gbsngProcessRecorder(i, 33)) & "," & _
'                    CStr(gbsngProcessRecorder(i, 34)) & "," & _
'                    CStr(gbsngProcessRecorder(i, 35)) & "," & _
'                    CStr(gbsngProcessRecorder(i, 36)) & "," & _
'                    CStr(gbsngProcessRecorder(i, 37)) & "," & _
'                    CStr(gbsngProcessRecorder(i, 38)) & "," & _
'                    CStr(gbsngProcessRecorder(i, 39))
'        End If
'
'
'        lngRet = WritePrivateProfileString("PROCESS_RECORD", CStr(i), strTemp, strFileName)
'
'    Next i
    'Call frmConfiguration.StopWatchDog
    If Not gbstrLogFilePath = "" Then
        If Not gbstrLogFilePath = "C:\Program Files\eRTA100" Then
            strFilePath = gbstrLogFilePath & "\Log"
            strDir = dir(strFilePath, vbDirectory)
            If strDir = "" Then MkDir strFilePath
            strFilePath = strFilePath & "\" & Year(Date)
            strDir = dir(strFilePath, vbDirectory)
            If strDir = "" Then MkDir strFilePath
            strFilePath = strFilePath & "\" & Month(Date)
            strDir = dir(strFilePath, vbDirectory)
            If strDir = "" Then MkDir strFilePath
            strFilePath = strFilePath & "\" & Day(Date)
            strDir = dir(strFilePath, vbDirectory)
            If strDir = "" Then MkDir strFilePath
            
            If gbintActiveModule_Barcode = 0 Or gbstrBN = "" Then
                strLogFileName = strFilePath & "\" & Mid(frmPlotProcess.Caption, 1, InStr(1, frmPlotProcess.Caption, ".") - 1) & "_" & Format(Time, "hhmmss") & ".log"
            Else
                strLogFileName = strFilePath & "\" & Mid(frmPlotProcess.Caption, 1, InStr(1, frmPlotProcess.Caption, ".") - 1) & "_" & Format(Time, "hhmmss") & "_" & gbstrBN & ".log"
            End If
            
            FileCopy StrFileName, strLogFileName
             
             
            
        End If
    End If
    
    If Para.UseCIM = 1 And Para.intCIMPort = 1 Then
        strFilePath = gbstrLogFilePath & "\Log"
        strDir = dir(strFilePath, vbDirectory)
        If strDir = "" Then MkDir strFilePath
        strFilePath = strFilePath & "\" & Year(Date)
        strDir = dir(strFilePath, vbDirectory)
        If strDir = "" Then MkDir strFilePath
        strFilePath = strFilePath & "\" & Month(Date)
        strDir = dir(strFilePath, vbDirectory)
        If strDir = "" Then MkDir strFilePath
        strFilePath = strFilePath & "\" & Day(Date)
        strDir = dir(strFilePath, vbDirectory)
        If strDir = "" Then MkDir strFilePath
    
        CurrProc.strLogFilePath = strFilePath
        CurrProc.strLogFileName = Mid(frmPlotProcess.Caption, 1, InStr(1, frmPlotProcess.Caption, ".") - 1) & "_" & Format(Time, "hhmmss") & ".log"
    End If
    
'    If Para.UseBarcodeServer = 1 Then
'        FileCopy strFileName, strServerFileName
'    End If
    
    frmConfiguration.StartWatchDog
    Exit Sub
ERR_PROCESSLOG_SAVE:
    frmConfiguration.StartWatchDog
    
End Sub


 Private Function ReadLogRecord(fileName As String) As String()
 
    Dim fileNumber As Integer
    Dim FilePath As String
    Dim lineContent As String
    Dim lines() As String
    Dim lineData() As String
    Dim Index As Integer
    Dim lineCount As Long
    On Error GoTo ERR_ReadLogRecord
    FilePath = fileName
    fileNumber = FreeFile
    Open FilePath For Input As fileNumber
    lineCount = 0
    ReDim lineData(1)
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, lineContent
        If lineContent <> "" And InStr(lineContent, "=") > 0 Then
          lineData = Split(lineContent, "=")
         If IsNumeric(lineData(0)) Then
          lineCount = lineCount + 1
          ReDim Preserve lines(lineCount)
          lines(lineCount - 1) = lineData(1)
'         lines.Add lineContent
         End If
        End If
    Loop
    Close fileNumber
    ReadLogRecord = lines
    Exit Function
ERR_ReadLogRecord:
  WriteLog ("讀取log數據失敗!")
    
End Function



Private Sub OpenProcessLog(StrFileName As String)
    Dim i                           As Long
    Dim j                           As Long
    Dim lngRet                As Long
    Dim iInputDevice        As Integer
    Dim StrData(GB_MAX_STEP_PROCESS, GB_MAX_STEP_PROCESS)    As String * 15
    Dim strSubData(35)    As String * 35
    Dim strProcessData   As String * 400
    Dim strTemp() As String
    Dim temp As Integer
    Dim intGetCol As Integer
    Dim LogData() As String
    
    
    On Error GoTo ERR_PROCESSLOG_OPEN
    
    Erase gbsngLogProcessRecord
    If dir(StrFileName) = "" Then GoTo ERR_PROCESSLOG_OPEN
    
    gbLogRecipe.arrayRecipe(0, 0) = 0
    For i = 1 To GB_MAX_STEP_PROCESS
        lngRet = GetPrivateProfileString("STEP " & CStr(i), "ACTION", "0", StrData(i, GB_PROCESS_ACTION), 20, StrFileName)
        lngRet = GetPrivateProfileString("STEP " & CStr(i), "TEMP", "0", StrData(i, GB_PROCESS_TEMP), 20, StrFileName)
        lngRet = GetPrivateProfileString("STEP " & CStr(i), "TIME", "0", StrData(i, GB_PROCESS_TIME), 20, StrFileName)
        
        
'        temp = gbintMaxGasEnable
'        If gbintMaxGasEnable > 2 Then
'            temp = 2
'        End If
        For j = 0 To 4
            lngRet = GetPrivateProfileString("STEP " & CStr(i), gbstrGasAlias(j), "0", StrData(i, GB_PROCESS_GAS1 + j), 20, StrFileName)
        Next j
        txtArray(GB_PROCESS_ACTION).text = StrData(i, GB_PROCESS_ACTION)
        txtArray(GB_PROCESS_TEMP) = StrData(i, GB_PROCESS_TEMP)
        txtArray(GB_PROCESS_TIME) = StrData(i, GB_PROCESS_TIME)
        For j = 0 To 4
            txtArray(GB_PROCESS_GAS1 + j) = StrData(i, GB_PROCESS_GAS1 + j)
        Next j
        
        gbLogRecipe.arrayRecipe(i, GB_PROCESS_ACTION) = txtArray(GB_PROCESS_ACTION)
        gbLogRecipe.arrayRecipe(i, GB_PROCESS_TEMP) = txtArray(GB_PROCESS_TEMP)
        gbLogRecipe.arrayRecipe(i, GB_PROCESS_TIME) = txtArray(GB_PROCESS_TIME)
        For j = 0 To 4
            gbLogRecipe.arrayRecipe(i, GB_PROCESS_GAS1 + j) = txtArray(GB_PROCESS_GAS1 + j)
        Next j
    Next i

    lngRet = GetPrivateProfileString("CONTROL_LOOP", "Proportional", "0", strSubData(0), 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "Proportional2", "0", strSubData(10), 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "Integrnal", "0", strSubData(1), 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "Derivational", "0", strSubData(2), 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "InputDevice", "0", strSubData(3), 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "Overshoot", "0", strSubData(4), 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "Factor1", "0", strSubData(5), 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "Factor2", "0", strSubData(6), 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "InputObject", "0", strSubData(7), 20, StrFileName)
    lngRet = GetPrivateProfileString("CONTROL_LOOP", "Integral2", "0", strSubData(9), 20, StrFileName)
    
    lbPIDValue(0).Caption = CStr(Val(strSubData(0))) ', "0.0000")
    lbPIDValue(2).Caption = CStr(Val(strSubData(10))) ', "0.0000")
    lbPIDValue(1).Caption = CStr(Val(strSubData(1))) ', "0.0000")
    lbPIDValue(3).Caption = CStr(Val(strSubData(9))) ', "0.0000")
'    lbPIDValue(3).Caption = CStr(Val(strSubData(5))) ', "0.0000")
'    lbPIDValue(4).Caption = CStr(Val(strSubData(6))) ', "0.0000")
'    lbPIDValue(5).Caption = CStr(Val(strSubData(9))) ', "0.0000")
    
    lngRet = GetPrivateProfileString("COMMENT", "PN", "NA", strSubData(0), 20, StrFileName)
    lbPN.Caption = strSubData(0)
    lngRet = GetPrivateProfileString("COMMENT", "BN", "NA", strSubData(0), 20, StrFileName)
    lbBN.Caption = strSubData(0)
    lngRet = GetPrivateProfileString("COMMENT", "ID1", "NA", strSubData(0), 20, StrFileName)
    lbID1.Caption = strSubData(0)
    lngRet = GetPrivateProfileString("COMMENT", "ID2", "NA", strSubData(0), 20, StrFileName)
    lbID2.Caption = strSubData(0)
    
    lngRet = GetPrivateProfileString("COMMENT", "PumpDownTime", "0", strSubData(0), 20, StrFileName)
    lbPumpDownTime.Caption = strSubData(0)
'
    
    
    LogData = ReadLogRecord(StrFileName)
    GbLogRcdCount = UBound(LogData) - 1
'    GbLogRcdCount = GetKeyCountInSection("PROCESS_RECORD", StrFileName) - 2
'    For i = 0 To 65535
   For i = 0 To GbLogRcdCount
'        lngRet = GetPrivateProfileString("PROCESS_RECORD", CStr(i), "0,0,0,0,0,0,0,0,0", strProcessData, 65535, StrFileName)
        strProcessData = LogData(i)
        strTemp = Split(strProcessData, ",")
        intGetCol = UBound(strTemp)
        If i > 0 And strTemp(0) = "0" Then Exit Sub
        gbsngLogProcessRecord(i, 0) = strTemp(0)   'Time
        gbsngLogProcessRecord(i, 1) = strTemp(1)   'Intensity
        gbsngLogProcessRecord(i, 2) = strTemp(2)   'TC
        gbsngLogProcessRecord(i, 3) = strTemp(3)   'PM
        
'        temp = gbintMaxGasEnable
'        If gbintMaxGasEnable > 2 Then
'            temp = 2
'        End If
        For j = 0 To 4
            gbsngLogProcessRecord(i, 4 + j) = strTemp(4 + j) 'Gas
        Next j
        
        gbsngLogProcessRecord(i, 9) = strTemp(9)   'Vacuum
        
        gbsngLogProcessRecord(i, 11) = strTemp(11)        'MTC1
        gbsngLogProcessRecord(i, 12) = strTemp(12)      'MTC2
        gbsngLogProcessRecord(i, 13) = strTemp(13)      'MTC3
        gbsngLogProcessRecord(i, 14) = strTemp(14)      'MTC4
        gbsngLogProcessRecord(i, 15) = strTemp(15)      'MTC5
        
        gbsngLogProcessRecord(i, 17) = strTemp(17)   'Intensity2
        gbsngLogProcessRecord(i, 18) = strTemp(18)   'Intensity3
        gbsngLogProcessRecord(i, 19) = strTemp(19)   'Intensity4
        gbsngLogProcessRecord(i, 20) = strTemp(20)   'Intensity5
'        gbsngLogProcessRecord(i, 15) = strTemp(15)   'Pow2
'        gbsngLogProcessRecord(i, 16) = strTemp(16)   'Pow3
'        gbsngLogProcessRecord(i, 17) = strTemp(17)   'Pow4
'        gbsngLogProcessRecord(i, 18) = strTemp(18)   'Pow5
        
        'gbsngLogProcessRecord(i, 19) = strTemp(19)   'Oxygen
        If intGetCol >= 21 Then
            gbsngLogProcessRecord(i, 20) = strTemp(20)   'MTC6
            gbsngLogProcessRecord(i, 21) = strTemp(21)   'MTC7
        End If
        
        If intGetCol >= 23 Then
            gbsngLogProcessRecord(i, 22) = strTemp(22)   'MTC6
            gbsngLogProcessRecord(i, 23) = strTemp(23)   'MTC7
        End If
        
'        If Para.UseMTC = 1 And intGetCol = 31 Then
        If Para.UseMTC = 1 And intGetCol >= 39 Then
             For j = 24 To 31
                gbsngLogProcessRecord(i, j) = strTemp(j)   'MTC8~15
            Next j
        End If
        If Para.UseMTCB = 1 And intGetCol >= 39 Then
             For j = 32 To 39
                gbsngLogProcessRecord(i, j) = strTemp(j)   'MTC16~23
            Next j
        End If
        DoEvents
    Next i
            
    Exit Sub
ERR_PROCESSLOG_OPEN:
    ShowMessageOK "記錄檔開啟失敗!"
End Sub

Private Sub Form_Paint()
    Dim i As Integer
    
    'picProcessLog(0).Refresh
End Sub








'============================================================================================================================

Private Sub picProcessLog_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If chkZoom.value = 1 Then
        ZoomRect.Left = X
        ZoomRect.Top = Y
        shpRect.Left = ZoomRect.Left
        shpRect.Top = ZoomRect.Top
        shpRect.Width = 1
        shpRect.Height = 1
        shpRect.Visible = True
    End If
End Sub

Private Sub picProcessLog_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim temp As Single
    
    If chkZoom.value = 1 Then
        shpRect.Visible = False
        chkZoom.value = 0
        blnZoom = True
        
        ZoomRect.Right = X
        ZoomRect.Bottom = Y
        
        If ZoomRect.Right < ZoomRect.Left Then
            temp = ZoomRect.Left
            ZoomRect.Left = ZoomRect.Right
            ZoomRect.Right = temp
        End If
        If ZoomRect.Bottom > ZoomRect.Top Then
            temp = ZoomRect.Bottom
            ZoomRect.Bottom = ZoomRect.Top
            ZoomRect.Top = temp
        End If
    
        PlotProcessLogChart
        cmdOpenLog.Enabled = True
        chkZoom.Enabled = True
        cmdZoomBack.Enabled = True
    End If
End Sub


Private Function TrimChars(ByVal str As String, ByVal charToTrim As String) As String
    Dim i As Integer
    For i = 1 To Len(str)
        If Mid(str, i, 1) <> charToTrim Then
            str = Mid(str, i)
            Exit For
        End If
    Next i
    TrimChars = str
End Function


Private Sub picProcessLog_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim CurrTimePixel As Single
    Dim CurrTempPixel As Single
    Dim i As Long
    Dim S() As String
    Dim ss As String
    Dim j As Integer
    Dim Intensity As Double
    Dim tc(30) As Double
    Dim vacuum As Double
    Dim Pressure As Double
    Dim gas(GB_GAS_MAX) As Double
    Dim Max As Integer
    Dim O2_Senser As Double
    Dim O2 As Double
    Dim Power(4) As Double
    Dim GasCount As Integer
    If chkZoom.value = 0 Then
        linMove.Visible = blnOpened
        lbCurrTime.Visible = blnOpened
'        lbData.Visible = blnOpened
        If blnOpened Then
            If X > 0 And X < lnGridW Then
                linMove.y1 = 0
                linMove.y2 = fraProcessLogChart.Height + 500
                linMove.x1 = X
                linMove.x2 = linMove.x1
                CurrTimePixel = ((lngMaxTime \ 15) * 15 + 15) / lnGridW * X
                CurrTempPixel = ((lngMaxTemp \ 100) * 100 + 300) / lnGridH * Y
                'If CurrTempPixel > lngMaxTemp Then CurrTempPixel = lngMaxTemp
                If CurrTempPixel < 0 Then CurrTempPixel = 0
                lbCurrTime.Caption = CInt(CurrTimePixel \ 60) & "'" & CInt(CurrTimePixel - (CurrTimePixel \ 60) * 60) & """" & "," & Format(CurrTempPixel, "0.0") & " C"
                lbCurrTime.Left = X
                For i = 0 To 65535
                    If TimeMap(i) >= (CurrTimePixel * 1000) Then
                        iniLog.Key = i
                        iniLog.Section = "PROCESS_RECORD"
                        S = Split(iniLog.value, ",")
                        Max = UBound(S)
                        For j = 0 To Max + 1
                            Select Case j
                                Case 1
                                    Intensity = Val(S(j)) * 10
                                Case 2
                                    tc(0) = Val(S(j))
                                Case 4
                                    gas(0) = Val(S(j))
                                Case 5
                                    gas(1) = Val(S(j))
                                Case 6
                                    gas(2) = Val(S(j))
                                Case 7
                                    gas(3) = Val(S(j))
                                Case 8
                                    gas(4) = Val(S(j))
'                                    If Max >= 21 Then vacuum = Val(S(j))
'                                    If Max >= 23 Then gas(4) = Val(S(j))
                                Case 9
'                                    If max >= 21 Then tc(1) = Val(S(j))
'                                    If max >= 23 Then vacuum = Val(S(j))
                                     Pressure = Val(S(j))
                                Case 10
                                        gas(5) = Val(S(j))
'                                    If max >= 21 Then tc(2) = Val(S(j))
                                Case 11
'                                    If max >= 21 Then tc(3) = Val(S(j))
'                                    If max >= 23 Then tc(1) = Val(S(j))
                                     tc(1) = Val(S(j))

                                Case 12
'                                    If max >= 21 Then tc(4) = Val(S(j))
'                                    If max >= 23 Then tc(2) = Val(S(j))
                                     tc(2) = Val(S(j))
                                Case 13
'                                    If max >= 21 Then tc(5) = Val(S(j))
'                                    If max >= 23 Then tc(3) = Val(S(j))
                                     tc(3) = Val(S(j))
                                Case 14
                                     tc(4) = Val(S(j))
                                Case 15
                                     tc(5) = Val(S(j))
                                Case 17
                                Power(0) = Val(S(j)) * 10
                                Case 18
                                  Power(1) = Val(S(j)) * 10
                                Case 19
                                Power(2) = Val(S(j)) * 10
                                Case 20
                                Power(3) = Val(S(j)) * 10
'                                    If max = 21 Then tc(6) = Val(S(j))
                                Case 21
                                   O2_Senser = Val(S(j))
                                Case 22
                                    tc(6) = Val(S(j))
                                Case 23
                                     O2 = Val(S(j))
                                Case 24
                                    tc(7) = 0
                                Case 25
                                    tc(8) = Val(S(j - 1))
                                Case 26
                                    tc(9) = Val(S(j - 1))
                                Case 27
                                    tc(10) = Val(S(j - 1))
                                Case 28
                                    tc(11) = Val(S(j - 1))
                                Case 29
                                    tc(12) = Val(S(j - 1))
                                Case 30
                                    tc(13) = Val(S(j - 1))
                                Case 31
                                    tc(14) = Val(S(j - 1))
                                Case 32
                                    tc(15) = Val(S(j - 1))
                                Case 33
                                    tc(16) = Val(S(j - 1))
                                Case 34
                                    tc(17) = Val(S(j - 1))
                                Case 35
                                    tc(18) = Val(S(j - 1))
                                Case 36
                                    tc(19) = Val(S(j - 1))
                                Case 37
                                    tc(20) = Val(S(j - 1))
                                Case 38
                                    tc(21) = Val(S(j - 1))
                                Case 39
                                    tc(22) = Val(S(j - 1))
                                  Case 40
                                    tc(23) = Val(S(j - 1))
                            End Select
                                               
                        Next j
                        ss = ""
                        For j = 0 To 4
                            If tc(j) >= 0 And tc(j) < 2000 Then
'                                If j = 0 Then
''                                    ss = "TC=" & Format(tc(j), "0.0")
'                                     ss = "TC1=" & Format(tc(j), "0.0")
'                                Else
                                    If chkPlotColor(j + 7).value = 1 Then ss = ss & ",TC" & CStr(j + 1) & "=" & Format(tc(j), "0.0")
                                       TempValue(j).text = Format(tc(j), "0.0")
'                                    ss = ss & ",M" & CStr(j) & "=" & Format(tc(j), "0.0")
'                                End If
                            End If
                        Next j
                        ss = ss & ",PS=" & Format(tc(5), "0.000")
                        TempValue(5).text = Format(tc(5), "0.000")
                        ss = ss & ",P=" & Format(tc(6), "0.000") + vbCrLf
                        TempValue(6).text = Format(tc(6), "0.000")
                        ss = ss & ",O2 S=" & Format(O2, "0.000")
                        TempValue(7).text = Format(O2, "0.000")
'                        ss = ss & ",O2 Sensor=" & Format(O2_Senser, "0.000") + vbCrLf
                        
                        ss = ss & ",Int=" & Format(Intensity, "0.00")
                        IntValue(0).text = Format(Intensity, "0.00")
                        IntValue(1).text = Format(Power(0), "0.00")
                        IntValue(2).text = Format(Power(1), "0.00")
                        IntValue(3).text = Format(Power(2), "0.00")
                        IntValue(4).text = Format(Power(3), "0.00")
                        ss = ss & ",Va=" & Format(Pressure, "0.000")
'                        VacValue.text = Format(Pressure, "0.000")
                        If ChkGasColor(0).value = 1 Then ss = ss & ",GN2=" & Format(gas(0), "0.0")
                        If ChkGasColor(1).value = 1 Then ss = ss & ",PN2=" & Format(gas(1), "0.0")
                        If ChkGasColor(2).value = 1 Then ss = ss & ",AR=" & Format(gas(2), "0.0")
                        If ChkGasColor(3).value = 1 Then ss = ss & ",APC=" & Format(gas(3), "0.0")
                        If ChkGasColor(4).value = 1 Then ss = ss & ",EX=" & Format(gas(4), "0.0") + vbCrLf
                        For GasCount = 0 To 6
                          If UCase(ChkGasColor(GasCount).Caption) <> "VAC" Then
                            GasValue(GasCount).text = Format(gas(GasCount), "0.00")
                          Else
                           GasValue(GasCount).text = Format(Pressure, "0.000")
                          End If
                      
                        Next GasCount
                        
'                        lbValue.Caption = ss
                        
                        If Para.UseMTC = 1 Then
'                            ss = ""
                            For j = 8 To 15
                                If tc(j) >= 0 And tc(j) < 2000 Then
'                                    If j = 0 Then
'                                        ss = "TC=" & Format(tc(j), "0.0")
'                                    Else
                                        If chkPlotColor(j + 7).value = 1 Then ss = ss & ",TC" & CStr(j + 1) & "=" & Format(tc(j), "0.0")
                                         TempValue(j).text = Format(tc(j), "0.0")
'                                        ss = ss & ",M" & CStr(j) & "=" & Format(tc(j), "0.0")
'                                    End If
                                End If
                            Next j
                     
'                            lbMTC.Caption = ss
'                            lbValue.Caption = ss
                        End If
                                ss = ss + vbCrLf
                        If Para.UseMTCB = 1 And Max >= 39 Then
'                            ss = ""
                            For j = 16 To 23
                                If tc(j) >= 0 And tc(j) < 2000 Then
'                                    If j = 0 Then
'                                        ss = "TC=" & Format(tc(j), "0.0")
'                                    Else
                                        If chkPlotColor(j + 7).value = 1 Then ss = ss & ",TC" & CStr(j + 1) & "=" & Format(tc(j), "0.0")
                                           TempValue(j).text = Format(tc(j), "0.0")
'                                        ss = ss & ",M" & CStr(j) & "=" & Format(tc(j), "0.0")
'                                    End If
                                End If
                            Next j
                              
'                            lbMTCB.Caption = ss
                        End If
                        lbData.Caption = TrimChars(ss, ",")
                        Exit For
                    End If
                Next i
                
            End If
        End If
    Else
        If Button = 1 Then
            'CurrRatio = CurrRatio + 1
            'Label1.Caption = CStr(ZoomRect.Left) & "," & CStr(ZoomRect.Top) & "---" & CStr(X) & "," & CStr(Y)
            
            shpRect.Width = Abs(X - ZoomRect.Left)
            shpRect.Height = Abs(Y - ZoomRect.Top)
            If (X - ZoomRect.Left) < 0 Then
                shpRect.Left = ZoomRect.Left - shpRect.Width
            End If
            If (Y - ZoomRect.Top) > 0 Then
                shpRect.Top = ZoomRect.Top + shpRect.Height
            End If
            'picProcessLog(0).Line (ZoomRect.Left, ZoomRect.Top)-(X, Y), RGB(255, 0, 0), B
        End If
    End If
    
    'Label1.Caption = CStr(X) & "," & CStr(Y)
    Erase S
End Sub



