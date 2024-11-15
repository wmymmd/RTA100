VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPlotProcess 
   Caption         =   "製程曲線"
   ClientHeight    =   11640
   ClientLeft      =   -4935
   ClientTop       =   -600
   ClientWidth     =   18825
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
   ScaleHeight     =   11640
   ScaleMode       =   0  'User
   ScaleWidth      =   19922.03
   WindowState     =   2  'Maximized
   Begin VB.Frame fraMTCB 
      Caption         =   "MTC2"
      Height          =   3495
      Left            =   16200
      TabIndex        =   201
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
      Begin VB.CheckBox chkTC 
         Caption         =   "Check1"
         Height          =   270
         Index           =   16
         Left            =   1320
         TabIndex        =   209
         Top             =   360
         Width           =   255
      End
      Begin VB.CheckBox chkTC 
         Caption         =   "Check1"
         Height          =   270
         Index           =   17
         Left            =   1320
         TabIndex        =   208
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox chkTC 
         Caption         =   "Check1"
         Height          =   270
         Index           =   18
         Left            =   1320
         TabIndex        =   207
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox chkTC 
         Caption         =   "Check1"
         Height          =   270
         Index           =   19
         Left            =   1320
         TabIndex        =   206
         Top             =   1440
         Width           =   255
      End
      Begin VB.CheckBox chkTC 
         Caption         =   "Check1"
         Height          =   270
         Index           =   20
         Left            =   1320
         TabIndex        =   205
         Top             =   1800
         Width           =   255
      End
      Begin VB.CheckBox chkTC 
         Caption         =   "Check1"
         Height          =   270
         Index           =   21
         Left            =   1320
         TabIndex        =   204
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox chkTC 
         Caption         =   "Check1"
         Height          =   270
         Index           =   22
         Left            =   1320
         TabIndex        =   203
         Top             =   2520
         Width           =   255
      End
      Begin VB.CheckBox chkTC 
         Caption         =   "Check1"
         Height          =   270
         Index           =   23
         Left            =   1320
         TabIndex        =   202
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label lbTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   16
         Left            =   480
         TabIndex        =   225
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "17"
         Height          =   375
         Index           =   15
         Left            =   120
         TabIndex        =   224
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "18"
         Height          =   375
         Index           =   14
         Left            =   120
         TabIndex        =   223
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lbTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   17
         Left            =   480
         TabIndex        =   222
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lbTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   18
         Left            =   480
         TabIndex        =   221
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lbTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   19
         Left            =   480
         TabIndex        =   220
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lbTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000040C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   20
         Left            =   480
         TabIndex        =   219
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lbTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FF00FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   21
         Left            =   480
         TabIndex        =   218
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lbTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   22
         Left            =   480
         TabIndex        =   217
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lbTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   23
         Left            =   480
         TabIndex        =   216
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "19"
         Height          =   375
         Index           =   13
         Left            =   120
         TabIndex        =   215
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "20"
         Height          =   375
         Index           =   12
         Left            =   120
         TabIndex        =   214
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "21"
         Height          =   375
         Index           =   11
         Left            =   120
         TabIndex        =   213
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "22"
         Height          =   375
         Index           =   10
         Left            =   120
         TabIndex        =   212
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "23"
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   211
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "24"
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   210
         Top             =   2880
         Width           =   255
      End
   End
   Begin VB.Frame fraBankH 
      Caption         =   "Lamp Current"
      Height          =   2055
      Left            =   120
      TabIndex        =   140
      Top             =   9480
      Width           =   17175
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgCTConfigProcess 
         Height          =   1695
         Left            =   0
         TabIndex        =   255
         Top             =   240
         Visible         =   0   'False
         Width           =   17055
         _ExtentX        =   30083
         _ExtentY        =   2990
         _Version        =   393216
         BorderStyle     =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   59
         Left            =   10920
         TabIndex        =   200
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   58
         Left            =   10440
         TabIndex        =   199
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   57
         Left            =   9960
         TabIndex        =   198
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   56
         Left            =   9480
         TabIndex        =   197
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   55
         Left            =   9000
         TabIndex        =   196
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   54
         Left            =   8400
         TabIndex        =   195
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   53
         Left            =   7920
         TabIndex        =   194
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   52
         Left            =   7440
         TabIndex        =   193
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   51
         Left            =   6960
         TabIndex        =   192
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   50
         Left            =   7920
         TabIndex        =   191
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   49
         Left            =   7440
         TabIndex        =   190
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   48
         Left            =   6960
         TabIndex        =   189
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   47
         Left            =   6120
         TabIndex        =   188
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   46
         Left            =   5640
         TabIndex        =   187
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   45
         Left            =   5160
         TabIndex        =   186
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   44
         Left            =   4440
         TabIndex        =   185
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   43
         Left            =   3960
         TabIndex        =   184
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   42
         Left            =   3480
         TabIndex        =   183
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   41
         Left            =   2760
         TabIndex        =   182
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   40
         Left            =   2280
         TabIndex        =   181
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   39
         Left            =   1800
         TabIndex        =   180
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   38
         Left            =   1080
         TabIndex        =   179
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   37
         Left            =   600
         TabIndex        =   178
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   36
         Left            =   120
         TabIndex        =   177
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   24
         Left            =   3480
         TabIndex        =   176
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   23
         Left            =   2760
         TabIndex        =   175
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   22
         Left            =   2280
         TabIndex        =   174
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   21
         Left            =   1800
         TabIndex        =   173
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   28
         Left            =   5640
         TabIndex        =   172
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   27
         Left            =   5160
         TabIndex        =   171
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   26
         Left            =   4440
         TabIndex        =   170
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   25
         Left            =   3960
         TabIndex        =   169
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   32
         Left            =   7920
         TabIndex        =   168
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   31
         Left            =   7440
         TabIndex        =   167
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   30
         Left            =   6960
         TabIndex        =   166
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   29
         Left            =   6120
         TabIndex        =   165
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   35
         Left            =   9720
         TabIndex        =   164
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   34
         Left            =   9240
         TabIndex        =   163
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   33
         Left            =   8760
         TabIndex        =   162
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   20
         Left            =   1080
         TabIndex        =   161
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   19
         Left            =   600
         TabIndex        =   160
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   18
         Left            =   120
         TabIndex        =   159
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   15
         Left            =   8760
         TabIndex        =   158
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   16
         Left            =   9240
         TabIndex        =   157
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   17
         Left            =   9720
         TabIndex        =   156
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   12
         Left            =   6960
         TabIndex        =   155
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   13
         Left            =   7440
         TabIndex        =   154
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   14
         Left            =   7920
         TabIndex        =   153
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   10
         Left            =   5640
         TabIndex        =   152
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   11
         Left            =   6120
         TabIndex        =   151
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   8
         Left            =   4440
         TabIndex        =   150
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   9
         Left            =   5160
         TabIndex        =   149
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   6
         Left            =   3480
         TabIndex        =   148
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   7
         Left            =   3960
         TabIndex        =   147
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   3
         Left            =   1800
         TabIndex        =   146
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   4
         Left            =   2280
         TabIndex        =   145
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   5
         Left            =   2760
         TabIndex        =   144
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   2
         Left            =   1080
         TabIndex        =   143
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   142
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbCT 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   141
         Top             =   240
         Width           =   495
      End
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
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   12000
      TabIndex        =   116
      Top             =   8760
      Width           =   3015
      Begin VB.Label lbName 
         Alignment       =   1  'Right Justify
         Caption         =   "氧氣:"
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   119
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lbName 
         Caption         =   "ppm"
         Height          =   375
         Index           =   3
         Left            =   2160
         TabIndex        =   118
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lbOxygen 
         Alignment       =   2  'Center
         Caption         =   "999"
         Height          =   255
         Left            =   1080
         TabIndex        =   117
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraMTC 
      Caption         =   "MTC1"
      Height          =   3495
      Left            =   14520
      TabIndex        =   95
      Top             =   0
      Width           =   1695
      Begin VB.CheckBox chkTC 
         Caption         =   "Check1"
         Height          =   270
         Index           =   15
         Left            =   1320
         TabIndex        =   139
         Top             =   2880
         Width           =   255
      End
      Begin VB.CheckBox chkTC 
         Caption         =   "Check1"
         Height          =   270
         Index           =   14
         Left            =   1320
         TabIndex        =   138
         Top             =   2520
         Width           =   255
      End
      Begin VB.CheckBox chkTC 
         Caption         =   "Check1"
         Height          =   270
         Index           =   13
         Left            =   1320
         TabIndex        =   137
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox chkTC 
         Caption         =   "Check1"
         Height          =   270
         Index           =   12
         Left            =   1320
         TabIndex        =   136
         Top             =   1800
         Width           =   255
      End
      Begin VB.CheckBox chkTC 
         Caption         =   "Check1"
         Height          =   270
         Index           =   11
         Left            =   1320
         TabIndex        =   135
         Top             =   1440
         Width           =   255
      End
      Begin VB.CheckBox chkTC 
         Caption         =   "Check1"
         Height          =   270
         Index           =   10
         Left            =   1320
         TabIndex        =   134
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox chkTC 
         Caption         =   "Check1"
         Height          =   270
         Index           =   9
         Left            =   1320
         TabIndex        =   133
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox chkTC 
         Caption         =   "Check1"
         Height          =   270
         Index           =   8
         Left            =   1320
         TabIndex        =   132
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "16"
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   111
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "15"
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   110
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "14"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   109
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "13"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   108
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "12"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   107
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "11"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   106
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lbTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   15
         Left            =   480
         TabIndex        =   105
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label lbTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   14
         Left            =   480
         TabIndex        =   104
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lbTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FF00FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   13
         Left            =   480
         TabIndex        =   103
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lbTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000040C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   12
         Left            =   480
         TabIndex        =   102
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lbTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   11
         Left            =   480
         TabIndex        =   101
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lbTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   10
         Left            =   480
         TabIndex        =   100
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lbTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   9
         Left            =   480
         TabIndex        =   99
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "10"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   98
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "9"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   97
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lbTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   8
         Left            =   480
         TabIndex        =   96
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame fraBarcode 
      Caption         =   "Barcode"
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
      Height          =   1815
      Left            =   15120
      TabIndex        =   76
      Top             =   3960
      Visible         =   0   'False
      Width           =   3015
      Begin VB.TextBox txtID2 
         Height          =   390
         Left            =   600
         TabIndex        =   80
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txtID1 
         Height          =   390
         Left            =   600
         TabIndex        =   79
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtBN 
         Height          =   390
         Left            =   600
         TabIndex        =   78
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtPN 
         Height          =   390
         Left            =   600
         TabIndex        =   77
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "ID"
         Height          =   270
         Index           =   15
         Left            =   120
         TabIndex        =   84
         Top             =   1320
         Width           =   225
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "EN"
         Height          =   270
         Index           =   6
         Left            =   120
         TabIndex        =   83
         Top             =   960
         Width           =   330
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "BN"
         Height          =   270
         Index           =   5
         Left            =   120
         TabIndex        =   82
         Top             =   600
         Width           =   330
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "PN"
         Height          =   270
         Index           =   4
         Left            =   120
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   81
         Top             =   240
         Width           =   330
      End
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   5175
      Left            =   11880
      TabIndex        =   50
      Top             =   3600
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "輸出"
      TabPicture(0)   =   "frmPlotProcess.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lbNameMFC(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbGasUnit(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbGasUnit(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbGasUnit(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbNameMFC(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbNameMFC(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbName(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbGasUnit(3)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbNameMFC(4)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lbNameMFC(3)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lbGasUnit(4)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lbNameMFC(5)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lbGasUnit(5)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtGas(2)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtGas(1)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtGas(0)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtIntensity"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtGas(4)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtGas(3)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "fraVacFunc"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtGas(5)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "PID"
      TabPicture(1)   =   "frmPlotProcess.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbName(19)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lbPIDValue(6)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lbName(14)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lbPIDValue(5)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lbPIDValue(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lbPIDValue(0)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lbName(11)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lbName(10)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lbScanCount"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lbOutput"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label5"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Loop"
      TabPicture(2)   =   "frmPlotProcess.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lbIntensityLabel(0)"
      Tab(2).Control(1)=   "lbIntensityLoop(0)"
      Tab(2).Control(2)=   "lbIntensityLabel(1)"
      Tab(2).Control(3)=   "lbIntensityLoop(1)"
      Tab(2).Control(4)=   "lbIntensityLabel(2)"
      Tab(2).Control(5)=   "lbIntensityLoop(2)"
      Tab(2).Control(6)=   "lbIntensityLabel(3)"
      Tab(2).Control(7)=   "lbIntensityLoop(3)"
      Tab(2).Control(8)=   "lbIntensityLabel(4)"
      Tab(2).Control(9)=   "lbIntensityLoop(4)"
      Tab(2).Control(10)=   "lbIntensityLoop(5)"
      Tab(2).Control(11)=   "lbIntensityLabel(5)"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "TCM"
      TabPicture(3)   =   "frmPlotProcess.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lbIntensityAz1(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lbIntensityAz(3)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lbIntensityAz1(2)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "lbIntensityAz(2)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "lbIntensityAz1(1)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "lbIntensityAz(1)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "lbIntensityAz1(0)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "lbIntensityAz(0)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "lbIntensityAz(4)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "lbIntensityAz1(4)"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "lbIntensityAz(5)"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "lbIntensityAz1(5)"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "lbIntensityAz(6)"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "lbIntensityAz1(6)"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "lbIntensityAz(7)"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "lbIntensityAz1(7)"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "chkTest"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "chkAT"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "ChkPower(0)"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "ChkPower(1)"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "ChkPower(2)"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).Control(21)=   "ChkPower(3)"
      Tab(3).Control(21).Enabled=   0   'False
      Tab(3).Control(22)=   "ChkPower(4)"
      Tab(3).Control(22).Enabled=   0   'False
      Tab(3).Control(23)=   "ChkPower(5)"
      Tab(3).Control(23).Enabled=   0   'False
      Tab(3).Control(24)=   "ChkPower(6)"
      Tab(3).Control(24).Enabled=   0   'False
      Tab(3).Control(25)=   "ChkPower(7)"
      Tab(3).Control(25).Enabled=   0   'False
      Tab(3).ControlCount=   26
      Begin VB.TextBox txtGas 
         Alignment       =   2  'Center
         Height          =   390
         Index           =   5
         Left            =   -73800
         TabIndex        =   253
         Text            =   "0.0"
         Top             =   3360
         Width           =   855
      End
      Begin VB.CheckBox ChkPower 
         Caption         =   "Check1"
         Height          =   270
         Index           =   7
         Left            =   480
         TabIndex        =   251
         Top             =   3960
         Width           =   255
      End
      Begin VB.CheckBox ChkPower 
         Caption         =   "Check1"
         Height          =   270
         Index           =   6
         Left            =   480
         TabIndex        =   250
         Top             =   3490
         Width           =   255
      End
      Begin VB.CheckBox ChkPower 
         Caption         =   "Check1"
         Height          =   270
         Index           =   5
         Left            =   480
         TabIndex        =   249
         Top             =   3000
         Width           =   255
      End
      Begin VB.CheckBox ChkPower 
         Caption         =   "Check1"
         Height          =   270
         Index           =   4
         Left            =   480
         TabIndex        =   248
         Top             =   2530
         Width           =   255
      End
      Begin VB.CheckBox ChkPower 
         Caption         =   "Check1"
         Height          =   270
         Index           =   3
         Left            =   480
         TabIndex        =   247
         Top             =   2050
         Width           =   255
      End
      Begin VB.CheckBox ChkPower 
         Caption         =   "Check1"
         Height          =   270
         Index           =   2
         Left            =   480
         TabIndex        =   246
         Top             =   1570
         Width           =   255
      End
      Begin VB.CheckBox ChkPower 
         Caption         =   "Check1"
         Height          =   270
         Index           =   1
         Left            =   480
         TabIndex        =   245
         Top             =   1090
         Width           =   255
      End
      Begin VB.CheckBox ChkPower 
         Caption         =   "Check1"
         Height          =   270
         Index           =   0
         Left            =   480
         TabIndex        =   244
         Top             =   610
         Width           =   255
      End
      Begin VB.CheckBox chkAT 
         Caption         =   "Auto Tuning"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   243
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CheckBox chkTest 
         Caption         =   "Test"
         Height          =   375
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   242
         Top             =   4440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame fraVacFunc 
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1335
         Left            =   -74880
         TabIndex        =   85
         Top             =   3720
         Width           =   2775
         Begin VB.CheckBox chkPurge 
            Caption         =   "破真空"
            Height          =   495
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   87
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox chkPumping 
            Caption         =   "開啟泵"
            Height          =   495
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   86
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lbVacuum 
            Alignment       =   2  'Center
            Caption         =   "760"
            Height          =   255
            Left            =   720
            TabIndex        =   90
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lbName 
            Caption         =   "Torr"
            Height          =   375
            Index           =   2
            Left            =   2160
            TabIndex        =   89
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lbName 
            Alignment       =   1  'Right Justify
            Caption         =   "壓力:"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   88
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.TextBox txtGas 
         Alignment       =   2  'Center
         Height          =   390
         Index           =   3
         Left            =   -73800
         TabIndex        =   72
         Text            =   "0.0"
         Top             =   2460
         Width           =   855
      End
      Begin VB.TextBox txtGas 
         Alignment       =   2  'Center
         Height          =   390
         Index           =   4
         Left            =   -73800
         TabIndex        =   71
         Text            =   "0.0"
         Top             =   2940
         Width           =   855
      End
      Begin VB.TextBox txtIntensity 
         Alignment       =   2  'Center
         Height          =   390
         Left            =   -73800
         TabIndex        =   62
         Text            =   "0.00"
         Top             =   540
         Width           =   855
      End
      Begin VB.TextBox txtGas 
         Alignment       =   2  'Center
         Height          =   390
         Index           =   0
         Left            =   -73800
         TabIndex        =   61
         Text            =   "0.0"
         Top             =   1020
         Width           =   855
      End
      Begin VB.TextBox txtGas 
         Alignment       =   2  'Center
         Height          =   390
         Index           =   1
         Left            =   -73800
         TabIndex        =   60
         Text            =   "0.0"
         Top             =   1500
         Width           =   855
      End
      Begin VB.TextBox txtGas 
         Alignment       =   2  'Center
         Height          =   390
         Index           =   2
         Left            =   -73800
         TabIndex        =   59
         Text            =   "0.0"
         Top             =   1980
         Width           =   855
      End
      Begin VB.Label lbGasUnit 
         Caption         =   "SLPM"
         Height          =   375
         Index           =   5
         Left            =   -72840
         TabIndex        =   254
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label lbNameMFC 
         Alignment       =   1  'Right Justify
         Caption         =   "NA"
         Height          =   255
         Index           =   5
         Left            =   -74760
         TabIndex        =   252
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label lbIntensityAz1 
         Alignment       =   2  'Center
         Caption         =   "0"
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   7
         Left            =   1800
         TabIndex        =   241
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label lbIntensityAz 
         AutoSize        =   -1  'True
         Caption         =   "Intensity 8:"
         Height          =   270
         Index           =   7
         Left            =   720
         TabIndex        =   240
         Top             =   3960
         Width           =   1080
      End
      Begin VB.Label lbIntensityAz1 
         Alignment       =   2  'Center
         Caption         =   "0"
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   6
         Left            =   1800
         TabIndex        =   239
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Label lbIntensityAz 
         AutoSize        =   -1  'True
         Caption         =   "Intensity 7:"
         Height          =   270
         Index           =   6
         Left            =   720
         TabIndex        =   238
         Top             =   3480
         Width           =   1080
      End
      Begin VB.Label lbIntensityAz1 
         Alignment       =   2  'Center
         Caption         =   "0"
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   5
         Left            =   1800
         TabIndex        =   237
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label lbIntensityAz 
         AutoSize        =   -1  'True
         Caption         =   "Intensity 6:"
         Height          =   270
         Index           =   5
         Left            =   720
         TabIndex        =   236
         Top             =   3000
         Width           =   1080
      End
      Begin VB.Label lbIntensityAz1 
         Alignment       =   2  'Center
         Caption         =   "0"
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   4
         Left            =   1800
         TabIndex        =   235
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lbIntensityAz 
         AutoSize        =   -1  'True
         Caption         =   "Intensity 5:"
         Height          =   270
         Index           =   4
         Left            =   720
         TabIndex        =   234
         Top             =   2520
         Width           =   1080
      End
      Begin VB.Label lbIntensityAz 
         AutoSize        =   -1  'True
         Caption         =   "Intensity 1:"
         Height          =   270
         Index           =   0
         Left            =   720
         TabIndex        =   233
         Top             =   600
         Width           =   1080
      End
      Begin VB.Label lbIntensityAz1 
         Alignment       =   2  'Center
         Caption         =   "0"
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   232
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lbIntensityAz 
         AutoSize        =   -1  'True
         Caption         =   "Intensity 2:"
         Height          =   270
         Index           =   1
         Left            =   720
         TabIndex        =   231
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label lbIntensityAz1 
         Alignment       =   2  'Center
         Caption         =   "0"
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   230
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lbIntensityAz 
         AutoSize        =   -1  'True
         Caption         =   "Intensity 3:"
         Height          =   270
         Index           =   2
         Left            =   720
         TabIndex        =   229
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label lbIntensityAz1 
         Alignment       =   2  'Center
         Caption         =   "0"
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   2
         Left            =   1800
         TabIndex        =   228
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lbIntensityAz 
         AutoSize        =   -1  'True
         Caption         =   "Intensity 4:"
         Height          =   270
         Index           =   3
         Left            =   720
         TabIndex        =   227
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label lbIntensityAz1 
         Alignment       =   2  'Center
         Caption         =   "0"
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   3
         Left            =   1800
         TabIndex        =   226
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lbIntensityLabel 
         AutoSize        =   -1  'True
         Caption         =   "Intensity 6:"
         Height          =   270
         Index           =   5
         Left            =   -74880
         TabIndex        =   131
         Top             =   3300
         Width           =   1080
      End
      Begin VB.Label lbIntensityLoop 
         Alignment       =   2  'Center
         Caption         =   "0"
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   5
         Left            =   -73680
         TabIndex        =   130
         Top             =   3300
         Width           =   1335
      End
      Begin VB.Label lbIntensityLoop 
         Alignment       =   2  'Center
         Caption         =   "0"
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   4
         Left            =   -73680
         TabIndex        =   129
         Top             =   2820
         Width           =   1335
      End
      Begin VB.Label lbIntensityLabel 
         AutoSize        =   -1  'True
         Caption         =   "Intensity 5:"
         Height          =   270
         Index           =   4
         Left            =   -74880
         TabIndex        =   128
         Top             =   2820
         Width           =   1080
      End
      Begin VB.Label lbIntensityLoop 
         Alignment       =   2  'Center
         Caption         =   "0"
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   3
         Left            =   -73680
         TabIndex        =   127
         Top             =   2340
         Width           =   1335
      End
      Begin VB.Label lbIntensityLabel 
         AutoSize        =   -1  'True
         Caption         =   "Intensity 4:"
         Height          =   270
         Index           =   3
         Left            =   -74880
         TabIndex        =   126
         Top             =   2340
         Width           =   1080
      End
      Begin VB.Label lbIntensityLoop 
         Alignment       =   2  'Center
         Caption         =   "0"
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   2
         Left            =   -73680
         TabIndex        =   125
         Top             =   1860
         Width           =   1335
      End
      Begin VB.Label lbIntensityLabel 
         AutoSize        =   -1  'True
         Caption         =   "Intensity 3:"
         Height          =   270
         Index           =   2
         Left            =   -74880
         TabIndex        =   124
         Top             =   1860
         Width           =   1080
      End
      Begin VB.Label lbIntensityLoop 
         Alignment       =   2  'Center
         Caption         =   "0"
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   1
         Left            =   -73680
         TabIndex        =   123
         Top             =   1380
         Width           =   1335
      End
      Begin VB.Label lbIntensityLabel 
         AutoSize        =   -1  'True
         Caption         =   "Intensity 2:"
         Height          =   270
         Index           =   1
         Left            =   -74880
         TabIndex        =   122
         Top             =   1380
         Width           =   1080
      End
      Begin VB.Label lbIntensityLoop 
         Alignment       =   2  'Center
         Caption         =   "0"
         ForeColor       =   &H8000000D&
         Height          =   375
         Index           =   0
         Left            =   -73680
         TabIndex        =   121
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label lbIntensityLabel 
         AutoSize        =   -1  'True
         Caption         =   "Intensity 1:"
         Height          =   270
         Index           =   0
         Left            =   -74880
         TabIndex        =   120
         Top             =   900
         Width           =   1080
      End
      Begin VB.Label Label5 
         Caption         =   "Output"
         Height          =   255
         Left            =   -74760
         TabIndex        =   115
         Top             =   3780
         Width           =   1095
      End
      Begin VB.Label lbOutput 
         Alignment       =   2  'Center
         Caption         =   "0"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -73680
         TabIndex        =   114
         Top             =   3780
         Width           =   1095
      End
      Begin VB.Label lbScanCount 
         Alignment       =   2  'Center
         Caption         =   "0"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -73680
         TabIndex        =   113
         Top             =   4260
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "ScanLoop"
         Height          =   255
         Left            =   -74760
         TabIndex        =   112
         Top             =   4260
         Width           =   1095
      End
      Begin VB.Label lbGasUnit 
         Caption         =   "SLPM"
         Height          =   375
         Index           =   4
         Left            =   -72840
         TabIndex        =   91
         Top             =   2940
         Width           =   735
      End
      Begin VB.Label lbNameMFC 
         Alignment       =   1  'Right Justify
         Caption         =   "NA"
         Height          =   255
         Index           =   3
         Left            =   -74760
         TabIndex        =   75
         Top             =   2460
         Width           =   855
      End
      Begin VB.Label lbNameMFC 
         Alignment       =   1  'Right Justify
         Caption         =   "NA"
         Height          =   255
         Index           =   4
         Left            =   -74760
         TabIndex        =   74
         Top             =   2940
         Width           =   855
      End
      Begin VB.Label lbGasUnit 
         Caption         =   "SLPM"
         Height          =   375
         Index           =   3
         Left            =   -72840
         TabIndex        =   73
         Top             =   2460
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "%"
         Height          =   255
         Left            =   -72600
         TabIndex        =   70
         Top             =   540
         Width           =   255
      End
      Begin VB.Label lbName 
         Caption         =   "Intensity"
         Height          =   375
         Index           =   0
         Left            =   -74760
         TabIndex        =   69
         Top             =   540
         Width           =   975
      End
      Begin VB.Label lbNameMFC 
         Alignment       =   1  'Right Justify
         Caption         =   "NA"
         Height          =   375
         Index           =   0
         Left            =   -74760
         TabIndex        =   68
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label lbNameMFC 
         Alignment       =   1  'Right Justify
         Caption         =   "NA"
         Height          =   375
         Index           =   1
         Left            =   -74760
         TabIndex        =   67
         Top             =   1500
         Width           =   855
      End
      Begin VB.Label lbGasUnit 
         Caption         =   "SLPM"
         Height          =   375
         Index           =   0
         Left            =   -72840
         TabIndex        =   66
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label lbGasUnit 
         Caption         =   "SLPM"
         Height          =   375
         Index           =   1
         Left            =   -72840
         TabIndex        =   65
         Top             =   1500
         Width           =   735
      End
      Begin VB.Label lbGasUnit 
         Caption         =   "SLPM"
         Height          =   375
         Index           =   2
         Left            =   -72840
         TabIndex        =   64
         Top             =   1980
         Width           =   735
      End
      Begin VB.Label lbNameMFC 
         Alignment       =   1  'Right Justify
         Caption         =   "NA"
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   63
         Top             =   1980
         Width           =   855
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "Integral"
         Height          =   270
         Index           =   10
         Left            =   -74880
         TabIndex        =   58
         Top             =   1620
         Width           =   750
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "Proportional"
         Height          =   270
         Index           =   11
         Left            =   -74880
         TabIndex        =   57
         Top             =   900
         Width           =   1275
      End
      Begin VB.Label lbPIDValue 
         Caption         =   "0"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   0
         Left            =   -73320
         TabIndex        =   56
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label lbPIDValue 
         Caption         =   "0"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   1
         Left            =   -73320
         TabIndex        =   55
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Label lbPIDValue 
         Caption         =   "0"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   5
         Left            =   -73320
         TabIndex        =   54
         Top             =   1980
         Width           =   1215
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "Integral2"
         Height          =   270
         Index           =   14
         Left            =   -74880
         TabIndex        =   53
         Top             =   1980
         Width           =   885
      End
      Begin VB.Label lbPIDValue 
         Caption         =   "0"
         ForeColor       =   &H80000002&
         Height          =   255
         Index           =   6
         Left            =   -73320
         TabIndex        =   52
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label lbName 
         AutoSize        =   -1  'True
         Caption         =   "Proportional2"
         Height          =   270
         Index           =   19
         Left            =   -74880
         TabIndex        =   51
         Top             =   1260
         Width           =   1410
      End
   End
   Begin VB.Timer tmrAlarmFlash 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11280
      Top             =   8040
   End
   Begin VB.Timer tmrIdleCheck 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   10800
      Top             =   8040
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   720
      TabIndex        =   7
      Top             =   480
      Width           =   5655
      Begin VB.CheckBox chkShowLine 
         Caption         =   "Check1"
         Height          =   270
         Index           =   4
         Left            =   4560
         TabIndex        =   32
         Top             =   0
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox chkShowLine 
         Caption         =   "Check1"
         Height          =   270
         Index           =   3
         Left            =   3480
         TabIndex        =   31
         Top             =   0
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox chkShowLine 
         Caption         =   "Check1"
         Height          =   270
         Index           =   2
         Left            =   2400
         TabIndex        =   30
         Top             =   0
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox chkShowLine 
         Caption         =   "Check1"
         Height          =   270
         Index           =   1
         Left            =   1320
         TabIndex        =   29
         Top             =   0
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox chkShowLine 
         Caption         =   "Check1"
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H000080FF&
         Caption         =   "Vacuum"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   0
         Value           =   2  'Grayed
         Width           =   735
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H000080FF&
         Caption         =   "NoUse"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         Value           =   2  'Grayed
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H00FF8080&
         Caption         =   "O2"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   360
         Value           =   2  'Grayed
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H000000FF&
         Caption         =   "TC"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Value           =   2  'Grayed
         Width           =   735
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H0000FFFF&
         Caption         =   "Intensity"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   0
         Value           =   2  'Grayed
         Width           =   975
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H00FF00FF&
         Caption         =   "N2"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Value           =   2  'Grayed
         Width           =   735
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H0000C000&
         Caption         =   "Ar"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Value           =   2  'Grayed
         Width           =   735
      End
      Begin VB.CheckBox chkPlotColor 
         BackColor       =   &H00C0C000&
         Caption         =   "O2"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   480
         Value           =   2  'Grayed
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "溫度"
      Height          =   3495
      Left            =   11880
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.CheckBox chkTC 
         Caption         =   "Check1"
         Height          =   270
         Index           =   7
         Left            =   2400
         TabIndex        =   92
         Top             =   2880
         Width           =   255
      End
      Begin VB.CheckBox chkTC 
         Caption         =   "Check1"
         Height          =   270
         Index           =   6
         Left            =   2400
         TabIndex        =   49
         Top             =   2520
         Width           =   255
      End
      Begin VB.CheckBox chkTC 
         Caption         =   "Check1"
         Height          =   270
         Index           =   5
         Left            =   2400
         TabIndex        =   38
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox chkTC 
         Caption         =   "Check1"
         Height          =   270
         Index           =   4
         Left            =   2400
         TabIndex        =   37
         Top             =   1800
         Width           =   255
      End
      Begin VB.CheckBox chkTC 
         Caption         =   "Check1"
         Height          =   270
         Index           =   3
         Left            =   2400
         TabIndex        =   36
         Top             =   1440
         Width           =   255
      End
      Begin VB.CheckBox chkTC 
         Caption         =   "Check1"
         Height          =   270
         Index           =   2
         Left            =   2400
         TabIndex        =   35
         Top             =   1080
         Width           =   255
      End
      Begin VB.CheckBox chkTC 
         Caption         =   "Check1"
         Height          =   270
         Index           =   1
         Left            =   2400
         TabIndex        =   34
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox chkTC 
         Caption         =   "Check1"
         Height          =   270
         Index           =   0
         Left            =   2400
         TabIndex        =   33
         Top             =   360
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lbTCA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TC8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   94
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lbTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   7
         Left            =   1320
         TabIndex        =   93
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label lbTCA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TC2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   48
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lbTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   6
         Left            =   1320
         TabIndex        =   26
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lbTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FF00FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   5
         Left            =   1320
         TabIndex        =   25
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lbTCA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TC7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   24
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lbTCA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TC6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   20
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lbTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000040C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   4
         Left            =   1320
         TabIndex        =   19
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lbTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   18
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lbTCA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TC3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lbTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   2
         Left            =   1320
         TabIndex        =   16
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lbTCA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TC4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lbTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   3
         Left            =   1320
         TabIndex        =   14
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lbTCA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TC5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lbTC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   0
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lbTCA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TC1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fraProcessChart 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9405
      Left            =   240
      TabIndex        =   3
      Top             =   0
      Width           =   11535
      Begin VB.PictureBox picProcess 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9000
         Index           =   0
         Left            =   120
         ScaleHeight     =   596
         ScaleMode       =   0  'User
         ScaleWidth      =   729
         TabIndex        =   4
         Top             =   240
         Width           =   11000
         Begin VB.Frame fraAlarm 
            BackColor       =   &H000000FF&
            Caption         =   "錯誤"
            ForeColor       =   &H0000FF00&
            Height          =   2415
            Left            =   4560
            TabIndex        =   39
            Top             =   960
            Visible         =   0   'False
            Width           =   5655
            Begin VB.Label lbAlarmTime 
               BackColor       =   &H000000FF&
               Caption         =   "2016/05/01 12:00:00"
               ForeColor       =   &H0000FF00&
               Height          =   375
               Left            =   3240
               TabIndex        =   47
               Top             =   480
               Width           =   2295
            End
            Begin VB.Label lbAlarm 
               BackColor       =   &H000000FF&
               Caption         =   "錯誤時間:"
               ForeColor       =   &H0000FF00&
               Height          =   375
               Index           =   3
               Left            =   2160
               TabIndex        =   46
               Top             =   480
               Width           =   1095
            End
            Begin VB.Label lbAlarmDo 
               BackColor       =   &H000000FF&
               Caption         =   "0000"
               ForeColor       =   &H0000FF00&
               Height          =   375
               Left            =   1320
               TabIndex        =   45
               Top             =   1800
               Width           =   3735
            End
            Begin VB.Label lbAlarmName 
               BackColor       =   &H000000FF&
               Caption         =   "0000 \n 1111"
               ForeColor       =   &H0000FF00&
               Height          =   855
               Left            =   1320
               TabIndex        =   44
               Top             =   960
               Width           =   3735
            End
            Begin VB.Label lbAlarmID 
               BackColor       =   &H000000FF&
               Caption         =   "0000"
               ForeColor       =   &H0000FF00&
               Height          =   375
               Left            =   1320
               TabIndex        =   43
               Top             =   480
               Width           =   615
            End
            Begin VB.Label lbAlarm 
               BackColor       =   &H000000FF&
               Caption         =   "處置方式:"
               ForeColor       =   &H0000FF00&
               Height          =   375
               Index           =   2
               Left            =   240
               TabIndex        =   42
               Top             =   1800
               Width           =   1095
            End
            Begin VB.Label lbAlarm 
               BackColor       =   &H000000FF&
               Caption         =   "錯誤名稱:"
               ForeColor       =   &H0000FF00&
               Height          =   375
               Index           =   1
               Left            =   240
               TabIndex        =   41
               Top             =   960
               Width           =   1095
            End
            Begin VB.Label lbAlarm 
               BackColor       =   &H000000FF&
               Caption         =   "錯誤代碼:"
               ForeColor       =   &H0000FF00&
               Height          =   375
               Index           =   0
               Left            =   240
               TabIndex        =   40
               Top             =   480
               Width           =   1095
            End
         End
         Begin VB.Label lblRecipeName 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6000
            TabIndex        =   27
            Top             =   120
            Width           =   3975
         End
         Begin VB.Label lbSec 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   270
            Left            =   1920
            TabIndex        =   6
            Top             =   120
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lbTemp 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   270
            Left            =   120
            TabIndex        =   5
            Top             =   2520
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Line linCourseVer 
            BorderColor     =   &H000000FF&
            Visible         =   0   'False
            X1              =   128
            X2              =   128
            Y1              =   24
            Y2              =   192
         End
         Begin VB.Line linCourseHor 
            BorderColor     =   &H000000FF&
            Visible         =   0   'False
            X1              =   0
            X2              =   128
            Y1              =   192
            Y2              =   192
         End
      End
   End
End
Attribute VB_Name = "frmPlotProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim Plot_Scale      As Single
      
'For Record Current Plot Point in Processing
Dim Plot_X_Point(50)    As Long
Dim Plot_Y_Point(50)    As Long
Dim Plot_X2_Point(50)    As Long
Dim Plot_Y2_Point(50)    As Long


'Dim sngProcessRecorder(65535, 5) As Single

'Dim strProcessRecipe(50, 6)   As String * 15
Dim strProcessRecipe(10) As String


Dim blnStartPIDLoop As Boolean
Dim iPreheatGainCount As Integer

Dim picObj(2) As PictureBox
Dim lngGasProfileColor(GB_GAS_MAX) As Long
    
    
'Cycle Run Test
Dim sngCycleRunTemp As Single

Public gbblnAlarmFlash As Boolean




Private Sub ChkPower_Click(Index As Integer)
Dim i As Integer
For i = 0 To 7
If i <> Index Then
If ChkPower(Index).value = 1 Then
ChkPower(i).Enabled = False
Else
ChkPower(i).Enabled = True
End If
End If
Next i
End Sub

Private Sub chkPumping_Click()
'    If Para.useTPump = 0 Then
'
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
'        End If
'    End If
     Dim countdown As Long
    Dim isMessageShown As Boolean
  
    isMessageShown = False
     If gbintReleaseOpenDelay > 0 Then
        Do
            DoEvents
                If frmDiagnosis.tmrSetReleaseOFF.Enabled Then
                    If Not isMessageShown Then
                        countdown = frmDiagnosis.tmrSetReleaseOFF.Interval / 1000
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
    If Not frmDiagnosis.tmrSetReleaseOFF.Enabled Then
        Call TogglePumping(chkPumping, Para.useTPump = 1, frmDiagnosis.tmrSetReleaseOFF)
        frmDiagnosis.chkPumping.value = IIf(PumpState, 1, 0)
    End If
    
End Sub


Private Sub chkPurge_Click()
    Dim sngGas(10) As Single
    Dim i As Integer
    
    If chkPurge.value = 1 Then
        If gbsngKeepPurge > 0 Then
            
            sngGas(0) = gbsngKeepPurge
            sngGas(1) = 0
            sngGas(2) = 0
            sngGas(3) = 0
            sngGas(4) = 0
            SetAO_MFC sngGas
            
           
            frmDiagnosis.tmrPurge.Enabled = True
            Kernel.IsPurge = 1
            Call frmHistory.AppendLogAlert(1, "Manual", 1015, "手動開啟破真空", 1)
            If gbintLoginRight > 1 Then Call mdifrmRTP.ShowTitleBar(False)
            frmDiagnosis.chkPurge.value = 1
        End If
        
    Else
        frmDiagnosis.tmrPurge.Enabled = False
        Kernel.IsPurge = 0
        For i = 0 To 4
                sngGas(i) = 0
        Next i
        SetAO_MFC sngGas
        Call mdifrmRTP.ShowTitleBar(True)
        frmDiagnosis.chkPurge.value = 0
    End If

'    If chkPurge.value = 1 Then
'        frmDiagnosis.sclSetGasValue(0).value = gbsngMaxGasSLMP(0) * gbsngGasFlowScale(0)
'        If gbsngKeepPurge > 0 Then
'            frmDiagnosis.sclSetGasValue(0).value = gbsngKeepPurge * gbsngGasFlowScale(0)
'        End If
'        frmDiagnosis.tmrPurge.Enabled = True
'        Kernel.IsPurge = 1
'        Call frmHistory.AppendLogAlert(1, "Manual", 1015, "手動開啟破真空", 1)
'        If gbintLoginRight > 1 Then Call mdifrmRTP.ShowTitleBar(False)
'    Else
'        frmDiagnosis.tmrPurge.Enabled = False
'        Kernel.IsPurge = 0
'        frmDiagnosis.txtSetGasValue(0).Text = 0
'        frmDiagnosis.sclSetGasValue(0).value = 0
'        Call mdifrmRTP.ShowTitleBar(True)
'    End If
End Sub

Private Sub chkAT_Click()
'    Dim i As Integer
'    Dim HoldData(3) As Integer
'    Dim value As Integer
'    value = chkAT.value
'
'    For i = 0 To 3
'        If Az1.blnUseLoop(i) = True Then HoldData(i) = value
'    Next i
'    Call frmAz1.WriteParas(125, HoldData, True)
'    For i = 0 To 3
'        If Az2.blnUseLoop(i) = True Then HoldData(i) = value
'    Next i
'    Call frmAz2.WriteParas(125, HoldData, True)
    gbintAz1ProcNo = IIf(chkAT.value = 1, 5, 4)
    gbintAz2ProcNo = IIf(chkAT.value = 1, 5, 4)
End Sub


Private Sub chkTest_Click()
    Dim value As Integer
    value = chkTest.value
    
    If Az1.blnUseAzbil Then
        gbintAz1ProcNo = value + 2
    End If
    If Az2.blnUseAzbil Then
        gbintAz2ProcNo = value + 2
    End If
        
End Sub

Private Sub Form_Activate()
    Dim i As Integer
        
    Me.Caption = frmRecipeEdit.lbRecipeName.Caption
    
'    For i = 1 To 7
'        If gbintMonitorTCActive(i - 1) = 1 Then
'            lbTCA(i).Visible = True
'            lbTC(i).Visible = True
'            chkTC(i).Visible = True
'        Else
'            lbTCA(i).Visible = False
'            lbTC(i).Visible = False
'            chkTC(i).Visible = False
'        End If
'    Next i

    For i = 0 To 7
        If gbintMonitorTCActive(i) = 1 Then
            lbTCA(i).Visible = True
            lbTC(i).Visible = True
            chkTC(i).Visible = True
        Else
            lbTCA(i).Visible = False
            lbTC(i).Visible = False
            chkTC(i).Visible = False
        End If
    Next i

    chkPlotColor(3).Caption = Trim(gbstrGasAlias(0))
    chkPlotColor(4).Caption = Trim(gbstrGasAlias(1))
    
    
    '110607 Josh Modified
    If gbintActiveModule_Vacuum = 1 Then
        fraVacFunc.Visible = True
        chkPlotColor(8).Visible = True
    Else
        fraVacFunc.Visible = False
        chkPlotColor(8).Visible = False
    End If
    
'    fraBarcode.Visible = IIf(gbintActiveModule_Barcode = 1, True, False)
    fraOxygen.Visible = IIf(gbintActiveModule_Oxygen = 1, True, False)
    
    If Para.UseCT = 0 Then
        fraBankH.Visible = False
        
    Else
        fraBankH.Visible = True
        
    End If
    fraMTC.Visible = IIf(Para.UseMTC = 1, True, False)
    fraMTCB.Visible = IIf(Para.UseMTCB = 1, True, False)
    tabMain.TabVisible(1) = IIf(Para.UseAz1 = 0, True, False)
    tabMain.TabVisible(2) = IIf(Para.UseAz1 = 0, True, False)
    tabMain.TabVisible(3) = IIf(Para.UseAz1 = 1, True, False)
    'fraAzbil.Visible = IIf(Para.UseAzbil = 1, True, False)
        
'    ShowStatus
End Sub

Private Sub Form_Initialize()
     
    Color_Choice = GB_ColorFewLightBlue ' GB_ColorLightGray
    'Coufigure ProcessTable Size
    Set_Percent = "100"  'Percent
    Set_Temp = 1000     'temperature
    Set_Time = 240 'sec

    With ProcessTable
        .Xsize = 10000
        .Ysize = 8000
        .ExtendXY_Size = 500
        .row = 40
        .Column = 15
'        .Xscale_size = 900 + .ExtendXY_Size
'        .Yscale_size = 600 + .ExtendXY_Size
        AssignX_Axis = Set_Time * 60
        AssignY_Axis = Set_Temp * 10
        X_Step = AssignX_Axis / (ProcessTable.Column + 1)
        Y_Step = AssignY_Axis / ProcessTable.row * 2
        .Xscale_size = AssignX_Axis + X_Step
        .Yscale_size = AssignY_Axis + Y_Step

    End With
End Sub

Private Sub Form_Load()
    Dim Return_Value    As Boolean
    Dim PicNO           As Long
    Dim i As Integer
        
    Plot_Scale = 10
    For PicNO = picProcess.LBound To picProcess.UBound
        picProcess(PicNO).ScaleMode = vbPixels
        picProcess(PicNO).AutoRedraw = True
        picProcess(PicNO).Width = ProcessTable.Xsize + ProcessTable.ExtendXY_Size * 2
        picProcess(PicNO).Height = ProcessTable.Ysize + ProcessTable.ExtendXY_Size * 2
'        picProcess(PicNO).Scale (-X_step, ProcessTable.Yscale_size + Y_step) _
'                            -(AssignX_Axis + X_step, -Y_step)
    Next PicNO
    Set_Percent = "100"  'Percent
    Set_Temp = 1000     'temperature
    Set_Time = 240 'sec
    Set picObj(0) = picProcess(0)
    Set picObj(1) = picProcess(0)
    Return_Value = Me.Plot_ProcessTable(0, 0, picObj)
    picProcess(0).AutoRedraw = True
    fraProcessChart.Height = picProcess(0).Height + 500
    fraProcessChart.Width = picProcess(0).Width + 500
    
    If CTDisplay = 1 Then
        hfgCTConfigProcess.Visible = True
        ShowCTTable
        
        For i = 0 To 59
            lbCT(i).Visible = False
        Next i
    End If
    
End Sub

Public Function Plot_ProcessTable(ByVal StartPic As Long, ByVal EndPic As Long, objPic() As PictureBox) As Boolean
    Dim m_X As Long
    Dim m_Y As Long
    Dim XPitch As Single
    Dim YPitch As Single
    Dim iCount As Integer
    Dim PercentPitch  As Integer
    Dim PercentStep As Integer
    Dim TempPitch As Integer
    Dim TempStep As Integer
    Dim TimePitch As Integer
    Dim TimeStep As Integer
    Dim m_Min As Integer
    Dim m_Sec As Integer
    Dim PicNO As Long
    Dim i As Integer
    
    With ProcessTable
        .Xsize = 10000
        .Ysize = 8000
        .ExtendXY_Size = 500
        .row = 40
        .Column = 15
        AssignX_Axis = Set_Time * 1000 '60    '120 sec 120 * 60 = 7200
        AssignY_Axis = Set_Temp * 10    '1200 degree 1200*10= 12000
        X_Step = AssignX_Axis / (ProcessTable.Column + 1)
        Y_Step = AssignY_Axis / ProcessTable.row * 2
        .Xscale_size = AssignX_Axis + X_Step
        .Yscale_size = AssignY_Axis + Y_Step
    End With
    
    For PicNO = StartPic To EndPic
        objPic(PicNO).FontSize = 12
        objPic(PicNO).Width = ProcessTable.Xsize + ProcessTable.ExtendXY_Size * 2
        objPic(PicNO).Height = ProcessTable.Ysize + ProcessTable.ExtendXY_Size * 2
        objPic(PicNO).Scale (-X_Step, ProcessTable.Yscale_size + Y_Step) _
                            -(AssignX_Axis + X_Step, -Y_Step)
    Next PicNO
    
    For PicNO = StartPic To EndPic
        objPic(PicNO).Cls
        objPic(PicNO).DrawWidth = 1
        Table_Xsize = ProcessTable.Xscale_size - X_Step 'Table's Weight
        Table_Ysize = ProcessTable.Yscale_size - Y_Step 'Table's Height
        XPitch = (Table_Xsize) / ProcessTable.Column
        YPitch = (Table_Ysize) / ProcessTable.row
        iCount = 0
        'This loop to plot line
        For m_Y = 0 To (Table_Ysize) Step YPitch
            For m_X = 0 To (Table_Xsize) Step XPitch
                objPic(PicNO).Line (m_X, 0)-(m_X, Table_Ysize)
            Next m_X
            
            If 0 = (iCount Mod 2) Then
                Color_Choice = vbBlack
            Else
                Color_Choice = GB_ColorLightBlue 'GB_ColorLightGray
            End If
            objPic(PicNO).Line (0, m_Y)-(Table_Xsize, m_Y), Color_Choice
            iCount = iCount + 1
        Next m_Y
    
       'Plot temperature,percent and sec. number
    
       PercentPitch = Val(Set_Percent) / (ProcessTable.row / 2)
       TempPitch = Val(Set_Temp) / (ProcessTable.row / 2)
       For i = 0 To 20
            'Percent number display
            PercentStep = i * PercentPitch
            objPic(PicNO).PSet (Table_Xsize + 10, (i * YPitch * 2) + _
                                    ScaleY((objPic(PicNO).FontSize / 2 * Table_Ysize / objPic(PicNO).Height) _
                                         , vbPoints, vbTwips)) 'sign point and get the currentX,Y
            If i = 20 Then
                objPic(PicNO).Print str(PercentStep) & "%"
            Else
                objPic(PicNO).Print str(PercentStep)
            End If
            'Temperature number Display
            TempStep = i * TempPitch
            objPic(PicNO).CurrentX = -XPitch '-ScaleX(objPic(PicNO).FontSize, vbPoints, vbTwips) * 3.5
            objPic(PicNO).CurrentY = (i * YPitch * 2) + _
                                         ScaleY((objPic(PicNO).FontSize / 2 * Table_Ysize / objPic(PicNO).Height) _
                                         , vbPoints, vbTwips)
             If i = 0 Then
                objPic(PicNO).Print str(TempStep) & " (C)"
            Else
                objPic(PicNO).Print str(TempStep)
            End If
        Next i
        
        'For time sec display
        TimePitch = Val(Set_Time) / ProcessTable.Column
        For i = 0 To 15
            TimeStep = i * TimePitch
            m_Min = Fix(TimeStep / 60)
            m_Sec = (TimeStep Mod 60)
            objPic(PicNO).CurrentX = (i * XPitch) - XPitch / 2 'ScaleY((objPic(PicNO).FontSize), vbPoints, vbTwips)
            objPic(PicNO).CurrentY = -Y_Step / 3 'table axis'
            objPic(PicNO).Print str(m_Min) & ":" & str(m_Sec)
        Next i
    Next PicNO

    Plot_ProcessTable = True
    
End Function

Public Sub DrawCurve()
    Dim i As Integer
    Dim sngTemp(50) As Single
    Dim mc(0 To 7) As Long

    ' ?置?色常量??
    mc(0) = &HC0FFC0
    mc(1) = &HFF00&
    mc(2) = &HFFFF00
    mc(3) = &HFFFF&
    mc(4) = &H40C0&
    mc(5) = &HFF00FF
    mc(6) = &HFF8080
    mc(7) = &H8080FF
        
    lngGasProfileColor(0) = GB_ColorPurple
    lngGasProfileColor(1) = GB_ColorGreen
    lngGasProfileColor(2) = GB_ColorFewLightBlue
    lngGasProfileColor(3) = &HFF00&
    lngGasProfileColor(4) = &HFFFF80
    For i = 0 To 23
        sngTemp(i) = CSng(gbsngDrawData(i))
        If sngTemp(i) < 0 Then sngTemp(i) = 0
    Next i
    
    Plot_Scale = CLng(AssignY_Axis / Set_Temp)
    Plot_X2_Point(0) = sngTemp(0)
    sngTemp(1) = sngTemp(1) * (AssignY_Axis / 100) * 10           'INTENSITY1
    sngTemp(2) = sngTemp(2) * Plot_Scale           'TC *10
    
    If gbintGasEnable(0) > 0 And gbsngMaxGasSLMP(0) > 0 Then sngTemp(4) = sngTemp(4) * (5 / gbsngMaxGasSLMP(0)) * (AssignY_Axis / 50) * gbintMFC_Ratio 'GAS 1~3
    If gbintGasEnable(1) > 0 And gbsngMaxGasSLMP(1) > 0 Then sngTemp(5) = sngTemp(5) * (5 / gbsngMaxGasSLMP(1)) * (AssignY_Axis / 50) * gbintMFC_Ratio 'GAS 1~3
    If gbintGasEnable(2) > 0 And gbsngMaxGasSLMP(2) > 0 Then sngTemp(6) = sngTemp(6) * (5 / gbsngMaxGasSLMP(2)) * (AssignY_Axis / 50) * gbintMFC_Ratio 'GAS 1~3
    If gbintGasEnable(3) > 0 And gbsngMaxGasSLMP(3) > 0 Then sngTemp(7) = sngTemp(7) * (5 / gbsngMaxGasSLMP(3)) * (AssignY_Axis / 50) * gbintMFC_Ratio 'GAS 1~3
    If gbintGasEnable(4) > 0 And gbsngMaxGasSLMP(4) > 0 Then sngTemp(8) = sngTemp(8) * (5 / gbsngMaxGasSLMP(4)) * (AssignY_Axis / 50) * gbintMFC_Ratio 'GAS 1~3
    
    If gbsngGaugeZoomIn > 0 And gbsngDrawData(9) < gbsngGaugeZoomIn Then sngTemp(9) = sngTemp(9) * 1000
    sngTemp(9) = sngTemp(9) * Plot_Scale           'Vacuum
    If sngTemp(9) > AssignY_Axis Then sngTemp(9) = AssignY_Axis
    
    sngTemp(11) = sngTemp(11) * Plot_Scale           'MTC1
    sngTemp(12) = sngTemp(12) * Plot_Scale           'MTC2
    sngTemp(13) = sngTemp(13) * Plot_Scale           'MTC3
    sngTemp(14) = sngTemp(14) * Plot_Scale           'MTC4
    sngTemp(15) = sngTemp(15) * Plot_Scale           'MTC5
    sngTemp(22) = sngTemp(22) * Plot_Scale           'MTC6
    sngTemp(23) = sngTemp(23) * Plot_Scale           'MTC7
        
        
    sngTemp(17) = sngTemp(17) * (AssignY_Axis / 100) * 10           'INTENSITY2
    sngTemp(18) = sngTemp(18) * (AssignY_Axis / 100) * 10           'INTENSITY3
    sngTemp(19) = sngTemp(19) * (AssignY_Axis / 100) * 10           'INTENSITY4
    sngTemp(20) = sngTemp(20) * (AssignY_Axis / 100) * 10           'INTENSITY5
    If Para.UseMTC = 1 Then
        For i = 24 To 31
            sngTemp(i) = CSng(gbsngDrawData(i))
            sngTemp(i) = sngTemp(i) * Plot_Scale           'MTC8~15
            If sngTemp(i) < 0 Then sngTemp(i) = 0
        Next i
    End If
    If Para.UseMTCB = 1 Then
        For i = 32 To 39
            sngTemp(i) = CSng(gbsngDrawData(i))
            sngTemp(i) = sngTemp(i) * Plot_Scale           'MTC16~23
            If sngTemp(i) < 0 Then sngTemp(i) = 0
        Next i
    End If
    
    If Plot_X2_Point(0) >= 0 Then
        
        picProcess(0).DrawWidth = 1
        
        If chkTC(0).value = 1 Then _
            picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(2))-(Plot_X2_Point(0), sngTemp(2)), vbRed 'TC
        If chkTC(1).value = 1 Then _
            picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(11))-(Plot_X2_Point(0), sngTemp(11)), &HFF00& 'MTC1
        If chkTC(2).value = 1 Then _
            picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(12))-(Plot_X2_Point(0), sngTemp(12)), &HFFFF00 'MTC2
        If chkTC(3).value = 1 Then _
            picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(13))-(Plot_X2_Point(0), sngTemp(13)), &HFFFF&  'MTC3
        If chkTC(4).value = 1 Then _
            picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(14))-(Plot_X2_Point(0), sngTemp(14)), &H40C0&  'MTC4
        If chkTC(5).value = 1 Then _
            picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(15))-(Plot_X2_Point(0), sngTemp(15)), &HFF00FF 'MTC5
        If chkTC(6).value = 1 Then _
            picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(22))-(Plot_X2_Point(0), sngTemp(22)), &HFF8080 'MTC6
        If chkTC(7).value = 1 Then _
            picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(23))-(Plot_X2_Point(0), sngTemp(23)), &H8080FF 'MTC7
               
        If Para.UseMTC = 1 Then
            If chkTC(8).value = 1 Then _
                picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(24))-(Plot_X2_Point(0), sngTemp(24)), &HC0FFC0 'MTC8
            If chkTC(9).value = 1 Then _
                picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(25))-(Plot_X2_Point(0), sngTemp(25)), &HFF00& 'MTC9
            If chkTC(10).value = 1 Then _
                picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(26))-(Plot_X2_Point(0), sngTemp(26)), &HFFFF00  'MTC10
            If chkTC(11).value = 1 Then _
                picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(27))-(Plot_X2_Point(0), sngTemp(27)), &HFFFF&  'MTC11
            If chkTC(12).value = 1 Then _
                picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(28))-(Plot_X2_Point(0), sngTemp(28)), &H40C0& 'MTC12
            If chkTC(13).value = 1 Then _
                picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(29))-(Plot_X2_Point(0), sngTemp(29)), &HFF00FF 'MTC13
            If chkTC(14).value = 1 Then _
                picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(30))-(Plot_X2_Point(0), sngTemp(30)), &HFF8080 'MTC14
            If chkTC(15).value = 1 Then _
                picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(31))-(Plot_X2_Point(0), sngTemp(31)), &H8080FF 'MTC15
        End If
        
        If Para.UseMTCB = 1 Then
            For i = 16 To 23
                If chkTC(i).value = 1 Then _
                    picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(i + 16))-(Plot_X2_Point(0), sngTemp(i + 16)), mc(i - 16) 'MTC16
            Next i
        End If
        
        
        
        picProcess(0).DrawWidth = 3
        If chkShowLine(0).value = 1 And ChkPower(0).value = 1 Then
            picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(1))-(Plot_X2_Point(0), sngTemp(1)), vbYellow 'INTENSITY1
        End If
        
       If chkShowLine(0).value = 1 And ChkPower(1).value = 1 Then
            picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(17))-(Plot_X2_Point(0), sngTemp(17)), vbYellow 'INTENSITY2
        End If
        
        If chkShowLine(0).value = 1 And ChkPower(2).value = 1 Then
            picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(18))-(Plot_X2_Point(0), sngTemp(18)), vbYellow 'INTENSITY3
        End If
        
        
        If chkShowLine(0).value = 1 And ChkPower(3).value = 1 Then
            picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(19))-(Plot_X2_Point(0), sngTemp(19)), vbYellow 'INTENSITY4
        End If
        
        
        If chkShowLine(0).value = 1 And ChkPower(4).value = 1 Then
            picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(20))-(Plot_X2_Point(0), sngTemp(20)), vbYellow 'INTENSITY5
        End If
        
        For i = 4 To 8
            picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(i))-(Plot_X2_Point(0), sngTemp(i)), lngGasProfileColor(i - 4) 'GAS1
        Next i
        
        '110607 Josh Added
        If chkShowLine(4).value = 1 Then
            picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(9))-(Plot_X2_Point(0), sngTemp(9)), GB_ColorOrange
        End If
        
        'Plot  the Course Hor (AssignY_Axis and Ver
        linCourseHor.BorderColor = GB_ColorPlusGray
        linCourseHor.Visible = True
        linCourseHor.x1 = 0
        linCourseHor.y1 = sngTemp(2)
        linCourseHor.x2 = sngTemp(0)
        linCourseHor.y2 = sngTemp(2)
        linCourseHor.Refresh
        
        linCourseVer.BorderColor = GB_ColorPlusGray
        linCourseVer.Visible = True
        linCourseVer.x1 = sngTemp(0)
        linCourseVer.y1 = AssignY_Axis
        linCourseVer.x2 = sngTemp(0)
        linCourseVer.y2 = sngTemp(2)
        linCourseVer.Refresh
        
        lbTemp.Visible = True
        lbTemp.Top = linCourseHor.y1
        lbTemp.Left = linCourseHor.x1 + 50
        lbTemp.Caption = Format(sngTemp(2) / Plot_Scale, "0")
        
        lbSec.Visible = True
        lbSec.Top = linCourseVer.y1 '+ 200
        lbSec.Left = linCourseVer.x1 - 50
        lbSec.Caption = Format(sngTemp(0) / 1000, "0")
        
        
        For i = 0 To GB_MAX_DRAW_COL
            Plot_X_Point(i) = sngTemp(0)
            Plot_Y_Point(i) = sngTemp(i)
        Next i
       
    End If
    
End Sub

Public Sub PlotProcessChart(ByVal GetName As String, GetData() As String)
    Dim tmpStr      As String
    Dim Ci, Cj, Ck  As Long
    Dim tmpVal, SelPic     As Long
    Dim GetPlotData(50) As Single
    Dim i As Integer
    Dim j As Integer
    Dim temp As Integer
    
    lngGasProfileColor(0) = GB_ColorPurple
    lngGasProfileColor(1) = GB_ColorGreen
    lngGasProfileColor(2) = GB_ColorFewLightBlue
    lngGasProfileColor(3) = &HFF00&
    lngGasProfileColor(4) = &HFFFF80
    picProcess(0).DrawWidth = 3

    Plot_Scale = CLng(AssignY_Axis / Set_Temp)
    GetPlotData(0) = Val(GetData(0)) 'Time
    If GetPlotData(0) > Set_Time * 1000 Then
    Exit Sub
    End If
    GetPlotData(1) = Val(GetData(1)) * Plot_Scale           'TC *10
    GetPlotData(2) = Val(GetData(2)) * Plot_Scale           'PM
    GetPlotData(3) = Val(GetData(3)) * (AssignY_Axis / 100) * 10 'INTENSITY
    
    If gbintGasEnable(0) > 0 And gbsngMaxGasSLMP(0) > 0 Then GetPlotData(4) = Val(GetData(4)) * (5 / gbsngMaxGasSLMP(0)) * (AssignY_Axis / 50) * 10 'N2
    If gbintGasEnable(1) > 0 And gbsngMaxGasSLMP(1) > 0 Then GetPlotData(5) = Val(GetData(5)) * (5 / gbsngMaxGasSLMP(1)) * (AssignY_Axis / 50) * 10 'N2
    If gbintGasEnable(2) > 0 And gbsngMaxGasSLMP(2) > 0 Then GetPlotData(6) = Val(GetData(6)) * (5 / gbsngMaxGasSLMP(2)) * (AssignY_Axis / 50) * 10 'N2
    If gbintGasEnable(3) > 0 And gbsngMaxGasSLMP(3) > 0 Then GetPlotData(7) = Val(GetData(7)) * (5 / gbsngMaxGasSLMP(3)) * (AssignY_Axis / 50) * 10 'N2
    If gbintGasEnable(4) > 0 And gbsngMaxGasSLMP(4) > 0 Then GetPlotData(8) = Val(GetData(8)) * (5 / gbsngMaxGasSLMP(4)) * (AssignY_Axis / 50) * 10 'N2
    
    GetPlotData(9) = Val(GetData(9)) * Plot_Scale
    GetPlotData(11) = Val(GetData(11)) * Plot_Scale           'TC *10
    GetPlotData(12) = Val(GetData(12)) * Plot_Scale           'TC *10
    GetPlotData(13) = Val(GetData(13)) * Plot_Scale           'TC *10
    GetPlotData(14) = Val(GetData(14)) * Plot_Scale           'TC *10
    GetPlotData(15) = Val(GetData(15)) * Plot_Scale           'TC *10
    
    GetPlotData(16) = Val(GetData(16)) * (AssignY_Axis / 100) * 10 'INTENSITY
    GetPlotData(17) = Val(GetData(17)) * (AssignY_Axis / 100) * 10 'INTENSITY
    GetPlotData(18) = Val(GetData(18)) * (AssignY_Axis / 100) * 10 'INTENSITY
    GetPlotData(19) = Val(GetData(19)) * (AssignY_Axis / 100) * 10 'INTENSITY
    GetPlotData(20) = Val(GetData(20)) * (AssignY_Axis / 100) * 10 'INTENSITY
    GetPlotData(22) = Val(GetData(22)) * Plot_Scale           'TC *10
    GetPlotData(23) = Val(GetData(23)) * Plot_Scale           'TC *10
    GetPlotData(24) = Val(GetData(24)) * Plot_Scale           'TC *10
    
    GetPlotData(25) = Val(GetData(25)) * Plot_Scale           'TC *10
    GetPlotData(26) = Val(GetData(26)) * Plot_Scale           'TC *10
    GetPlotData(27) = Val(GetData(27)) * Plot_Scale           'TC *10
    GetPlotData(28) = Val(GetData(28)) * Plot_Scale           'TC *10
    GetPlotData(29) = Val(GetData(29)) * Plot_Scale           'TC *10
    GetPlotData(30) = Val(GetData(30)) * Plot_Scale           'TC *10
    GetPlotData(31) = Val(GetData(31)) * Plot_Scale           'TC *10
    GetPlotData(32) = Val(GetData(32)) * Plot_Scale           'TC *10
    
'    temp = gbintMaxGasEnable
'    If gbintMaxGasEnable > 2 Then
'        temp = 2
'    End If
'    For i = 0 To 4
'        GetPlotData(GB_PROCESS_GAS1 + i) = Val(GetData(GB_PROCESS_GAS1 + i)) _
'            * (5 / gbsngMaxGasSLMP(i)) * (AssignY_Axis / 50)  'N2
'
'
'    Next i
    
    '110607 Josh Added
    If gbintActiveModule_Vacuum Then
        If GetPlotData(9) > 760 Then
            GetPlotData(9) = 760
        End If
        GetPlotData(9) = Val(GetData(9)) * Plot_Scale          'Vacuum
        If GetPlotData(9) > AssignY_Axis Then
            GetPlotData(9) = AssignY_Axis
        End If
    End If
    
    
    For i = 0 To GB_MAX_DRAW_COL
        If GetPlotData(i) < -10 Or GetPlotData(i) > 1000000 Then
            GetPlotData(i) = 0
        End If
    Next i
    
    tmpStr = GetData(0)
    
    SelPic = 0
    If Val(tmpStr) >= 0 Then
        'Draw Temperature
        picProcess(0).DrawWidth = 1
        
        If chkShowLine(1).value = 1 Then
            picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(0))-(GetPlotData(0), GetPlotData(1)), vbRed 'TC
            
            'Rev8.0.1.7
            If chkTC(1).value = 1 And gbintMonitorTCActive(0) = 1 Then _
                picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(11))-(GetPlotData(0), GetPlotData(11)), &HFF00&        'TC2
            If chkTC(2).value = 1 And gbintMonitorTCActive(1) = 1 Then _
                picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(12))-(GetPlotData(0), GetPlotData(12)), &HFFFF00        'TC2
            If chkTC(3).value = 1 And gbintMonitorTCActive(2) = 1 Then _
                picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(13))-(GetPlotData(0), GetPlotData(13)), &HFFFF&     'TC3
            If chkTC(4).value = 1 And gbintMonitorTCActive(3) = 1 Then _
                picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(14))-(GetPlotData(0), GetPlotData(14)), &H40C0&        'TC4
            If chkTC(5).value = 1 And gbintMonitorTCActive(4) = 1 Then _
                picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(15))-(GetPlotData(0), GetPlotData(15)), &HFF00FF               'TC5
            If chkTC(6).value = 1 And gbintMonitorTCActive(5) = 1 Then _
                picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(22))-(GetPlotData(0), GetPlotData(22)), &HFF8080                  'TC6
            If chkTC(7).value = 1 And gbintMonitorTCActive(6) = 1 Then _
                picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(23))-(GetPlotData(0), GetPlotData(23)), &H8080FF                     'TC6
        
        End If
        'Draw Intensity
        picProcess(0).DrawWidth = 3
        If chkShowLine(0).value = 1 Then
            picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(3))-(GetPlotData(0), GetPlotData(3)), vbYellow 'INTENSITY
        End If
        
'        temp = gbintMaxGasEnable
'        If gbintMaxGasEnable > 2 Then
'            temp = 2
'        End If
'        For i = 0 To temp
'            picProcess(0).Line (Plot_X_Point(GB_PROCESS_GAS1 - 1 + i), Plot_Y_Point(GB_PROCESS_GAS1 - 1 + i))-(GetPlotData(0), GetPlotData(GB_PROCESS_GAS1 + i)), lngGasProfileColor(i) 'GAS1
'        Next i
        
        picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(4))-(GetPlotData(0), GetPlotData(4)), lngGasProfileColor(0) 'GAS1
        picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(5))-(GetPlotData(0), GetPlotData(5)), lngGasProfileColor(1) 'GAS2
        picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(6))-(GetPlotData(0), GetPlotData(6)), lngGasProfileColor(2) 'GAS3
        picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(7))-(GetPlotData(0), GetPlotData(7)), lngGasProfileColor(3) 'GAS4
        picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(8))-(GetPlotData(0), GetPlotData(8)), lngGasProfileColor(4) 'GAS5
                
        
        '110607 Josh Added
        If chkShowLine(4).value = 1 Then
            If gbintActiveModule_Vacuum And GetPlotData(9) > 0 Then
               picProcess(0).Line (Plot_X_Point(0), Plot_Y_Point(9))-(GetPlotData(0), GetPlotData(9)), GB_ColorOrange
            End If
        End If
        
               
        'Plot  the Course Hor (AssignY_Axis and Ver
        linCourseHor.BorderColor = GB_ColorPlusGray
        linCourseHor.Visible = True
        linCourseHor.x1 = 0
        linCourseHor.y1 = GetPlotData(1)
        linCourseHor.x2 = GetPlotData(0)
        linCourseHor.y2 = GetPlotData(1)
        If frmRecipeEdit.intRecipeTempInputType = 1 Then
            linCourseHor.y1 = GetPlotData(1)
            linCourseHor.y2 = GetPlotData(1)
        ElseIf frmRecipeEdit.intRecipeTempInputType = 2 Then
            linCourseHor.y1 = GetPlotData(2)
            linCourseHor.y2 = GetPlotData(2)
        End If
        
        linCourseHor.Refresh
        linCourseVer.BorderColor = GB_ColorPlusGray
        linCourseVer.Visible = True
        linCourseVer.x1 = GetPlotData(0)
        linCourseVer.y1 = AssignY_Axis
        linCourseVer.x2 = GetPlotData(0)
        linCourseVer.y2 = GetPlotData(1)
        If frmRecipeEdit.intRecipeTempInputType = 1 Then
            linCourseVer.y1 = AssignY_Axis
            linCourseVer.y2 = GetPlotData(1)
        ElseIf frmRecipeEdit.intRecipeTempInputType = 2 Then
            linCourseVer.y1 = AssignY_Axis
            linCourseVer.y2 = GetPlotData(2)
        End If
        
        linCourseVer.Refresh
        lbTemp.Visible = True
        lbTemp.Top = linCourseHor.y1
        lbTemp.Left = linCourseHor.x1 + 50
        lbTemp.Caption = Format(Kernel.sngTC(0), "0")
        '
        lbSec.Visible = True
        lbSec.Top = linCourseVer.y1 '+ 200
        lbSec.Left = linCourseVer.x1 - 50
        lbSec.Caption = Format(GetPlotData(0) / 1000, "0")
        
        Plot_X_Point(0) = CLng(GetPlotData(0))
        For i = 1 To 32
            Plot_Y_Point(i) = CLng(GetPlotData(i))
        Next i
              
        
    
    End If
    
    Erase GetPlotData
    'DoEvents
End Sub

Public Sub PlotProcessChartClean()
    Dim i As Integer
    
    For i = picProcess.LBound To picProcess.UBound
        picProcess(i).Cls
    Next i
    Call Plot_ProcessTable(picProcess.LBound, picProcess.UBound, picObj)
    fraProcessChart.Height = picProcess(0).Height + 500
    fraProcessChart.Width = picProcess(0).Width + 500
    lblRecipeName.Caption = ""
    frmRecipeEdit.lbRecipeName.Caption = " "
    Kernel.strCurrRecipe = ""
End Sub

Public Sub InitPlotChart()
    Dim i As Integer
    
    InitPIDParameter
    
    frmRecipeEdit.RecipeLoad
    If Not PlotRecipeChart Then Exit Sub
    
    For i = 0 To 20
        Plot_X_Point(i) = 0
        Plot_Y_Point(i) = 0
    Next i
    InitPIDParameter
    SetPIDParameter 0, _
                                    frmRecipeEdit.sngRecipeProportional, _
                                    frmRecipeEdit.sngRecipeIntegral, _
                                    frmRecipeEdit.sngRecipeDerivational, _
                                    frmRecipeEdit.sngRecipePredit, _
                                    frmRecipeEdit.sngRecipeFeedForward
    blnStartPIDLoop = False
    
End Sub


Private Sub lbIntensityAz_DblClick(Index As Integer)
'    Dim i As Integer
'    Dim lngRet                As Long
'
'    If Az1.blnUseAzbil = True And Az1.blnUseLoop = True Then
'        Dim tc As Single
'        Dim SetTemp As Single
'        Dim sum As Single
'        Dim ss As String
'
'
'        tc = Az1.sngPV(Index)
'        SetTemp = m_sngSetTemperature
'        sum = MultiLoop.sngLoopIN(Index) * (SetTemp / tc) ^ 3
'        ss = Format(sum, "0.000000")
'        lngRet = WritePrivateProfileString("MultiLoop", "IN" & CStr(Index), ss, Kernel.strCurrRecipeFile)
'        MultiLoop.sngLoopIN(Index) = sum
'        lbIntensityLabel(Index).Caption = Format(sum * 10000, "0.00000")
'
'
'    End If
End Sub



Public Sub lbIntensityAz1_DblClick(Index As Integer)
    CalAzbilRT (Index)
    
End Sub



Private Sub lbIntensityLoop_DblClick(Index As Integer)
    Dim i As Integer
    Dim lngRet                As Long
    
    If MultiLoop.blnUseMultiLoop = True Then
        If MultiLoop.blnUseLoop(Index) = True Then
            Dim tc As Double
            Dim mtc As Double
            Dim sum As Double
                        
            tc = Kernel.sngTC(MultiLoop.intLoopTC(Index))
            sum = 0
            If MultiLoop.intLoopMA(Index) > 0 Then
                mtc = Kernel.sngTC(MultiLoop.intLoopMA(Index) - 1)
                sum = sum + mtc / tc
                i = i + 1
            End If
            If MultiLoop.intLoopMB(Index) > 0 Then
                mtc = Kernel.sngTC(MultiLoop.intLoopMB(Index) - 1)
                sum = sum + mtc / tc
                i = i + 1
            End If
            If MultiLoop.intLoopMC(Index) > 0 Then
                mtc = Kernel.sngTC(MultiLoop.intLoopMC(Index) - 1)
                sum = sum + mtc / tc
                i = i + 1
            End If
            If MultiLoop.intLoopMD(Index) > 0 Then
                mtc = Kernel.sngTC(MultiLoop.intLoopMD(Index) - 1)
                sum = sum + mtc / tc
                i = i + 1
            End If
            If MultiLoop.intLoopME(Index) > 0 Then
                mtc = Kernel.sngTC(MultiLoop.intLoopME(Index) - 1)
                sum = sum + mtc / tc
                i = i + 1
            End If
            If MultiLoop.intLoopMF(Index) > 0 Then
                mtc = Kernel.sngTC(MultiLoop.intLoopMF(Index) - 1)
                sum = sum + mtc / tc
                i = i + 1
            End If
            If MultiLoop.intLoopMG(Index) > 0 Then
                mtc = Kernel.sngTC(MultiLoop.intLoopMG(Index) - 1)
                sum = sum + mtc / tc
                i = i + 1
            End If
            If MultiLoop.intLoopMH(Index) > 0 Then
                mtc = Kernel.sngTC(MultiLoop.intLoopMH(Index) - 1)
                sum = sum + mtc / tc
                i = i + 1
            End If
            If MultiLoop.intLoopMJ(Index) > 0 Then
                mtc = Kernel.sngTC(MultiLoop.intLoopMJ(Index) - 1)
                sum = sum + mtc / tc
                i = i + 1
            End If

            If i > 0 Then
                sum = sum / i
                MultiLoop.sngLoopRT(Index) = MultiLoop.sngLoopRT(Index) * sum
                
                lngRet = WritePrivateProfileString("MultiLoop", "RT" & CStr(Index), CStr(MultiLoop.sngLoopRT(Index)), Kernel.strCurrRecipeFile)
            End If
                     
        
        End If
        
    End If
End Sub

Private Sub lbName_DblClick(Index As Integer)
    If Index = 4 And gbintLoginRight < 3 And gbintActiveModule_PNRecipe = 1 Then
        frmPNRecipe.Show
    End If
End Sub





Private Sub picProcess_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyControl Then
        gbintPlayFakeBall = gbintPlayFakeBall + 1
        If gbintPlayFakeBall > 2 Then
            gbblnPlayFakeBall = True
            gbintPlayFakeBall = 0
            
        End If
    End If
End Sub



Private Sub tmrAlarmFlash_Timer()
    gbblnAlarmFlash = Not gbblnAlarmFlash
    fraAlarm.Visible = gbblnAlarmFlash
    fraProcessChart.BackColor = IIf(gbblnAlarmFlash = True, &HFF&, &H8000000F)
    frmDiagnosis.shpAlarm.BackColor = IIf(gbblnAlarmFlash = True, &HFF&, &H80&)
    
End Sub

Private Sub tmrIdleCheck_Timer()
    If Kernel.IsRun = 1 Then
        tmrIdleCheck.Enabled = False
        gbsngIdleCount = 0
        Kernel.IsNeedTestRun = 0
    Else
        gbsngIdleCount = gbsngIdleCount + 10
        If gbsngIdleCount > gbsngIdleWarning And gbsngIdleWarning > 0 Then
            tmrIdleCheck.Enabled = False
            gbsngIdleCount = 0
            Kernel.IsNeedTestRun = 1
            SetLampCooling False
            ShowMessageOK "系統閒置時間過長,CDA關閉,請先執行報廢片空跑!"
            Call frmHistory.AppendLogAlert(1, "Alarm", 3008, "系統閒置時間過長,CDA關閉", 1)
        End If
    End If
End Sub

Public Function PlotRecipeChart() As Boolean
    Dim picTemp(1) As PictureBox
    
    Set picTemp(0) = Me.picProcess(0)
    Set picTemp(1) = Me.picProcess(1)
    
    PlotRecipeChart = Plot_Recipe(m_Recipe, 0, 0, picTemp)
    linCourseHor.Visible = False
    linCourseVer.Visible = False
    lbSec.Visible = False
    lbTemp.Visible = False

End Function

Public Sub DrawProcessChartData(sngPower As Single, sngCurrTemperature As Single, lngTimeStamp As Long)
    Dim sngOutput   As Single
    Dim strValue(50)    As String
    Dim i As Integer
    Dim temp As Integer
    
    strValue(0) = CStr(lngTimeStamp)
    strValue(1) = CStr(sngCurrTemp(0))
    strValue(2) = CStr(gbsngPM)
    strValue(3) = CStr(sngPower)
    strValue(4) = CStr(SysAI.sngMFC(0))
    strValue(5) = CStr(SysAI.sngMFC(1))
    strValue(6) = CStr(SysAI.sngMFC(2))
    strValue(7) = CStr(SysAI.sngMFC(3))
    strValue(8) = CStr(SysAI.sngMFC(4))
    strValue(9) = CStr(Kernel.sngPressure)
    If gbsngGaugeZoomIn > 0 And Kernel.sngPressure < gbsngGaugeZoomIn Then
        strValue(7) = CStr(Kernel.sngPressure * 1000)
    End If
    
    
    strValue(11) = CStr(sngCurrTemp(1))
    
    
    'Rev9.0.0.0
    strValue(12) = CStr(sngCurrTemp(2))
    strValue(13) = CStr(sngCurrTemp(3))
    strValue(14) = CStr(sngCurrTemp(4))
    strValue(15) = CStr(sngCurrTemp(5))
    
        
    'txtIntensity.Text = Format(CStr(sngPower * 10), "0")
    
    strValue(16) = CStr(sngPower)
    strValue(17) = CStr(gbsngPower(1))
    strValue(18) = CStr(gbsngPower(2))
    strValue(19) = CStr(gbsngPower(3))
    strValue(20) = CStr(gbsngPower(4))
    
'    temp = gbintMaxGasEnable
'    If gbintMaxGasEnable > 2 Then
'        temp = 2
'    End If
'    For i = 0 To temp
'        strValue(GB_PROCESS_GAS1 + i) = CStr(SysAI.sngMFC(i))
'    Next i
    strValue(21) = Format(gbsngOxygenPPM, "0000")
    strValue(22) = CStr(sngCurrTemp(6))
    strValue(23) = CStr(sngCurrTemp(7))
    
    Me.PlotProcessChart "1", strValue
    'DoEvents

End Sub

Public Sub SetPIDValue()
    lbPIDValue(0).Caption = CStr(frmRecipeEdit.sngRecipeProportional) ', "0.0000")
    lbPIDValue(6).Caption = CStr(frmRecipeEdit.sngRecipeProportional2) ', "0.0000")
    lbPIDValue(1).Caption = CStr(frmRecipeEdit.sngRecipeIntegral) ', "0.0000")
    lbPIDValue(5).Caption = CStr(frmRecipeEdit.sngRecipeIntegral2) ', "0.0000")
End Sub

Public Function SetFormat(data As Single, Digit As Integer) As String
Dim DigitStr As String
Dim ForMatString As String
Dim i As Integer
DigitStr = "0."
For i = 1 To Digit
DigitStr = DigitStr + "0"
Next i
ForMatString = Format(data, DigitStr)
SetFormat = ForMatString
End Function
 


Public Sub ShowStatus()
    Dim i As Integer

    lbVacuum.BackColor = IIf(SysDI.IsChamberGaugeL = 0, &HFF00&, &H8000000F)
    lbVacuum.Caption = IIf(Para.useTPump = 1, Format(Kernel.sngPressure, "0.000000"), Format(Kernel.sngPressure, "0.000"))
    
    
    If Kernel.IsAlarm = 0 Then
        fraProcessChart.BackColor = IIf(Kernel.IsRun = 0, &H8000000F, &HFF00&)
        tmrAlarmFlash.Enabled = False
        fraAlarm.Visible = False
    Else
        fraAlarm.Visible = True
    End If
    For i = 0 To 7
     lbTCA(i).Caption = gbstrNameTC(i)
    Next i
    If GbTcoffset_Switch = 1 And GbHoldState = True Then
       For i = 0 To 4
        lbTC(i).Caption = SetFormat(Kernel.sngTC(i) + TempOffset(i), gbintPrecisionDigit(i))
       Next i
   Else
       For i = 0 To 4
        lbTC(i).Caption = SetFormat(Kernel.sngTC(i), gbintPrecisionDigit(i))
       Next i
   End If
    For i = 5 To 7
        lbTC(i).Caption = SetFormat(Kernel.sngTC(i), gbintPrecisionDigit(i))
    Next i
    
    lbOxygen.Caption = Format(Kernel.sngOxygen, "0.00")
    
    If Para.UseMTC = 1 Then
        For i = 8 To 15
            lbTC(i).Caption = Format(Kernel.sngTC(i), "0.0")
        Next i
    End If
    If Para.UseMTCB = 1 Then
        For i = 16 To 23
            lbTC(i).Caption = Format(Kernel.sngTC(i), "0.0")
        Next i
    End If
    For i = 0 To 5
        TxtGas(i).BackColor = IIf(Kernel.sngCurrOutMFC(i) > 0, &HFFFFC0, &H80000005)
        If lbNameMFC(i).Caption <> "APC" Then
        If gblngAI_MFC_Read(i) >= 0 Then TxtGas(i).text = Format(SysAI.sngMFC(i), "0.0")
        Else
        Dim j As Integer
        For j = 5 To 7
        If gbstrNameTC(j) = "PS" Then
         TxtGas(i).text = Format(Kernel.sngTC(j), "0.0")
        Exit For
        End If
        Next j
        End If
'        If gblngAI_MFC_Read(i) >= 0 Then txtGas(i).text = Format(SysAI.sngMFC(i), "0.0")
    Next i

    If Kernel.strCurrRecipe <> "" Then
        lblRecipeName.Caption = Kernel.strCurrRecipe
        Me.Caption = Kernel.strCurrRecipe
    End If
    
    If Para.UseCT = 1 Then
        If CTDisplay = 0 Then
            For i = 0 To 59
                If Kernel.intOverCT(i) = 0 And Kernel.intUnderCT(i) = 0 Then
                    lbCT(i).Caption = Format(Kernel.dblCT(i), "#0.0")
                    lbCT(i).BackColor = &H8000000F
                Else
                    If Kernel.intOverCT(i) = 1 Then lbCT(i).BackColor = &HFF&
                    If Kernel.intUnderCT(i) = 1 Then lbCT(i).BackColor = &HFF00FF
                End If
            Next i
        Else
            ShowCTTable
        End If
    End If
    If MultiLoop.blnUseMultiLoop = True Then
        For i = 0 To GB_MAX_LOOPS - 1
            lbIntensityLoop(i).Caption = Format(MultiLoop.sngLoopOut(i) * 10, "0.00")
            lbIntensityLoop(i).BackColor = IIf(MultiLoop.sngLoopOut(i) > 0, &H80FFFF, &H80000005)
        Next i
        Kernel.sngIntensity = MultiLoop.sngLoopOut(0)
    End If
    
    If Para.UseAz1 = 1 Then
        For i = 0 To 3
            lbIntensityAz1(i).Caption = Format(Az1.sngMV(i), "0.00")
            If Az1.blnStart(i) Then
                lbIntensityAz1(i).BackColor = IIf(Az1.intMode(i) = 0, &H80FFFF, &H80FFFF)
            Else
                lbIntensityAz1(i).BackColor = &H80000005
            End If
        Next i
    End If
    If Para.UseAz2 = 1 Then
        For i = 0 To 3
            lbIntensityAz1(i + 4).Caption = Format(Az2.sngMV(i), "0.00")
            If Az2.blnStart(i) Then
                lbIntensityAz1(i + 4).BackColor = IIf(Az2.intMode(i) = 0, &H80FFFF, &H80FFFF)
            Else
                lbIntensityAz1(i + 4).BackColor = &H80000005
            End If
        Next i
    End If
    
    
    txtIntensity.text = Format(Kernel.sngIntensity * 10, "0.00")
    txtIntensity.BackColor = IIf(Kernel.sngIntensity > 0, &H80FFFF, &H80000005)
    lbScanCount.Caption = CurrProc.lngCurrentTime - CurrProc.lngPrevTime
    'lbOutput.Caption = CurrProc.sngOutput
    lbOutput.Caption = m_sngSetTemperature
End Sub

Private Sub txtBN_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        gbstrBN = txtBN.Text
'        txtID1.SelStart = 0
'        txtID1.SelLength = 100
'        txtID1.SetFocus
'    End If
End Sub

Private Sub txtID1_KeyDown(KeyCode As Integer, Shift As Integer)
    
    '
End Sub

Private Sub txtID1_KeyPress(KeyAscii As Integer)
    Dim StrFileName As String
    Dim strFilePath As String
    Dim strDir As String
    Dim lngRet                As Long
    If KeyAscii = 13 Then
        gbstrID1 = txtID1.text
        'strFilePath = gbSystemPath & "\Recipe" & "\op\"
        strFilePath = gbstrRecipeFilePath
        strDir = dir(strFilePath, vbDirectory)
        If strDir = "" Then MkDir strFilePath
        
        gbstrPNRecipeFile = strFilePath & gbstrID1 & ".rcp"
        gbblnPNLoad = True
        frmRecipeEdit.cmdRecipeOpen_Click
        gbblnPNLoad = False
    '    txtID2.SelStart = 0
    '    txtID2.SelLength = 100
    '    txtID2.SetFocus
    End If
End Sub

Private Sub txtID2_KeyPress(KeyAscii As Integer)
    Dim StrFileName As String
    Dim strFilePath As String
    Dim strDir As String
    Dim lngRet                As Long
    If KeyAscii = 13 Then
        gbstrID2 = txtID2.text
        
        'strFilePath = gbSystemPath & "\Recipe" & "\op\"
        strFilePath = gbstrRecipeFilePath
        strDir = dir(strFilePath, vbDirectory)
        If strDir = "" Then MkDir strFilePath
        StrFileName = strFilePath & Mid(gbstrID2, 2, 3) & Mid(gbstrID2, 7, 1) & ".rcp"
        If FileExists(StrFileName) = True Then
        
            gbstrPNRecipeFile = StrFileName
        
            gbblnPNLoad = True
            frmRecipeEdit.cmdRecipeOpen_Click
            gbblnPNLoad = False
        Else
            ShowMessageOK "找不到檔案 " & StrFileName
        End If
    
    End If


'    If KeyAscii = 13 Then
'        gbstrID2 = txtID2.Text
'        txtPN.SelStart = 0
'        txtPN.SelLength = 100
'        txtPN.SetFocus
'        mdifrmRTP.tbrRTP_ButtonClick mdifrmRTP.tbrRTP.Buttons("iRun")
'
'    End If
End Sub

Private Sub txtPN_KeyPress(KeyAscii As Integer)
    
    
    If KeyAscii = 13 Then
        
        txtPN.text = txtPN.text & ","
        txtPN.SelStart = Len(txtPN.text)
        If gbintActiveModule_PNRecipe = 1 Then
            Call frmPNRecipe.LoadPNRecipe
        End If
'        gbblnPNLoad = True
'        gbstrPNRecipeFile = "D:\Project\RTP\RTP-125-ZK\Src\Recipe\0-130722.rcp"
'        frmRecipeEdit.cmdRecipeOpen_Click
'        gbblnPNLoad = False
'        gbstrPN = txtPN.Text
'        txtBN.SelStart = 0
'        txtBN.SelLength = 100
'        txtBN.SetFocus
    End If
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
