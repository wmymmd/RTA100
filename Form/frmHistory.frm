VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmHistory 
   Caption         =   "Access"
   ClientHeight    =   9780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "標楷體"
      Size            =   12
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Access"
   MDIChild        =   -1  'True
   ScaleHeight     =   9780
   ScaleWidth      =   14595
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   9495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14325
      _ExtentX        =   25268
      _ExtentY        =   16748
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   617
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Message"
      TabPicture(0)   =   "frmHistory.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbLastRun"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dtpSearchEndDate"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cdFile"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "dtpSearchDate"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "adoHistory"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "dgdHistoryAlert"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdSearch"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chkDate"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chkType"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmbType"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdMoveFirst"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdMoveLast"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdMovePrior"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdMoveNext"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdExportCSV"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Report"
      TabPicture(1)   =   "frmHistory.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dtpCurrReport"
      Tab(1).Control(1)=   "sgdReport"
      Tab(1).Control(2)=   "Label1"
      Tab(1).ControlCount=   3
      Begin VB.CommandButton cmdExportCSV 
         Caption         =   "導出CSV"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10680
         TabIndex        =   18
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdMoveNext 
         Caption         =   "下筆"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   13080
         TabIndex        =   15
         Top             =   4920
         Width           =   1095
      End
      Begin VB.CommandButton cmdMovePrior 
         Caption         =   "上筆"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   13080
         TabIndex        =   14
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmdMoveLast 
         Caption         =   "最後頁"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   13080
         TabIndex        =   13
         Top             =   6840
         Width           =   1095
      End
      Begin VB.CommandButton cmdMoveFirst 
         Caption         =   "最前頁"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   13080
         TabIndex        =   12
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox cmbType 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmHistory.frx":0038
         Left            =   6960
         List            =   "frmHistory.frx":003A
         TabIndex        =   10
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox chkType 
         Caption         =   "Check1"
         Height          =   255
         Left            =   6120
         TabIndex        =   9
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox chkDate 
         Caption         =   "Check1"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "搜尋"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8880
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpCurrReport 
         Height          =   375
         Left            =   -73920
         TabIndex        =   3
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   221839361
         CurrentDate     =   40030
      End
      Begin MSFlexGridLib.MSFlexGrid sgdReport 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   2
         Top             =   960
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   6376
         _Version        =   393216
         Rows            =   6
         Cols            =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dgdHistoryAlert 
         Height          =   7695
         Left            =   240
         TabIndex        =   1
         Top             =   1080
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   13573
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   21
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "Date"
            Caption         =   "Date"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Time"
            Caption         =   "Time"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Type"
            Caption         =   "Type"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Code"
            Caption         =   "Code"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Description"
            Caption         =   "Description"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            ScrollBars      =   2
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column02 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   1665.071
            EndProperty
            BeginProperty Column03 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
               ColumnWidth     =   1289.764
            EndProperty
            BeginProperty Column04 
               ColumnAllowSizing=   0   'False
               Locked          =   -1  'True
            EndProperty
            BeginProperty Column05 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   6240.189
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc adoHistory 
         Height          =   7650
         Left            =   12600
         Top             =   1080
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   13494
         ConnectMode     =   3
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   1
         Enabled         =   -1
         Connect         =   $"frmHistory.frx":003C
         OLEDBString     =   $"frmHistory.frx":00F0
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Alert"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtpSearchDate 
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   221839361
         CurrentDate     =   40030
      End
      Begin MSComDlg.CommonDialog cdFile 
         Left            =   12360
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker dtpSearchEndDate 
         Height          =   375
         Left            =   4320
         TabIndex        =   20
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   221839361
         CurrentDate     =   40030
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "結束時間:"
         Height          =   375
         Left            =   3240
         TabIndex        =   19
         Top             =   620
         Width           =   1095
      End
      Begin VB.Label lbLastRun 
         Alignment       =   2  'Center
         Caption         =   "NA"
         Height          =   375
         Left            =   2040
         TabIndex        =   17
         Top             =   9000
         Width           =   10455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "前次執行記錄:"
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   9000
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "類別:"
         Height          =   375
         Left            =   6360
         TabIndex        =   11
         Top             =   620
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "開始時間:"
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   620
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "日期:"
         Height          =   375
         Left            =   -74880
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private ReportViewIni As New cInifile

Public Sub AppendLogAlert(lngNo As Long, strType As String, lngCode As Long, strDesc As String, intLogType As Integer)
    adoHistory.Recordset.AddNew
    adoHistory.Recordset(0) = lngNo
    adoHistory.Recordset(1) = Date
    adoHistory.Recordset(2) = Time
    adoHistory.Recordset(3) = strType   '類別
    adoHistory.Recordset(4) = lngCode   '訊息碼
    adoHistory.Recordset(5) = strDesc   '內容
    
    adoHistory.Recordset.Update
    
    's As String
    's = "insert into log ([Name],[Level],[Description],[Time]) values ('"+Name+"','"+Level+"','"+Desc+"','"+Date+
    'adoHistory.RecordSource = "insert into log ([Name],[Level],[Description],[Time]) values ('jjj','2','edccdsc','1992/11/23 PM 3:05:54')"
    'adoHistory.Refresh
End Sub

'Public Sub AppendLogOperation(lngNo As Long, strType As String, lngCode As Long, strDesc As String, intLogType As Integer)
'    adoHistoryOp.Recordset.AddNew
'    adoHistoryOp.Recordset(0) = lngNo 'Name
'    adoHistoryOp.Recordset(1) = Date  'Level
'    adoHistoryOp.Recordset(2) = Time 'Desc
'    adoHistoryOp.Recordset(3) = strType 'Now
'    adoHistoryOp.Recordset(4) = lngCode 'LogType
'    adoHistoryOp.Recordset(5) = strDesc 'LogType
'
'    adoHistoryOp.Recordset.Update
'    's As String
'    's = "insert into log ([Name],[Level],[Description],[Time]) values ('"+Name+"','"+Level+"','"+Desc+"','"+Date+
'    'adoHistory.RecordSource = "insert into log ([Name],[Level],[Description],[Time]) values ('jjj','2','edccdsc','1992/11/23 PM 3:05:54')"
'    'adoHistory.Refresh
'End Sub

Public Sub SearchLog(dt As String, SearchType As Integer, LogType As Integer, OrderType As Integer, Desc As Boolean)
    'SearchType =   0   Show All Data
    '               1   Show range of Data
    '               2   Show range of Data & LogType
    Dim S As String
    Dim d1 As String
    Dim d2 As String
    Dim d3 As String
        
    d1 = dt + " 00:00:00"
    d2 = dt + " 23:59:59"
        
    d3 = ""
    Select Case OrderType
        Case 0
            d3 = " order by  [Time]"
        Case 1
            d3 = " order by  [Name]"
        Case 2
            d3 = " order by  [Level]"
        Case 3
            d3 = " order by  [Description]"
    End Select
    
    If Desc = True Then
        d3 = d3 + " desc"
    End If
    
    Select Case SearchType
    Case 0
        S = "select  * from [Log]" + d3
    Case 1
        S = "select  * from [Log] where Time between #" + d1 + "# and #" + d2 + "#" + d3
    Case 2
        S = "select  * from [Log] where Type=" + str(LogType) + " and Time between #" + d1 + "# and #" + d2 + "#" + d3
    Case 3
        S = "select  * from [Log] where Type=" + str(LogType) + d3
    End Select
    
    adoHistory.RecordSource = S
    adoHistory.Refresh
End Sub

Public Sub SearchLogDT(dt As String, EndDt As String, LogType As String, SearchType As Integer)
    Dim S As String
    Dim d1 As String
    Dim d2 As String
    Dim d3 As String
        
    
    'S = "select  * from log where date=#" & dt & "#" & d3
    
    d3 = " order by  date,time"


    Select Case SearchType
    Case 0
        S = "select  * from [Log]" + d3
    Case 1
'        S = "select  * from [Log] where date =#" + dt + "#" + d3
         If dt <> EndDt Then
          S = "select  * from [Log] where date Between #" + dt + "# and #" + EndDt + "#" + d3
         Else
          S = "select  * from [Log] where date =#" + dt + "#" + d3
         End If
      
    Case 2
'        S = "select  * from [Log] where date =#" + dt + "# and Type='" + LogType + "'" + d3
         If dt <> EndDt Then
          S = "select  * from [Log] where date Between #" + dt + "# and #" + EndDt + "# and Type='" + LogType + "'" + d3
         Else
          S = "select  * from [Log] where date =#" + dt + "# and Type='" + LogType + "'" + d3
         End If
    Case 3
        S = "select  * from [Log] where Type='" + LogType + "'" + d3
    End Select
    
    adoHistory.RecordSource = S
    adoHistory.Refresh
    
End Sub

Public Sub DeleteLogAlert(dtFrom As String, dtTo As String)
    Dim S As String
    Dim d1 As String
    Dim d2 As String
    
    d1 = dtFrom + " 00:00:00"
    d2 = dtTo + " 23:59:59"
    S = "delete  * from [Log] where Time between #" + d1 + "# and #" + d2 + "#"
    
    On Error Resume Next
    adoHistory.RecordSource = S
    adoHistory.Refresh
    
    S = "select  * from [Log]"
    adoHistory.RecordSource = S
    adoHistory.Refresh
    

End Sub
'
'Public Sub DeleteLogOp(dtFrom As String, dtTo As String)
'    Dim s As String
'    Dim d1 As String
'    Dim d2 As String
'
'    d1 = dtFrom + " 00:00:00"
'    d2 = dtTo + " 23:59:59"
'    s = "delete  * from [Operation] where Time between #" + d1 + "# and #" + d2 + "#"
'
'    On Error Resume Next
'    adoHistoryOp.RecordSource = s
'    adoHistoryOp.Refresh
'
'    s = "select  * from [Operation]"
'    adoHistoryOp.RecordSource = s
'    adoHistoryOp.Refresh
'
'
'End Sub

Public Sub AppendReport(starttime As String, status As String, recipe As String, en As String, bn As String, pn As String)
    Dim S As String
    Dim i As Integer
    Dim strFilePath As String
    Dim strDir As String
    
    strFilePath = gbSystemPath & "\Report"
    strDir = dir(strFilePath, vbDirectory)
    If strDir = "" Then MkDir strFilePath
    strFilePath = strFilePath & "\" & Year(Date)
    strDir = dir(strFilePath, vbDirectory)
    If strDir = "" Then MkDir strFilePath
    strFilePath = strFilePath & "\" & Month(Date)
    strDir = dir(strFilePath, vbDirectory)
    If strDir = "" Then MkDir strFilePath
    
    strFilePath = gbSystemPath & "\Report\" & Year(Date) & "\" & Month(Date) & "\" & Day(Date) & ".txt"
    cReportIni.Path = strFilePath
    cReportIni.Section = "Report"
    
    
    S = starttime & "," & Format(Time, "hh:mm:ss") & "," & status & "," & recipe & "," & UCase(en) & "," & UCase(bn) & "," & UCase(pn)
    cReportIni.Key = "CurrReportIndex"
    i = Val(cReportIni.value) + 1
    cReportIni.value = i
    cReportIni.Key = CStr(i)
    cReportIni.value = S

End Sub

Private Sub cmdExportCSV_Click()
    Dim StrFileName As String
    Dim strNewFileName As String
    Dim TextLine As String
    Dim NewLine As String
    
    If adoHistory.Recordset.EOF Then GoTo ERRHNADLE
    
    gbblnNoModalForm = True
    On Error GoTo ERRHNADLE
    Call frmConfiguration.StopWatchDog
    cdFile.InitDir = gbSystemPath & "\Log"
    cdFile.Filter = "*.csv|*.csv"
    cdFile.FilterIndex = 1
    cdFile.CancelError = True
    cdFile.ShowOpen
    gbblnNoModalForm = False
    
    If cdFile.fileName <> "" Then
        StrFileName = cdFile.fileName
        adoHistory.Recordset.MoveFirst
        Open StrFileName For Output As #1
        Print #1, "Date,Time,Type,Code,Description"
        Do While Not adoHistory.Recordset.EOF
            NewLine = CStr(adoHistory.Recordset(1)) + "," + CStr(adoHistory.Recordset(2)) + "," + CStr(adoHistory.Recordset(3)) + "," + CStr(adoHistory.Recordset(4)) + "," + CStr(adoHistory.Recordset(5))
            Print #1, NewLine
            adoHistory.Recordset.MoveNext
        Loop
        Close #1
    End If
    frmConfiguration.StartWatchDog
    Exit Sub
ERRHNADLE:
    Close #1
    frmConfiguration.StartWatchDog
    ShowMessageOK "導出CSV失敗!"
End Sub

Private Sub cmdMoveFirst_Click()
    If adoHistory.Recordset.EOF = False Then adoHistory.Recordset.MoveFirst
End Sub

Private Sub cmdMoveLast_Click()
    If adoHistory.Recordset.EOF = False Then adoHistory.Recordset.MoveLast
End Sub

Private Sub cmdMoveNext_Click()
    If adoHistory.Recordset.EOF = False Then adoHistory.Recordset.MoveNext
    
End Sub

Private Sub cmdMovePrior_Click()
    If adoHistory.Recordset.BOF = False Then adoHistory.Recordset.MovePrevious
End Sub

Private Sub cmdSearch_Click()
    If chkDate.value = 1 Then
        If dtpSearchDate.value > dtpSearchEndDate.value Then
           MsgBox "開始時間不得大於結束時間!!!!"
           Exit Sub
        End If
        If chkType.value = 1 Then
            SearchLogDT CStr(dtpSearchDate), CStr(dtpSearchEndDate), cmbType.text, 2
        Else
            SearchLogDT CStr(dtpSearchDate), CStr(dtpSearchEndDate), cmbType.text, 1
        End If
    Else
        If chkType.value = 1 Then
            SearchLogDT CStr(dtpSearchDate), CStr(dtpSearchEndDate), cmbType.text, 3
        Else
            SearchLogDT CStr(dtpSearchDate), CStr(dtpSearchEndDate), cmbType.text, 0
        End If
    End If
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub dtpCurrReport_Change()
    Dim S As String
    Dim Count As Integer
    Dim i As Integer
    Dim j As Integer
    Dim sValue() As String
    
    S = gbSystemPath & "\Report\" & CStr(dtpCurrReport.Year) & "\" & CStr(dtpCurrReport.Month) & "\" & CStr(dtpCurrReport.Day) & ".txt"
    ReportViewIni.Path = S
    ReportViewIni.Section = "Report"
    ReportViewIni.Key = "CurrReportIndex"
    Count = Val(ReportViewIni.value)
    If Count = 0 Then
        sgdReport.Rows = 2
       
        For i = 0 To 6
            sgdReport.TextMatrix(1, i) = ""
        Next i
    Else
        sgdReport.Rows = Count + 1
        For i = 1 To Count
            ReportViewIni.Key = CStr(i)
            sValue = Split(ReportViewIni.value, ",")
            S = ""
            sgdReport.TextMatrix(i, 0) = CStr(i)
            For j = 0 To UBound(sValue)
                If j < 6 Then
                    sgdReport.TextMatrix(i, j + 1) = sValue(j)
                Else
                    S = S & sValue(j) & ","
                End If
            Next j
            sgdReport.TextMatrix(i, 7) = S
        Next i
    End If
End Sub

Private Sub Form_Activate()
    Dim s1 As String
    Dim s2 As String
    Dim s3 As String
    Dim s4 As String
    
'    If gbintLoginRight < 3 Then
'        cmdDelete.Visible = True
'        dtpFrom.Visible = True
'        dtpTo.Visible = True
'    Else
'        cmdDelete.Visible = False
'        dtpFrom.Visible = False
'        dtpTo.Visible = False
'    End If
    dtpCurrReport_Change
    iniPara.Section = "Debug"
    iniPara.Key = "Start"
    s1 = iniPara.value
    iniPara.Key = "Idle"
    s2 = iniPara.value
    iniPara.Key = "RampUp"
    s3 = iniPara.value
    iniPara.Key = "Hold"
    s4 = iniPara.value
    lbLastRun.Caption = "Start=" & s1 & ",Idle=" & s2 & ",RampUp=" & s3 & ",Hold=" & s4
    
End Sub

Private Sub Form_DblClick()
    'Call DeleteLogAlert("2009/9/15", "2009/9/15")
End Sub

Private Sub Form_Load()
    If CheckDB = 0 Then
               
        dgdHistoryAlert.Columns(0).Caption = "日期"
        dgdHistoryAlert.Columns(1).Caption = "時間"
        dgdHistoryAlert.Columns(2).Caption = "類別"
        dgdHistoryAlert.Columns(3).Caption = "訊息碼"
        dgdHistoryAlert.Columns(4).Caption = "內容"
        
    
        
        dgdHistoryAlert.Columns(0).Width = 1500
        dgdHistoryAlert.Columns(1).Width = 2000
        dgdHistoryAlert.Columns(2).Width = 1000
        dgdHistoryAlert.Columns(3).Width = 1000
        dgdHistoryAlert.Columns(4).Width = dgdHistoryAlert.Width _
                                            - dgdHistoryAlert.Columns(0).Width _
                                            - dgdHistoryAlert.Columns(1).Width _
                                            - dgdHistoryAlert.Columns(2).Width _
                                            - dgdHistoryAlert.Columns(3).Width _
                                            - 800
    
        cmbType.AddItem "Alarm"
        cmbType.AddItem "Manual"
        cmbType.AddItem "Process"
        cmbType.AddItem "Check"
    Else
        ShowMessageOK "Log檔案開啟失敗"
    End If
    
    'dgdHistoryAlert.row = 10
    
    
    
    sgdReport.ColWidth(0) = 600    'Index
    sgdReport.ColWidth(1) = 1200    'Start time
    sgdReport.ColWidth(2) = 1200    'End time
    sgdReport.ColWidth(3) = 800     'Status
    sgdReport.ColWidth(4) = 2000    'Recipe Name
    sgdReport.ColWidth(5) = 2000    'EN
    sgdReport.ColWidth(6) = 2000    'BN
    sgdReport.ColWidth(7) = 12000    'PN
    sgdReport.TextMatrix(0, 0) = "Index"
    sgdReport.TextMatrix(0, 1) = "Start Time"
    sgdReport.TextMatrix(0, 2) = "End Time"
    sgdReport.TextMatrix(0, 3) = "Status"
    sgdReport.TextMatrix(0, 4) = "Recipe"
    sgdReport.TextMatrix(0, 5) = "EN"
    sgdReport.TextMatrix(0, 6) = "BN"
    sgdReport.TextMatrix(0, 7) = "PN"
    
    dtpCurrReport.value = Date
    dtpCurrReport_Change
    
    dtpSearchDate.value = Date
    
    dtpSearchEndDate.value = Date
    cmbType.text = "Alarm"
End Sub

Public Function CheckDB() As Integer
    Dim strDataSource As String
    Dim strProvider As String
    Dim strPassowrd As String
    Dim strConectionString As String
    Dim strDBFile As String
    
    On Error GoTo Err0
    
    strDataSource = gbSystemPath & "\System\Log.mdb"
    strDBFile = dir(strDataSource)
    If strDBFile = "" Then
        MsgBox "Can't find the history db,vbokonly,Alert"
        CheckDB = -1
    End If
    strDataSource = "Data Source=" & strDataSource & ";"
    strProvider = "Provider=Microsoft.Jet.OLEDB.4.0;"
    strPassowrd = "Persist Security Info=False;Jet OLEDB:Database Password=gianttek"
    
    strConectionString = strProvider & strDataSource & strPassowrd
    
    adoHistory.ConnectionString = strConectionString
    adoHistory.RecordSource = "select  * from log order by date desc,time desc"
    adoHistory.Refresh
    
    Set dgdHistoryAlert.DataSource = adoHistory
    CheckDB = 0
    
    Exit Function
Err0:
    CheckDB = -9
End Function

