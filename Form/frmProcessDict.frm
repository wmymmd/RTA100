VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmProcessDict 
   Caption         =   "進程字典"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   7920
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox Combo_CustomerProcName 
      Height          =   300
      Left            =   480
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   240
      Width           =   1695
   End
   Begin VB.ComboBox Combo_SysProcName 
      Height          =   300
      Left            =   4320
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Btn_Save 
      Caption         =   "保存"
      Height          =   495
      Left            =   6480
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   7646
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "SimSun"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmProcessDict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit





Private Sub DataGrid1_Click()
If DataGrid1.ColSel = 1 Then
        With Combo_SysProcName
            .text = DataGrid1.textm(hfgRecipe.RowSel, GB_PROCESS_ACTION)
            strOriginActionName = hfgRecipe.TextMatrix(hfgRecipe.RowSel, GB_PROCESS_ACTION)
            .Move fraRecipe.Left + tabRecipe.Left + hfgRecipe.Left + hfgRecipe.ColPos(hfgRecipe.ColSel) + 20, _
                      fraRecipe.Top + tabRecipe.Top + hfgRecipe.Top + hfgRecipe.RowPos(hfgRecipe.RowSel) + 25, _
                  hfgRecipe.ColWidth(hfgRecipe.ColSel)
            .Visible = True
            .ZOrder
            .SetFocus
        End With
'    Else
'        With txtRecipeEdit
'            .text = hfgRecipe.TextMatrix(hfgRecipe.RowSel, hfgRecipe.ColSel)
'            .Move fraRecipe.Left + tabRecipe.Left + hfgRecipe.Left + hfgRecipe.ColPos(hfgRecipe.ColSel) + 20, _
'                  fraRecipe.Top + tabRecipe.Top + hfgRecipe.Top + hfgRecipe.RowPos(hfgRecipe.RowSel) + 25, _
'                  hfgRecipe.ColWidth(hfgRecipe.ColSel)
'            .Visible = True
'            .ZOrder
''            .SetFocus
'        End With
    End If
    intCurrRowSel = hfgRecipe.RowSel
    intCurrColSel = hfgRecipe.ColSel
End Sub

Private Sub Form_Load()
InitForm
End Sub


Private Sub InitForm()
With DataGrid1
.Columns(0).Caption = "SysProcName"
.Columns(0).Width = 3000
.Columns(1).Caption = "CustomerProcName"
.Columns(1).Width = 3000
.RowHeight(0) = Combo_SysProcName.Height
With Combo_SysProcName
       .AddItem GB_ACTION_IDLE
        '.AddItem GB_ACTION_PREHEAT
        .AddItem GB_ACTION_RAMPUP
        .AddItem GB_ACTION_HOLD
        .AddItem GB_ACTION_STOP
        '.AddItem GB_ACTION_PURGE
        .AddItem GB_ACTION_RAMPDOWN
        .AddItem GB_ACTION_IOCONTROL

End With
End With
End Sub
