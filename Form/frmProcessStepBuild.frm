VERSION 5.00
Begin VB.Form frmProcessStepBuild 
   Caption         =   "Process Step自定義"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6270
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6270
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox TxtGas 
      Height          =   375
      Index           =   6
      Left            =   4080
      TabIndex        =   26
      Top             =   2880
      Width           =   1815
   End
   Begin VB.ComboBox Combo_SysAction 
      Height          =   300
      Left            =   1080
      TabIndex        =   24
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton CmdBtn_Query 
      Caption         =   "查詢"
      Height          =   495
      Left            =   2760
      TabIndex        =   22
      Top             =   120
      Width           =   1015
   End
   Begin VB.TextBox TxtGas 
      Height          =   375
      Index           =   5
      Left            =   1080
      TabIndex        =   21
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox TxtGas 
      Height          =   375
      Index           =   4
      Left            =   4080
      TabIndex        =   19
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox TxtGas 
      Height          =   375
      Index           =   3
      Left            =   1080
      TabIndex        =   17
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox TxtGas 
      Height          =   375
      Index           =   2
      Left            =   4080
      TabIndex        =   15
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox TxtGas 
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   13
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox TxtGas 
      Height          =   375
      Index           =   0
      Left            =   4080
      TabIndex        =   11
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Text_Time 
      Height          =   375
      Left            =   1080
      TabIndex        =   9
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Text_TP 
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton CmdBtn_Add 
      Caption         =   "+保存"
      Height          =   495
      Left            =   5040
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox Combo_Action 
      Height          =   300
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox Text_Action 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton CmdBtn_Delete 
      Caption         =   "-作廢"
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lbGas 
      Caption         =   "NA"
      Height          =   255
      Index           =   6
      Left            =   3480
      TabIndex        =   25
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Lbl_SysAction 
      Caption         =   "SysAction:"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3480
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6240
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lbGas 
      Caption         =   "NA"
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   20
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label lbGas 
      Caption         =   "NA"
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   18
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lbGas 
      Caption         =   "NA"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   16
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lbGas 
      Caption         =   "NA"
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   14
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lbGas 
      Caption         =   "NA"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   12
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lbGas 
      Caption         =   "NA"
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   10
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label_Time 
      Caption         =   "Time:"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label_TP 
      Caption         =   "Tp:"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label_Action 
      Caption         =   "Action:"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label_Selection 
      Caption         =   "Action:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmProcessStepBuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TxtRow As Long
'Const ProcessStep_Path As String = App.Path + "\Config\ProcessStep.txt"
'Dim RowNum As Long
Dim TxtNull As Boolean












Private Sub CmdBtn_Add_Click()
Dim SumString As String
Dim i As Integer
If Text_Action.text <> "" Then
  If ActionIsExist(Text_Action.text) = False Then
  SumString = Text_Action.text + ";"
  If Text_TP.text <> "" Then
  SumString = SumString + Text_TP.text + ";"
  Else
  SumString = SumString + "0;"
  End If
  If Text_Time.text <> "" Then
  SumString = SumString + Text_Time.text + ";"
  Else
  SumString = SumString + "0;"
  End If
  For i = 1 To frmRecipeEdit.GasNames.Count
    If lbGas(i - 1).Caption <> "NA" Then
    If TxtGas(i - 1).text <> "" Then
    SumString = SumString + TxtGas(i - 1).text + ";"
    Else
    SumString = SumString + "0;"
    End If
    End If
    Next i
  If Combo_SysAction.text <> "" Then
     Call WriteLineToFile(App.Path + ProcessStep_Path, SumString)
     Call InitForm
    Call WritePrivateProfileString("Procdict", Text_Action.text, Combo_SysAction.text, App.Path + ProcDict_Path)
     MsgBox "Action:" + Text_Action.text + "保存成功！"
     Else
      MsgBox "請選擇sysAction值！"
  End If

  Else
  Select Case MsgBox("是否要更新此Action", vbYesNo + vbQuestion, "確認")
        Case vbYes
        SumString = Text_Action.text + ";"
        If Text_TP.text <> "" Then
        SumString = SumString + Text_TP.text + ";"
        Else
        SumString = SumString + "0;"
        End If
        If Text_Time.text <> "" Then
        SumString = SumString + Text_Time.text + ";"
        Else
        SumString = SumString + "0;"
        End If
        For i = 1 To frmRecipeEdit.GasNames.Count
        If lbGas(i - 1).Caption <> "NA" Then
        If TxtGas(i - 1).text <> "" Then
        SumString = SumString + TxtGas(i - 1).text + ";"
        Else
        SumString = SumString + "0;"
        End If
        End If
        Next i
        If Combo_SysAction.text <> "" Then
        Call ModifyTextFile(App.Path + ProcessStep_Path, TxtRow, SumString)
        Call WritePrivateProfileString("Procdict", Text_Action.text, Combo_SysAction.text, App.Path + ProcDict_Path)
        MsgBox "Action:" + Text_Action.text + "修改成功！"
        Else
        MsgBox "請選擇sysAction值！"
        End If
    End Select
  End If
Else
MsgBox "請輸入Action值！"
End If
End Sub

Private Function ActionIsExist(Action As String) As Boolean
Dim result As Boolean
Dim TxtContent As Collection
Dim RetrunCollection As Collection
Dim TxtLine() As String
Dim i As Integer
result = False
If Combo_Action.text <> "" Then
Set TxtContent = ReadTextFileToArray(App.Path + ProcessStep_Path)
If TxtContent.Count > 0 Then
TxtNull = False
For i = 1 To TxtContent.Count
TxtLine = Split(TxtContent(i), ";")
If TxtLine(0) = Action Then
result = True
TxtRow = i
Exit For
End If
Next i
End If
End If
ActionIsExist = result
End Function

Private Sub RemoveItemFromCombo(itemToRemove As String)
    Dim i As Integer
    For i = 0 To Combo_Action.ListCount - 1
        If Combo_Action.List(i) = itemToRemove Then
            Combo_Action.RemoveItem i
            Exit For
        End If
    Next i
End Sub

Private Sub CmdBtn_Delete_Click()
If Combo_Action.text <> "" Then
If ActionIsExist(Combo_Action.text) = True Then
Call ModifyTextFile(App.Path + ProcessStep_Path, TxtRow, "")
Call RemoveItemFromCombo(Combo_Action.text)
'DeleteKeyFromIni App.Path + ProcDict_Path, "Procdict", Combo_Action.text
End If
Else
 MsgBox "請先下拉菜單選擇需要作廢的Action！"
End If
End Sub

Private Sub CmdBtn_Query_Click()
Dim TxtContent As Collection
Dim RetrunCollection As Collection
Dim TxtLine() As String
Dim i As Integer
Dim j As Integer

Dim buffer As String
Dim result As Long
buffer = String(255, 0)
If Combo_Action.text <> "" Then
Set TxtContent = ReadTextFileToArray(App.Path + ProcessStep_Path)
If TxtContent.Count > 0 Then
For i = 1 To TxtContent.Count
TxtLine = Split(TxtContent(i), ";")
If TxtLine(0) = Combo_Action.text Then
  Text_Action.text = TxtLine(0)
  Text_TP.text = TxtLine(1)
  Text_Time.text = TxtLine(2)
  For j = 3 To UBound(TxtLine) - 1
  TxtGas(j - 3).text = TxtLine(j)
  Next j
  result = GetPrivateProfileString("Procdict", Text_Action.text, "", buffer, Len(buffer), App.Path + ProcDict_Path)
  buffer = Left(buffer, result)
  Combo_SysAction.text = buffer
  Exit For
End If
Next i
End If
Else
MsgBox "請輸入查詢條件Action！"
End If

End Sub




Private Sub Form_Load()
InitForm

       With Combo_SysAction
       .AddItem GB_ACTION_IDLE
        '.AddItem GB_ACTION_PREHEAT
        .AddItem GB_ACTION_RAMPUP
        .AddItem GB_ACTION_HOLD
        .AddItem GB_ACTION_STOP
        '.AddItem GB_ACTION_PURGE
        .AddItem GB_ACTION_RAMPDOWN
'        .AddItem GB_ACTION_IOCONTROL

End With


End Sub
Private Function IsDuplicate(item As String) As Boolean
    Dim i As Integer
    For i = 0 To Combo_Action.ListCount - 1
        If Combo_Action.List(i) = item Then
            IsDuplicate = True
            Exit Function
        End If
    Next i
    IsDuplicate = False
End Function

Private Sub InitForm()
Dim TxtContent As Collection
Dim TxtLine() As String
Dim i As Integer
Dim s1 As String
Dim S As String
Set TxtContent = ReadTextFileToArray(App.Path + ProcessStep_Path)
If TxtContent.Count > 0 Then
For i = 1 To TxtContent.Count
Dim Count As Integer
TxtLine = Split(TxtContent(i), ";")
If IsDuplicate(TxtLine(0)) <> True Then
  Combo_Action.AddItem (TxtLine(0))
End If
Next i
End If
For i = 1 To frmRecipeEdit.GasNames.Count
S = frmRecipeEdit.GasNames(i) + ":"
lbGas(i - 1).Caption = S
Next i
End Sub

'Private Sub InitForm()
'    Dim i As Integer, j, k, L As Integer
'    Dim sngTotalGridWidth As Single
'    Dim S As String
'    Dim s1 As String
'    Dim TxtContent As Collection
'    Dim TxtLine() As String
'    Dim TxtNonGas() As String
'    Dim TxtGas() As String
'    Dim TxtCell As String
'    Set TxtContent = ReadTextFileToArray("D:\123.txt")
'         With HfgProcessStep
'        '.Top = 500
'        .Left = 200
'        .FixedCols = 1
'        .FixedRows = 1
'        .Rows = 51
'        .Cols = 10
'        .ColWidth(0) = 800
'        .ColWidth(1) = 1500
'        For i = 2 To .Cols - 1
'            .ColWidth(i) = 1500
'        Next i
'        For i = 0 To .Cols - 1
'            sngTotalGridWidth = sngTotalGridWidth + .ColWidth(i)
'            .ColAlignmentFixed = flexAlignCenterCenter
'            .ColAlignment(i) = flexAlignCenterCenter
'        Next i
'        .Width = sngTotalGridWidth + 350
'        .TextMatrix(0, GB_PROCESS_STEP) = "Step"
'        .TextMatrix(0, GB_PROCESS_ACTION) = "Action"
'        .TextMatrix(0, GB_PROCESS_TEMP) = "T(℃)/ P(%)"
'        .TextMatrix(0, GB_PROCESS_TIME) = "Time (Sec)"
'
'        For j = 0 To GB_GAS_MAX - 1
'            s1 = gbstrGasUnit(j)
'            If s1 = "SLM" Then s1 = "LM"
'            If s1 = "SCCM" Then s1 = "CM"
'            S = gbstrGasAlias(j) & "(" & s1 & "~" & CStr(gbsngMaxGasSLMP(j)) & ")"
'            .TextMatrix(0, GB_PROCESS_GAS1 + j) = S
'        Next j
'        .RowHeight(0) = CmdBtn_Delete.Height
'         If TxtContent.Count = 0 Then
'            For i = 1 To .Rows - 1
'                .RowHeight(i) = CmdBtn_Delete.Height
'                .TextMatrix(i, GB_PROCESS_STEP) = str(i)
'                .TextMatrix(i, GB_PROCESS_ACTION) = ""
'                .TextMatrix(i, GB_PROCESS_TEMP) = ""
'                .TextMatrix(i, GB_PROCESS_TIME) = ""
'                For j = 0 To GB_GAS_MAX - 1
'                    .TextMatrix(i, GB_PROCESS_GAS1 + j) = ""
'                Next j
'            Next i
'           Else
'            For k = 1 To TxtContent.Count
'             TxtLine = Split(TxtContent(k), "|")
'             TxtNonGas = Split(TxtLine(0), ";")
'                .TextMatrix(k, GB_PROCESS_STEP) = str(k)
'                .TextMatrix(k, GB_PROCESS_ACTION) = TxtNonGas(0)
'                .TextMatrix(k, GB_PROCESS_TEMP) = TxtNonGas(1)
'                .TextMatrix(k, GB_PROCESS_TIME) = TxtNonGas(2)
'              TxtGas = Split(TxtLine(1), ";")
'                For j = 0 To GB_GAS_MAX - 1
'                    .TextMatrix(k, GB_PROCESS_GAS1 + j) = TxtGas(j)
'                Next j
'            Next k
'            End If
'        .Refresh
'        .AllowUserResizing = flexResizeNone
'    End With
'End Sub





'
'Private Sub HfgProcessStep_Click()
'  If HfgProcessStep.ColSel = 1 Then
'        With cmbRecipeAction
'            .Text = HfgProcessStep.TextMatrix(hfgRecipe.RowSel, GB_PROCESS_ACTION)
'            strOriginActionName = HfgProcessStep.TextMatrix(hfgRecipe.RowSel, GB_PROCESS_ACTION)
'            .Move fraRecipe.Left + tabRecipe.Left + hfgRecipe.Left + hfgRecipe.ColPos(hfgRecipe.ColSel) + 20, _
'                      fraRecipe.Top + tabRecipe.Top + hfgRecipe.Top + hfgRecipe.RowPos(hfgRecipe.RowSel) + 25, _
'                  hfgRecipe.ColWidth(hfgRecipe.ColSel)
'            .Visible = True
'            .ZOrder
'            .SetFocus
'        End With
'    Else
'        With txtRecipeEdit
'            .Text = hfgRecipe.TextMatrix(hfgRecipe.RowSel, hfgRecipe.ColSel)
'            .Move fraRecipe.Left + tabRecipe.Left + hfgRecipe.Left + hfgRecipe.ColPos(hfgRecipe.ColSel) + 20, _
'                  fraRecipe.Top + tabRecipe.Top + hfgRecipe.Top + hfgRecipe.RowPos(hfgRecipe.RowSel) + 25, _
'                  hfgRecipe.ColWidth(hfgRecipe.ColSel)
'            .Visible = True
'            .ZOrder
'            .SetFocus
'        End With
'    End If
'    intCurrRowSel = hfgRecipe.RowSel
'    intCurrColSel = hfgRecipe.ColSel
'End Sub
Private Sub Form_Unload(Cancel As Integer)
Call frmRecipeEdit.ReFreshAction
End Sub


