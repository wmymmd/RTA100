VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmTcOffset 
   Caption         =   "溫度偏執量設定"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   10245
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CdRecipe 
      Left            =   120
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox TxtTcOffset 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Save 
      Caption         =   "Save"
      Height          =   855
      Left            =   8280
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame FraTcOffset 
      Caption         =   "TcOffSet"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   9975
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgTcOffset 
         Height          =   4215
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   7435
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "frmTcOffset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public intCurrRowSel As Integer
Public intCurrColSel As Integer
Public StrFileName As String

Private Sub InitForm()
Dim i As Integer
  With fgTcOffset
        .Cols = 8
        .Rows = 31
        .ColWidth(0) = 400
        .ColWidth(1) = 4800
        .ColWidth(2) = 800
        .ColWidth(3) = 800
        .ColWidth(4) = 800
        .ColWidth(5) = 800
        .ColWidth(6) = 800
        .ColWidth(7) = 800
        
        .TextMatrix(0, 0) = "No"
        .TextMatrix(0, 1) = "RecipeName"
        .TextMatrix(0, 2) = "Period"
        .TextMatrix(0, 3) = "TC1"
        .TextMatrix(0, 4) = "TC2"
        .TextMatrix(0, 5) = "TC3"
        .TextMatrix(0, 6) = "TC4"
        .TextMatrix(0, 7) = "TC5"
 
        For i = 1 To fgTcOffset.Rows - 1
            .RowHeight(i) = 360
            .TextMatrix(i, 0) = CStr(i)
            .TextMatrix(i, 1) = "NA"
            .TextMatrix(i, 2) = "0"
            .TextMatrix(i, 3) = "0"
            .TextMatrix(i, 4) = "0"
            .TextMatrix(i, 5) = "0"
            .TextMatrix(i, 6) = "0"
            .TextMatrix(i, 7) = "0"
        Next i
    End With
End Sub

Private Sub Cmd_Save_Click()
Dim i As Integer
Dim lngRet As Long
For i = 1 To fgTcOffset.Rows - 1
  If fgTcOffset.TextMatrix(i, 1) <> "NA" Then
    lngRet = WritePrivateProfileString("TcOffset", "Set" + fgTcOffset.TextMatrix(i, 0) + "_RcpName", fgTcOffset.TextMatrix(i, 1), StrFileName)
    lngRet = WritePrivateProfileString("TcOffset", "Set" + fgTcOffset.TextMatrix(i, 0) + "_Period", fgTcOffset.TextMatrix(i, 2), StrFileName)
    lngRet = WritePrivateProfileString("TcOffset", "Set" + fgTcOffset.TextMatrix(i, 0) + "_TC1", fgTcOffset.TextMatrix(i, 3), StrFileName)
    lngRet = WritePrivateProfileString("TcOffset", "Set" + fgTcOffset.TextMatrix(i, 0) + "_TC2", fgTcOffset.TextMatrix(i, 4), StrFileName)
    lngRet = WritePrivateProfileString("TcOffset", "Set" + fgTcOffset.TextMatrix(i, 0) + "_TC3", fgTcOffset.TextMatrix(i, 5), StrFileName)
    lngRet = WritePrivateProfileString("TcOffset", "Set" + fgTcOffset.TextMatrix(i, 0) + "_TC4", fgTcOffset.TextMatrix(i, 6), StrFileName)
    lngRet = WritePrivateProfileString("TcOffset", "Set" + fgTcOffset.TextMatrix(i, 0) + "_TC5", fgTcOffset.TextMatrix(i, 7), StrFileName)
  End If
Next i
End Sub


Private Sub fgTcOffset_DblClick()
 fgTcOffset.Refresh
'  With TxtTcOffset
   If fgTcOffset.col = 1 Then
       CdRecipe.InitDir = gbSystemPath & "\Recipe"
       CdRecipe.Filter = "*.rcp|*.rcp"
       CdRecipe.FilterIndex = 1
       CdRecipe.CancelError = False
       CdRecipe.ShowOpen
       If CdRecipe.fileName <> "" Then
            With TxtTcOffset
            .text = CdRecipe.fileName
            .Move FraTcOffset.Left + fgTcOffset.Left + fgTcOffset.ColPos(fgTcOffset.ColSel), _
                  FraTcOffset.Top + fgTcOffset.Top + fgTcOffset.RowPos(fgTcOffset.RowSel), _
                  fgTcOffset.ColWidth(fgTcOffset.ColSel)
            .Visible = True
            .ZOrder
            .SetFocus
          End With
       End If
    Else
        With TxtTcOffset
           .text = fgTcOffset.TextMatrix(fgTcOffset.row, fgTcOffset.col)
            .Move FraTcOffset.Left + fgTcOffset.Left + fgTcOffset.ColPos(fgTcOffset.ColSel), _
                  FraTcOffset.Top + fgTcOffset.Top + fgTcOffset.RowPos(fgTcOffset.RowSel), _
                  fgTcOffset.ColWidth(fgTcOffset.ColSel)
            .Visible = True
            .ZOrder
            .SetFocus
          End With
      
    End If

'   End With
    intCurrRowSel = fgTcOffset.RowSel
    intCurrColSel = fgTcOffset.ColSel
End Sub

Private Sub Form_Load()
TxtTcOffset.Visible = False
StrFileName = App.Path + ProcDict_Path
InitForm
ReadData
End Sub

Private Sub ReadData()
Dim i As Integer
If SectionExistsInIni(StrFileName, "TcOffset") Then
  For i = 1 To 30
  If CommnonReadini("TcOffset", "Set" + CStr(i) + "_RcpName", StrFileName) <> "" Then
   fgTcOffset.TextMatrix(i, 1) = CommnonReadini("TcOffset", "Set" + CStr(i) + "_RcpName", StrFileName)
   fgTcOffset.TextMatrix(i, 2) = CommnonReadini("TcOffset", "Set" + CStr(i) + "_Period", StrFileName)
   fgTcOffset.TextMatrix(i, 3) = CommnonReadini("TcOffset", "Set" + CStr(i) + "_TC1", StrFileName)
   fgTcOffset.TextMatrix(i, 4) = CommnonReadini("TcOffset", "Set" + CStr(i) + "_TC2", StrFileName)
   fgTcOffset.TextMatrix(i, 5) = CommnonReadini("TcOffset", "Set" + CStr(i) + "_TC3", StrFileName)
   fgTcOffset.TextMatrix(i, 6) = CommnonReadini("TcOffset", "Set" + CStr(i) + "_TC4", StrFileName)
   fgTcOffset.TextMatrix(i, 7) = CommnonReadini("TcOffset", "Set" + CStr(i) + "_TC5", StrFileName)
  End If
  Next i
End If
End Sub


Public Function GetTcOffset(fileName As String) As Single()
Dim i As Integer
Dim FilePath As String
Dim TpOffset(6) As Single
Dim StrFileName As String
StrFileName = App.Path + ProcDict_Path
If SectionExistsInIni(StrFileName, "TcOffset") Then
  For i = 1 To 30
  FilePath = CommnonReadini("TcOffset", "Set" + CStr(i) + "_RcpName", StrFileName)
  If StrFileName <> "" And FilePath = fileName Then
   TpOffset(0) = CSng(CommnonReadini("TcOffset", "Set" + CStr(i) + "_TC1", StrFileName))
   TpOffset(1) = CSng(CommnonReadini("TcOffset", "Set" + CStr(i) + "_TC2", StrFileName))
   TpOffset(2) = CSng(CommnonReadini("TcOffset", "Set" + CStr(i) + "_TC3", StrFileName))
   TpOffset(3) = CSng(CommnonReadini("TcOffset", "Set" + CStr(i) + "_TC4", StrFileName))
   TpOffset(4) = CSng(CommnonReadini("TcOffset", "Set" + CStr(i) + "_TC5", StrFileName))
   TpOffset(5) = CSng(CommnonReadini("TcOffset", "Set" + CStr(i) + "_Period", StrFileName))
   Exit For
  End If
  Next i
  GetTcOffset = TpOffset
End If
End Function




Private Sub TxtTcOffset_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
     fgTcOffset.TextMatrix(fgTcOffset.RowSel, fgTcOffset.ColSel) = TxtTcOffset.text
     TxtTcOffset.text = ""
     TxtTcOffset.Visible = False
     Cmd_Save.SetFocus
 End If
End Sub
