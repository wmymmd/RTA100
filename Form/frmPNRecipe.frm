VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPNRecipe 
   Caption         =   "PN-Recipe"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10725
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   12
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   10725
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   615
      Left            =   2880
      TabIndex        =   6
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   615
      Left            =   4440
      TabIndex        =   5
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   615
      Left            =   6000
      TabIndex        =   4
      Top             =   5040
      Width           =   1455
   End
   Begin VB.TextBox txtPN 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   10335
   End
   Begin VB.ListBox lstPNRecipe 
      Height          =   3180
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   10335
   End
   Begin MSComDlg.CommonDialog cdFile 
      Left            =   120
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "PN (ex A1234,W5678,W[1000:2000](for continuous number))"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   10215
   End
   Begin VB.Label Label1 
      Caption         =   "Recipe List"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmPNRecipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strFileName         As String

Private CurrKey As String
Private CurrValue As String
Private IsChanged As Boolean

Public Function LoadPNRecipe() As Boolean
    Dim strTemp As String
    Dim RecipeName As String
    Dim s1 As String
    Dim s2 As String
    Dim s3 As String
    Dim s4 As String
    Dim s5 As String
    Dim v1 As Integer
    Dim v2 As Integer
    Dim v3 As Integer
    Dim v4 As Integer
    Dim v5 As Integer
    Dim ss1() As String     'Input Code List
    Dim ss2() As String     'Ini Value List
    Dim ss3() As String     'Ini Key list
    Dim i As Integer
    Dim j As Integer
    Dim K As Integer
    Dim IsFound As Boolean
    Dim iCount As Long
    Dim IsValue As Boolean
    
    
    On Error GoTo ERRHNADLE
    strTemp = UCase(frmPlotProcess.txtPN)
    cIni.Section = "PNRecipeList"
    ss1 = Split(strTemp, ",")
    cIni.EnumerateCurrentSection ss3, iCount
    
    'find and load recipe from the first code
    If UBound(ss1) >= 0 And UBound(ss3) >= 0 Then
        If Len(ss1(0)) > 10 Then    'at least 11 words
            IsValue = True
            s2 = Mid(ss1(0), 8, 4)
            For j = 1 To 4
                If Asc(Mid(s2, j, 1)) > 57 Then     'not 0~9
                    IsValue = False
                    Exit For
                End If
            Next j
            
            If IsValue = True Then
            
                For i = 1 To UBound(ss3)
                    cIni.Key = ss3(i)
                    ss2 = Split(cIni.Value, ",")
                    If UBound(ss2) >= 0 Then
                        IsFound = False
                                           
                        s1 = Mid(ss1(0), 7, 1)
                        v1 = Mid(ss1(0), 8, 4)
                        For j = 0 To UBound(ss2)
                            s2 = Mid(ss2(j), 1, 1)
                            s3 = Mid(ss2(j), 2, 1)
                            
                            If s1 = s2 Then
                                If s3 = "[" Then
                                    v4 = InStr(ss2(j), ":")
                                    v5 = InStr(ss2(j), "]")
                                    s4 = Mid(ss2(j), 3, v4 - 3)
                                    s5 = Mid(ss2(j), v4 + 1, 4)
                                    v4 = Int(s4)
                                    v5 = Int(s5)
                                    If v1 >= v4 And v1 <= v5 Then
                                        IsFound = True
                                        Exit For
                                    End If
                                Else
                                    s2 = Mid(ss1(0), 7, 5)
                                    If s2 = ss2(j) Then
                                        IsFound = True
                                        Exit For
                                    End If
                                    
                                End If
                            End If
                        
                        Next j
                        
                        RecipeName = ""
                        If IsFound Then
                            RecipeName = ss3(i)
                            Exit For
                        End If
                        
                    End If
                
                Next i
            Else
                IsFound = False
                For i = 1 To UBound(ss3)
                    cIni.Key = ss3(i)
                    ss2 = Split(cIni.Value, ",")
                    If UBound(ss2) >= 0 Then
                        
                        s1 = Mid(ss1(0), 7, 5)
                        For j = 0 To UBound(ss2)
                            If s1 = ss2(j) Then
                                IsFound = True
                                Exit For
                            End If
                                                       
                        Next j
                        
                        RecipeName = ""
                        If IsFound Then
                            RecipeName = ss3(i)
                            Exit For
                        End If
                        
                    End If
                Next i
                        
            End If
            If IsFound Then
                gbblnPNLoad = True
                gbstrPNRecipeFile = RecipeName
                frmRecipeEdit.cmdRecipeOpen_Click
                gbblnPNLoad = False
            Else
                strTemp = "Product No. is not defined"
                
                ShowAlarm strTemp
                Call frmHistory.AppendLogAlert(1, "Alarm", 3014, "Product No. is not defined", 1)
                LoadPNRecipe = False
                Exit Function
            End If
        End If
        
        If IsFound Then
            'check code list as same recipe
            cIni.Key = RecipeName
            ss2 = Split(cIni.Value, ",")
            
            If UBound(ss1) > 0 And UBound(ss2) >= 0 Then
                For i = 1 To UBound(ss1)
                    If Len(ss1(i)) > 10 Then
                        IsValue = True
                        s2 = Mid(ss1(i), 8, 4)
                        For j = 1 To 4
                            If Asc(Mid(s2, j, 1)) > 57 Then     'not 0~9
                                IsValue = False
                                Exit For
                            End If
                        Next j
                        
                        If IsValue = True Then
                            s1 = Mid(ss1(i), 7, 1)
                            v1 = Mid(ss1(i), 8, 4)
                            IsFound = False
                            For j = 0 To UBound(ss2)
                                s2 = Mid(ss2(j), 1, 1)
                                s3 = Mid(ss2(j), 2, 1)
                                
                                If s1 = s2 Then
                                    If s3 = "[" Then
                                        v4 = InStr(ss2(j), ":")
                                        v5 = InStr(ss2(j), "]")
                                        s4 = Mid(ss2(j), 3, v4 - 3)
                                        s5 = Mid(ss2(j), v4 + 1, 4)
                                        v4 = Int(s4)
                                        v5 = Int(s5)
                                        If v1 >= v4 And v1 <= v5 Then
                                            IsFound = True
                                        End If
                                        
                                    Else
                                        s2 = Mid(ss1(i), 7, 5)
                                                                            
                                        If s2 = ss2(j) Then
                                            IsFound = True
                                            Exit For
                                        End If
                                        
                                    End If
                                End If
                            Next j
                        Else
                            IsFound = False
                            s1 = Mid(ss1(i), 7, 5)
                            For j = 0 To UBound(ss2)
                                If s1 = ss2(j) Then
                                    IsFound = True
                                    Exit For
                                End If
                            Next j
                        End If
                        If IsFound = False Then
                            strTemp = "Product No =" & ss1(i) & " is not defined"
                            
                            ShowAlarm strTemp
                            Call frmHistory.AppendLogAlert(1, "Alarm", 3014, "Product No. is not defined", 1)
                            LoadPNRecipe = False
                            Exit Function
                        End If
                    ElseIf Len(ss1(i)) > 0 Then
                        
                        ShowAlarm "Error Product No. (< 11 words)"
                        Call frmHistory.AppendLogAlert(1, "Alarm", 3015, "Error Product No. (< 11 words)", 1)
                        LoadPNRecipe = False
                        Exit Function
                    End If
                Next i
            End If
        End If
    End If
    If IsFound = False Then
        strTemp = "Product No. is not defined"
        ShowAlarm strTemp
        Call frmHistory.AppendLogAlert(1, "Alarm", 3014, "Product No. is not defined", 1)
        LoadPNRecipe = False
        Exit Function
    End If
        
    LoadPNRecipe = True
    Exit Function
    
    
                            
    
ERRHNADLE:
    ShowMessageOK "檔案開啟失敗"
    Call frmHistory.AppendLogAlert(1, "Alarm", 3016, "PN Define error", 1)
    LoadPNRecipe = False
End Function

Private Sub cmdDelete_Click()
    
    QuestionAns = vbNo
    ShowMessageYN "確定要刪除?"
    If QuestionAns = vbYes Then
        cIni.Key = CurrKey
        cIni.DeleteKey
        lstPNRecipe.RemoveItem (lstPNRecipe.ListIndex)
        If lstPNRecipe.ListCount > 0 Then
            lstPNRecipe.ListIndex = 0
            lstPNRecipe_Click
        End If
        
    End If
'    Dim lngRet                As Long
'    Dim i As Integer
'    Dim s As String
'    Dim ss As String * 1000
'    Dim sItem() As String
'
'
'    s = "PNRecipe" & CStr(lstPNRecipe.ListIndex)
'    ss = vbNullString
'    lngRet = WritePrivateProfileString("PNRecipeList", s, ss, strFileName)
'    lstPNRecipe.RemoveItem (lstPNRecipe.ListIndex)
End Sub

Private Sub cmdNew_Click()
    Dim S As String
    Dim i As Integer
    Dim b As Boolean
    
    cdFile.InitDir = gbSystemPath & "\Recipe"
    cdFile.Filter = "*.rcp|*.rcp"
    cdFile.FilterIndex = 1
    cdFile.ShowOpen
    If cdFile.FileName <> "" Then
        cIni.Key = cdFile.FileName
        
        
        If cIni.Value <> "" Then
            S = cdFile.FileName & " is exists."
            MsgBox S
        Else
            b = False
            For i = 0 To lstPNRecipe.ListCount - 1
                If cdFile.FileName = lstPNRecipe.List(i) Then
                    b = True
                    Exit For
                End If
            Next i
            If b Then
                lstPNRecipe.Selected(lstPNRecipe.ListCount - 1) = True
            Else
                CurrKey = cdFile.FileName
                CurrValue = ""
                txtPN.Text = ""
                lstPNRecipe.AddItem (cdFile.FileName)
                lstPNRecipe.Selected(lstPNRecipe.ListCount - 1) = True
                txtPN.SetFocus
            End If
        End If
        IsChanged = True
        
    End If
    
End Sub

Private Sub cmdSave_Click()
'    Dim lngRet                As Long
'    Dim i As Integer
'    Dim s As String
'    Dim ss As String * 1000
'    Dim sItem() As String
'
'
'    s = "PNRecipe" & CStr(lstPNRecipe.ListIndex)
'    ss = lstPNRecipe.List(lstPNRecipe.ListIndex) & ";" & UCase(txtPN.Text)
'    lngRet = WritePrivateProfileString("PNRecipeList", s, ss, strFileName)
    cIni.Key = CurrKey
    cIni.Value = UCase(txtPN.Text)
    CurrValue = txtPN.Text
    lstPNRecipe.List(lstPNRecipe.ListIndex) = CurrKey & "=" & UCase(CurrValue)
    
    lstPNRecipe_Click
    IsChanged = False
End Sub

Private Sub Form_Load()
    Dim S As String
    Dim sKey() As String, iCount As Long, i As Long
    
    strFileName = gbSystemPath & "\System\system.cfg"
    
    cIni.Path = strFileName
    cIni.Section = "PNRecipeList"
    cIni.EnumerateCurrentSection sKey(), iCount
    cIni.Default = ""
    
    If (iCount > 0) Then
        For i = 1 To iCount
            cIni.Key = sKey(i)
            S = sKey(i) & "=" & cIni.Value
            lstPNRecipe.AddItem (S)
        Next i
    End If
    
'    Dim sOut As String
'    With m_cIni
'        .Path = txtInfo(0)
'        .Section = txtInfo(1)
'        .EnumerateCurrentSection sKey(), iCount
'        If (iCount > 0) Then
'            For i = 1 To iCount
'                sOut = sOut & vbCrLf & "    " & sKey(i)
'            Next i
'            MsgBox "Section contains:" & sOut, vbInformation
'        Else
'            MsgBox "Section is empty.", vbInformation
'        End If
'    End With
    
    
'    For i = 0 To 100
'        s = "PNRecipe" & CStr(i)
'        lngRet = GetPrivateProfileString("PNRecipeList", s, "", ss, 1000, strFileName)
'
'        If lngRet > 0 Then
'            sItem = Split(ss, ";")
'            lstPNRecipe.AddItem (sItem(0))
'
'
'        Else
'            Exit For
'        End If
'    Next i
       
    
    If lstPNRecipe.ListCount > 0 Then
        lstPNRecipe.Selected(lstPNRecipe.ListCount - 1) = True
    End If
    
    IsChanged = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If IsChanged Then
        QuestionAns = vbNo
        ShowMessageYN "是否要儲存?"
        If QuestionAns = vbYes Then
            cmdSave_Click
        End If
    End If
End Sub

Private Sub lstPNRecipe_Click()
    Dim S As String
    Dim sKey() As String, iCount As Long, i As Long
        
    txtPN.Text = ""
    S = lstPNRecipe.List(lstPNRecipe.ListIndex)
    sKey = Split(S, "=")
    
    If UBound(sKey) > 0 Then
        cIni.Key = sKey(0)
        txtPN.Text = cIni.Value
        CurrKey = sKey(0)
        CurrValue = cIni.Value
    End If
'    s = "PNRecipe" & CStr(lstPNRecipe.ListIndex)
'    lngRet = GetPrivateProfileString("PNRecipeList", s, "", ss, 1000, strFileName)
'    If lngRet > 0 Then
'        sItem = Split(ss, ";")
'        txtPN.Text = ""
'        If UBound(sItem) > 0 Then
'            txtPN.Text = sItem(1)
'        End If
'    End If
End Sub

Private Sub txtPN_Change()
    IsChanged = True
End Sub
