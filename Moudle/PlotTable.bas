Attribute VB_Name = "mdlPlotTable"
 '========================================================================================================
'Copyright: Aries Liu
'Author: Aries Liu
'Date: 7/05/2006
'========================================================================================================
Option Explicit
    
'--------------------------------------------------------------------
'Define chart default layout contant
'--------------------------------------------------------------------
    Public Const CHART_DEF_TIME = 120  'Sec
    Public Const CHART_DEF_TEMP = 1200 'Degree
    Public Const CHART_DEF_PERCENT = 100 'Percent
'--------------------------------------------------------------------
    Public Const MAX_RECORD_DATA = 65535
    
    Dim i, j, K             As Integer          'For loop or count
    Public Table_Xsize      As Long          'PID Table X asix Scale
    Public Table_Ysize      As Long          'PID Table Y asix Scale
    Public Set_Temp         As Single           'Set Temperature of the table display
    Public Set_Time         As Single           'Set Time by second unit of the table display
    Public Set_Percent      As Long             'Set Percent of the table display
    Public tmpTable         As Long
    
    'This type for recode plot recipe point on the table
    Type PlotRecipePoint
        rcpX As Single
        rcpY As Single
    End Type
    
    'Just to let form moudles call Plot_Recipe by structure
    Type RecipeTable
        arrayRecipe(100, 20) As String
    End Type
    Dim RcpXY As RecipeTable
    
    Dim Check_Status As String      'Check the recipe status (idel or preheat or ramp,or stop)
    Dim m_SecPoint(100) As Single    'Because Process Step limit 20
    Dim m_TempPoint(100)  As Single  'declear Tempurature point max 20 point
    Dim m_Slope As Double           'ramp's slope
    Dim tmpX As Single              'recode X point by one's status
    Dim tmpY As Single
    Dim PointCounts As Long         'get  total point in the process one's cycle
    Dim CalCounts As Long           'Calculate count
    Dim m_PlotXY(100) As PlotRecipePoint   'get recipe from table
    Dim lngMaxTemp As Long
    Dim lngMaxTime As Long
    Dim bRet As Boolean
    
    
    Const PointsPerSec As Long = 60     'set adjust point number per second
    
    
    Public m_Recipe As RecipeTable
    Public gbLogRecipe As RecipeTable
    Public gbsngLogProcessRecord(65535, 50) As Single

'--------------------------------------------------------------------
'Define Process Table size and pitch
'--------------------------------------------------------------------
Public X_Step          As Long             'pitch
Public Y_Step          As Long
Public AssignX_Axis    As Long         'size
Public AssignY_Axis    As Long
Public Color_Choice    As Long   'Choice Color to plot

Public Type Draw_Table      'Define table spec.
     Xsize          As Long
     Ysize          As Long
     ExtendXY_Size  As Long
     Xscale_size    As Long
     Yscale_size    As Long
     XPitch         As Single
     YPitch         As Single
     row            As Long
     Column         As Long
End Type
   
Public ProcessTable    As Draw_Table

Public gbProcessLogTable    As Draw_Table
'--------------------------------------------------------------------
'This sub function for to plot recipe draw on PID table
'--------------------------------------------------------------------
Public Function PlotProcessTable(ByVal StartPic As Long, ByVal EndPic As Long, picObject() As PictureBox) As Boolean
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
    
    With ProcessTable
        .Xsize = 10000
        .Ysize = 8000
        .ExtendXY_Size = 500
        .row = 40
        .Column = 15
'        .Xscale_size = 900 + .ExtendXY_Size
'        .Yscale_size = 600 + .ExtendXY_Size
        AssignX_Axis = Set_Time * 60    '120 sec 120 * 60 = 7200
        AssignY_Axis = Set_Temp * 10    '1200 degree 1200*10= 12000
        X_Step = AssignX_Axis / (ProcessTable.Column + 1)
        Y_Step = AssignY_Axis / ProcessTable.row * 2
        .Xscale_size = AssignX_Axis + X_Step
        .Yscale_size = AssignY_Axis + Y_Step

    End With
    
    For PicNO = StartPic To EndPic
        'picObject(PicNO).AutoRedraw = True
        picObject(PicNO).FontSize = 12
        picObject(PicNO).Width = ProcessTable.Xsize + ProcessTable.ExtendXY_Size * 2
        picObject(PicNO).Height = ProcessTable.Ysize + ProcessTable.ExtendXY_Size * 2
        picObject(PicNO).Scale (-X_Step, ProcessTable.Yscale_size + Y_Step) _
                            -(AssignX_Axis + X_Step, -Y_Step)
'        fraProcessChart.Height = picObject(PicNO).Height + 500
'        fraProcessChart.Width = picObject(PicNO).Width + 500
    Next PicNO
    
    For PicNO = StartPic To EndPic
        picObject(PicNO).Cls
        picObject(PicNO).DrawWidth = 1
        Table_Xsize = ProcessTable.Xscale_size - X_Step 'Table's Weight
        Table_Ysize = ProcessTable.Yscale_size - Y_Step 'Table's Height
        XPitch = (Table_Xsize) / ProcessTable.Column
        YPitch = (Table_Ysize) / ProcessTable.row
        iCount = 0
        'This loop to plot line
        For m_Y = 0 To (Table_Ysize) Step YPitch
            For m_X = 0 To (Table_Xsize) Step XPitch
                picObject(PicNO).Line (m_X, 0)-(m_X, Table_Ysize)
            Next m_X
            
            If 0 = (iCount Mod 2) Then
                Color_Choice = vbBlack
            Else
                Color_Choice = GB_ColorLightBlue 'GB_ColorLightGray
            End If
            picObject(PicNO).Line (0, m_Y)-(Table_Xsize, m_Y), Color_Choice
            iCount = iCount + 1
        Next m_Y
    
       'Plot temperature,percent and sec. number
    
       PercentPitch = Val(Set_Percent) / (ProcessTable.row / 2)
       TempPitch = Val(Set_Temp) / (ProcessTable.row / 2)
       For i = 0 To 20
            'Percent number display
            PercentStep = i * PercentPitch
            picObject(PicNO).PSet (Table_Xsize + 10, (i * YPitch * 2) + _
                                    picObject(PicNO).ScaleY((picObject(PicNO).FontSize / 2 * Table_Ysize / picObject(PicNO).Height) _
                                         , vbPoints, vbTwips)) 'sign point and get the currentX,Y
            If i = 20 Then
                picObject(PicNO).Print str(PercentStep) & "%"
            Else
                picObject(PicNO).Print str(PercentStep)
            End If
            'Temperature number Display
            TempStep = i * TempPitch
            picObject(PicNO).CurrentX = -XPitch '-ScaleX(picObject(PicNO).FontSize, vbPoints, vbTwips) * 3.5
            picObject(PicNO).CurrentY = (i * YPitch * 2) + _
                                         picObject(PicNO).ScaleY((picObject(PicNO).FontSize / 2 * Table_Ysize / picObject(PicNO).Height) _
                                         , vbPoints, vbTwips)
             If i = 0 Then
                picObject(PicNO).Print str(TempStep) & " (C)"
            Else
                picObject(PicNO).Print str(TempStep)
            End If
        Next i
        
        'For time sec display
        TimePitch = Val(Set_Time) / ProcessTable.Column
        For i = 0 To 15
            TimeStep = i * TimePitch
            m_Min = Fix(TimeStep / 60)
            m_Sec = (TimeStep Mod 60)
            picObject(PicNO).CurrentX = (i * XPitch) - XPitch / 2 'ScaleY((picObject(PicNO).FontSize), vbPoints, vbTwips)
            picObject(PicNO).CurrentY = -Y_Step / 3 'table axis'
            picObject(PicNO).Print str(m_Min) & ":" & str(m_Sec)
        Next i
    Next PicNO

    PlotProcessTable = True
    
End Function
'--------------------------------------------------------------------
'This sub function for to plot recipe draw on PID table
'--------------------------------------------------------------------
Public Function Plot_Recipe(RecipeXY As RecipeTable, StartPic As Integer, EndPic As Integer, picObject() As PictureBox) As Boolean
      
'    Dim lngMaxTemp As Long
'    Dim lngMaxTime As Long
    Dim sngRetCurveData(5) As Single
    Dim sngCurveDistance As Single
    Dim intCurvePoint(20) As Integer
    Dim intCalCurveCount As Integer
    
    Dim sngDrawScale As Single
    Dim SysProName As String
    Dim HoldTimes As Integer

    On Error GoTo ERR_PLOT_RECIPE
    
    bRet = False

    picObject(0).ScaleMode = vbPixels
    Check_Status = ""
    Set_Time = CHART_DEF_TIME
    Set_Temp = CHART_DEF_TEMP
    CalCounts = 0
    m_PlotXY(0).rcpX = 0
    m_PlotXY(0).rcpY = 0
      
    Check_Status = Trim(CStr(RecipeXY.arrayRecipe(1, GB_PROCESS_ACTION)))
    SysProName = Readini(Check_Status)
    If SysProName <> "" Then
    Check_Status = SysProName
    End If
    If Check_Status = "STOP" Or Check_Status = "" Then GoTo ERR_PLOT_RECIPE
    
    With RecipeXY
    lngMaxTime = 0
    lngMaxTemp = 0
    For i = 1 To 50
        lngMaxTime = lngMaxTime + CLng(RecipeXY.arrayRecipe(i, GB_PROCESS_TIME))
        lngMaxTemp = IIf(.arrayRecipe(i, GB_PROCESS_TEMP) >= lngMaxTemp, .arrayRecipe(i, GB_PROCESS_TEMP), lngMaxTemp)
    Next i
      gbdblTotalProcessTime = lngMaxTime
    'Adjust PIDTable's temperature display
    'If (lngMaxTemp - Set_Temp - 100) <= 0 Then
    Set_Temp = (lngMaxTemp \ 100) * 100 + 300
        'bRet = frmPlotProcess.Plot_ProcessTable(0, 0)
    'End If
    
    'If (lngMaxTime - Set_Time - 15) <= 0 Then
        Set_Time = (lngMaxTime \ 15) * 15 + 15
        
        bRet = frmPlotProcess.Plot_ProcessTable(0, 0, picObject)
        'bRet = frmPlotProcess.Plot_ProcessTable(0, 0)
    'End If
    HoldTimes = 0
    For i = 1 To GB_MAX_STEP_PROCESS
        CalCounts = CalCounts + 1
        Check_Status = Trim(CStr(.arrayRecipe(i, GB_PROCESS_ACTION)))
        SysProName = Readini(Check_Status)
        If SysProName <> "" Then
        Check_Status = SysProName
        End If
        Select Case Check_Status
            Case "Idle"
                m_SecPoint(i) = .arrayRecipe(i, GB_PROCESS_TIME)
                m_TempPoint(i) = .arrayRecipe(i, GB_PROCESS_TEMP)
                If i = 1 Then
                    tmpX = m_PlotXY(CalCounts - 1).rcpX + ((m_SecPoint(i) * (Table_Xsize / Val(Set_Time))))
                    tmpY = 0 'm_TempPoint(i) * (Table_Ysize / Val(Set_Temp))
                    m_PlotXY(CalCounts).rcpX = tmpX
                    m_PlotXY(CalCounts).rcpY = tmpY
                Else
                    tmpX = m_PlotXY(CalCounts - 1).rcpX
                    tmpY = 0 'm_TempPoint(i) * (Table_Ysize / Val(Set_Temp))
                    m_PlotXY(CalCounts).rcpX = tmpX
                    m_PlotXY(CalCounts).rcpY = tmpY
                    CalCounts = CalCounts + 1
                    tmpX = m_PlotXY(CalCounts - 1).rcpX + ((m_SecPoint(i) * (Table_Xsize / Val(Set_Time))))
                    tmpY = 0 'm_TempPoint(i) * (Table_Ysize / Val(Set_Temp))
                    m_PlotXY(CalCounts).rcpX = tmpX
                    m_PlotXY(CalCounts).rcpY = tmpY
                End If
                
            Case "PreHeat"
                'CalCounts = CalCounts - 1
                m_SecPoint(i) = RecipeXY.arrayRecipe(i, GB_PROCESS_TIME)
                m_TempPoint(i) = RecipeXY.arrayRecipe(i, GB_PROCESS_TEMP)
                tmpX = m_PlotXY(CalCounts - 1).rcpX + ((m_SecPoint(i) * (Table_Xsize / Val(Set_Time))))
                tmpY = m_TempPoint(i) * (Table_Ysize / Val(Set_Temp))
                m_PlotXY(CalCounts).rcpX = tmpX
                m_PlotXY(CalCounts).rcpY = tmpY
            Case "Ramp up"
                m_SecPoint(i) = RecipeXY.arrayRecipe(i, GB_PROCESS_TIME)
                m_TempPoint(i) = RecipeXY.arrayRecipe(i, GB_PROCESS_TEMP)
'                If i > 0 Then
'                    m_Slope = (m_TempPoint(i) - m_TempPoint(i - 1)) / m_SecPoint(i)
'                Else
'                    m_Slope = (m_TempPoint(i) - 0) / m_SecPoint(i)
'                End If
                tmpX = m_PlotXY(CalCounts - 1).rcpX + m_SecPoint(i) * (Table_Xsize / Val(Set_Time))
                tmpY = m_TempPoint(i) * (Table_Ysize / Val(Set_Temp))
                m_PlotXY(CalCounts).rcpX = tmpX
                m_PlotXY(CalCounts).rcpY = tmpY
                
                                             
            Case "Hold"
                HoldTimes = HoldTimes + 1
                m_SecPoint(i) = RecipeXY.arrayRecipe(i, GB_PROCESS_TIME)
                If HoldTimes = Val(CommnonReadini("Special_Setting", "Hold_Times", App.Path + ProcDict_Path)) Then
                m_SecPoint(i) = RecipeXY.arrayRecipe(i, GB_PROCESS_TIME) + Val(CommnonReadini("Special_Setting", "Hold_Offset", App.Path + ProcDict_Path))
                End If
''               m_SecPoint(i) = RecipeXY.arrayRecipe(i, GB_PROCESS_TIME) + Val(CommnonReadini("Special_Setting", "Hold_Offset", App.Path + ProcDict_Path))
                m_TempPoint(i) = RecipeXY.arrayRecipe(i, GB_PROCESS_TEMP)
                tmpX = m_PlotXY(CalCounts - 1).rcpX + m_SecPoint(i) * (Table_Xsize / Val(Set_Time))
                tmpY = m_TempPoint(i) * (Table_Ysize / Val(Set_Temp))
                m_PlotXY(CalCounts).rcpX = tmpX
                m_PlotXY(CalCounts).rcpY = tmpY
                
            Case "Vent"
                m_SecPoint(i) = .arrayRecipe(i, GB_PROCESS_TIME)
                m_TempPoint(i) = .arrayRecipe(i, GB_PROCESS_TEMP)
                If i = 1 Then
                    tmpX = m_PlotXY(CalCounts - 1).rcpX + ((m_SecPoint(i) * (Table_Xsize / Val(Set_Time))))
                    tmpY = 10 'm_TempPoint(i) * (Table_Ysize / Val(Set_Temp))
                    m_PlotXY(CalCounts - 1).rcpY = tmpY
                    m_PlotXY(CalCounts).rcpX = tmpX
                    m_PlotXY(CalCounts).rcpY = tmpY
                Else
                    tmpX = m_PlotXY(CalCounts - 1).rcpX
                    tmpY = 10 'm_TempPoint(i) * (Table_Ysize / Val(Set_Temp))
                    m_PlotXY(CalCounts).rcpX = tmpX
                    m_PlotXY(CalCounts).rcpY = tmpY
                    CalCounts = CalCounts + 1
                    tmpX = m_PlotXY(CalCounts - 1).rcpX + ((m_SecPoint(i) * (Table_Xsize / Val(Set_Time))))
                    tmpY = 10 'm_TempPoint(i) * (Table_Ysize / Val(Set_Temp))
                    m_PlotXY(CalCounts).rcpX = tmpX
                    m_PlotXY(CalCounts).rcpY = tmpY
                End If
            
            Case "Purge"
                m_SecPoint(i) = .arrayRecipe(i, GB_PROCESS_TIME)
                m_TempPoint(i) = .arrayRecipe(i, GB_PROCESS_TEMP)
                If i = 1 Then
                    tmpX = m_PlotXY(CalCounts - 1).rcpX + ((m_SecPoint(i) * (Table_Xsize / Val(Set_Time))))
                    tmpY = 10 'm_TempPoint(i) * (Table_Ysize / Val(Set_Temp))
                    m_PlotXY(CalCounts - 1).rcpY = tmpY
                    m_PlotXY(CalCounts).rcpX = tmpX
                    m_PlotXY(CalCounts).rcpY = tmpY
                Else
                    tmpX = m_PlotXY(CalCounts - 1).rcpX
                    tmpY = 10 'm_TempPoint(i) * (Table_Ysize / Val(Set_Temp))
                    m_PlotXY(CalCounts).rcpX = tmpX
                    m_PlotXY(CalCounts).rcpY = tmpY
                    CalCounts = CalCounts + 1
                    tmpX = m_PlotXY(CalCounts - 1).rcpX + ((m_SecPoint(i) * (Table_Xsize / Val(Set_Time))))
                    tmpY = 10 'm_TempPoint(i) * (Table_Ysize / Val(Set_Temp))
                    m_PlotXY(CalCounts).rcpX = tmpX
                    m_PlotXY(CalCounts).rcpY = tmpY
                End If
            
            Case "Stop"
                m_SecPoint(i) = RecipeXY.arrayRecipe(i, GB_PROCESS_TIME)
                m_TempPoint(i) = RecipeXY.arrayRecipe(i, GB_PROCESS_TEMP)
                tmpX = m_PlotXY(CalCounts - 1).rcpX + m_SecPoint(i) * (Table_Xsize / Val(Set_Time))
                tmpY = m_TempPoint(i) * (Table_Ysize / Val(Set_Temp))
                m_PlotXY(CalCounts).rcpX = tmpX
                m_PlotXY(CalCounts).rcpY = tmpY
                Exit For
            Case "IO Control"
                m_SecPoint(i) = RecipeXY.arrayRecipe(i, GB_PROCESS_TIME)
                m_TempPoint(i) = RecipeXY.arrayRecipe(i, GB_PROCESS_TEMP)
                tmpX = m_PlotXY(CalCounts - 1).rcpX + m_SecPoint(i) * (Table_Xsize / Val(Set_Time))
                tmpY = m_TempPoint(i) * (Table_Ysize / Val(Set_Temp))
                m_PlotXY(CalCounts).rcpX = tmpX
                m_PlotXY(CalCounts).rcpY = tmpY
                
            Case "Pump Down"
                m_SecPoint(i) = .arrayRecipe(i, GB_PROCESS_TIME)
                m_TempPoint(i) = .arrayRecipe(i, GB_PROCESS_TEMP)
                If i = 1 Then
                    tmpX = m_PlotXY(CalCounts - 1).rcpX + ((m_SecPoint(i) * (Table_Xsize / Val(Set_Time))))
                Else
                    tmpX = m_PlotXY(CalCounts - 1).rcpX '+ ((m_SecPoint(i) * (Table_Xsize / Val(Set_Time))))
                End If
                tmpY = 0 'm_TempPoint(i) * (Table_Ysize / Val(Set_Temp))
                m_PlotXY(CalCounts).rcpX = tmpX
                m_PlotXY(CalCounts).rcpY = tmpY
            Case "Ramp Down"
                m_SecPoint(i) = RecipeXY.arrayRecipe(i, GB_PROCESS_TIME)
                m_TempPoint(i) = RecipeXY.arrayRecipe(i, GB_PROCESS_TEMP)

                tmpX = m_PlotXY(CalCounts - 1).rcpX + m_SecPoint(i) * (Table_Xsize / Val(Set_Time))
                tmpY = m_TempPoint(i) * (Table_Ysize / Val(Set_Temp))
                m_PlotXY(CalCounts).rcpX = tmpX
                m_PlotXY(CalCounts).rcpY = tmpY
        End Select
'        End If
        
        
    Next i
    'Call GetMaxTempAndTime(RcpXY)

    
    DoEvents        ' If coding error ..be saviful
    'ret = False
    'If bRet = True Then Call Plot_Recipe(m_Recipe, 0, 0, picObject)
    bRet = False
    'Plot Line of the Recipe
    For i = StartPic To EndPic
        picObject(i).DrawWidth = 3
    Next i
'    If m_PlotXY(1).rcpX > 0 Then
'        For i = StartPic To EndPic
'            picObject(i).Line (m_PlotXY(0).rcpX, m_PlotXY(0).rcpY)-(m_PlotXY(1).rcpX, m_PlotXY(1).rcpY), vbBlue
'        Next i
'    End If
'    sngDrawScale = (picObject(0).ScaleWidth / picObject(0).Width) * (picObject(0).Height / -picObject(0).ScaleHeight)
    sngDrawScale = (picObject(0).ScaleWidth / -picObject(0).ScaleHeight) * (picObject(0).Height / picObject(0).Width)
    
    sngCurveDistance = (gbsngSmoothTime * (Table_Xsize / Val(Set_Time)))
    For j = 1 To CalCounts - 1
        For i = StartPic To EndPic
        Check_Status = Trim(CStr(RecipeXY.arrayRecipe(j, GB_PROCESS_ACTION)))
        SysProName = Readini(Check_Status)
        If SysProName <> "" Then
        Check_Status = SysProName
        End If
            If gbintSmoothDisplay = 1 Then
                If Check_Status = "Ramp up" Then
                    Check_Status = Trim(CStr(RecipeXY.arrayRecipe(j + 1, GB_PROCESS_ACTION)))
                    If Check_Status = "Hold" Then
                        Call CalSmoothCurve(sngCurveDistance, _
                                            m_PlotXY(j - 1).rcpX, m_PlotXY(j - 1).rcpY, _
                                            m_PlotXY(j).rcpX, m_PlotXY(j).rcpY, _
                                            sngRetCurveData)
                        picObject(i).Line (m_PlotXY(j - 1).rcpX, m_PlotXY(j - 1).rcpY) _
                                            -(m_PlotXY(j).rcpX - (sngCurveDistance * Sin(sngRetCurveData(3))), _
                                              m_PlotXY(j).rcpY - (sngCurveDistance * Cos(sngRetCurveData(3)))), vbBlue
                        'picObject(i).Circle(Cx,Cy,R,color,startDegree,EndDegree)
                        picObject(i).Circle (sngRetCurveData(0), sngRetCurveData(1)), (sngRetCurveData(2) * sngDrawScale), _
                                            vbRed, GB_PI / 2, GB_PI - sngRetCurveData(3), sngDrawScale
                                            
                        picObject(i).Line (m_PlotXY(j).rcpX + sngCurveDistance, m_PlotXY(j).rcpY) _
                                            -(m_PlotXY(j + 1).rcpX, m_PlotXY(j + 1).rcpY), vbBlue
                                            
                        j = j + 1
                    Else
                        picObject(i).Line (m_PlotXY(j - 1).rcpX, m_PlotXY(j - 1).rcpY) _
                                            -(m_PlotXY(j).rcpX, m_PlotXY(j).rcpY), vbBlue
                    End If
                Else
                    picObject(i).Line (m_PlotXY(j - 1).rcpX, m_PlotXY(j - 1).rcpY) _
                                            -(m_PlotXY(j).rcpX, m_PlotXY(j).rcpY), vbBlue
                End If
            Else
                picObject(i).Line (m_PlotXY(j - 1).rcpX, m_PlotXY(j - 1).rcpY) _
                                        -(m_PlotXY(j).rcpX, m_PlotXY(j).rcpY), vbBlue
            End If
        Next i
    Next j

'    For j = 1 To CalCounts - 1
'        For i = StartPic To EndPic
'            picObject(i).Line (m_PlotXY(j).rcpX, m_PlotXY(j).rcpY) _
'                                        -(m_PlotXY(j + 1).rcpX, m_PlotXY(j + 1).rcpY), vbBlue
'        Next i
'    Next j
    
    End With
    Plot_Recipe = True
    Exit Function
    'frmPlotProcess.tabPlotProcess.Tab = tmpTable
ERR_PLOT_RECIPE:
    intCalCurveCount = i
    Plot_Recipe = False
End Function

Public Sub GetMaxTempAndTime(RecipeXY As RecipeTable)
    Dim tmpStr              As String
    Dim cx, cy, Cz          As Long
    Dim tmpVal              As Long
    
    Set_Time = CHART_DEF_TIME
    Set_Temp = CHART_DEF_TEMP
    tmpVal = 0
    'tmpStr = frmAutoProcess.cmbLotID.Text
'    For Cx = 0 To 49
''        For Cy = 0 To 9
'        tmpVal = tmpVal + Val(RecipeXY.arrayRecipe(Cx, 1))
'        If tmpVal > Set_Time Then
'            Set_Time = tmpVal
'        End If
'
'        If Set_Temp < Val(RecipeXY.arrayRecipe(Cx, 2)) Then
'            Set_Temp = Val(RecipeXY.arrayRecipe(Cx, 2))
'        End If
''        Next Cy
'    Next Cx
End Sub

Public Sub CalSmoothCurve(sngDistance As Single, sngX1 As Single, sngY1 As Single, sngX2 As Single, sngY2 As Single, sngResult() As Single)
    Dim sngTheta As Single
    Dim sngA As Single
    Dim sngB As Single
    Dim sngC As Single
    Dim sngR As Single
    Dim sngCx As Single
    Dim sngCy As Single
    'Dim sngResult(5) As Single
    Dim i As Integer
    
    For i = 0 To 5
        sngResult(i) = 0
    Next i
    
    If (sngX2 - sngX1) <> 0 Then
        sngTheta = Atn((sngX2 - sngX1) / (sngY2 - sngY1))
    
        sngB = sngDistance
        sngA = sngB * Sin(sngTheta)
        sngR = (sngA + sngB) / Cos(sngTheta)
        sngCx = sngX2 + sngB
        sngCy = sngY2 - sngR
        
        sngResult(0) = sngCx
        sngResult(1) = sngCy
        sngResult(2) = sngR
        sngResult(3) = sngTheta
    End If
    
    gbsngRampSmoothDist = sngB
    gbsngRampSmoothDistX = sngA
    gbsngRampSmoothDistY = sngB * Cos(sngTheta)
    gbsngRampsmoothTheta = sngTheta
    gbsngRampSmoothCx = sngCx
    gbsngRampSmoothCy = sngCy
    gbsngRampSmoothR = sngR
    gbsngRampSmoothStart = sngX2 - sngA
    gbsngRampSmoothMid = sngX2
    'Rev4.1.6 Fix the bug
    gbsngRampSmoothEnd = sngX2 + sngB
    'gbsngRampSmoothEnd = sngX2 + sngA
End Sub
