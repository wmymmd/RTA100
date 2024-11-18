Attribute VB_Name = "mdlPIDControl"
Option Explicit

Public Type PIDParameterInfo
    Kp(100) As Single
    Ki(100) As Single
    Kd(100) As Single
    Kpd(100) As Single
    Kff(100) As Single
    Kpp(100) As Single
End Type

Public PIDParameter As PIDParameterInfo

'Define variable of the PID
Dim CO              As Single               'Controller Output
Dim Kp              As Single               'Proportional Constant
Dim Ki              As Single               'Integral Contant
Dim Kd              As Single               'Differential Contant
Dim Kpd              As Single               'Predit Contant
Dim Kff             As Single               'FeedForward Contant
Dim PID_Time        As Single               'PID time
Dim PID_E           As Single               'Error
Dim SP              As Single               'Set Point
Dim PV              As Single               'Process Variable
Dim Proportional    As Single               'Proportional Control
Dim Integernal      As Single               'Integernal Control
Dim Derivative      As Single               'Derivative Control
Dim PreError        As Single               'Pre_times Error Value

Dim VacuumSigma        As Single

'--------------------------------------------------------------------------------------
'Init PID Parameter
'--------------------------------------------------------------------------------------
Public Sub InitPIDParameter()
    gbProcessSecondCount = 0
    PID_Time = 1
    PV = 0
    Proportional = 0
    Integernal = 0
    Derivative = 0
End Sub
'--------------------------------------------------------------------------------------
'Set PID Parameter
'--------------------------------------------------------------------------------------
Public Sub SetPIDParameter(sngSP As Single, sngGain As Single, sngIntegral As Single, sngDerivative As Single, _
                                                    sngPredit As Single, sngFeedfForward As Single)
    SP = sngSP
    Kp = sngGain
    Ki = sngIntegral
    Kd = sngDerivative
    Kpd = sngPredit
    Kff = sngFeedfForward
End Sub

'--------------------------------------------------------------------------------------
'PID Control Loop
'--------------------------------------------------------------------------------------
Public Function PID_Loop(ByVal PresentValue As Single, ByVal lngInterval As Long) As Single 'or (ByVal Model As Long)
    Dim sngSlope As Single
    Dim sngPreditTemp As Single
    Dim sngRecipeSlope As Single
    Dim sngFeedForward As Single
    'Basic PID Control
    CO = 0
   
    PV = PresentValue
    PID_E = SP - PV
    
    PID_Time = 1
    sngSlope = PID_E / lngInterval 'PID_Time
    sngPreditTemp = PV + (sngSlope * Kpd) 'predit

    If gbProcessSecondCount > 0 Then
        sngRecipeSlope = (gbProcessRecipe(gbProcessSecondCount).sngTemperature - gbProcessRecipe(gbProcessSecondCount - 1).sngTemperature) / lngInterval
    Else
        sngRecipeSlope = 0
    End If
    'sngFeedForward = sngRecipeSlope * Kff
    sngFeedForward = gbsngProcessSlope * Kff

    
    PID_E = SP - sngPreditTemp
    Proportional = Kp * (PID_E)
    'Rev4.1.3 calculate by second
    Integernal = Integernal + (Ki * (PID_E * (lngInterval / 1000)))
        
    Derivative = Kd * ((PID_E) / (lngInterval / 1000))
    'Derivative = Kd * ((PID_E - PreError) / (lngInterval / 1000))
    'Integernal = Integernal + (Ki * (PID_E * lngInterval))
    'Derivative = Kd * ((PID_E - PreError) / lngInterval)
    CO = Proportional + Integernal + Derivative + sngFeedForward
    

    If CO > 10 Then
        CO = 10
    ElseIf CO < 0 Then
        CO = 0
    End If
    PreError = PID_E
    PID_Loop = CO
End Function

Public Function CalMultiPID(LoopNo As Integer, IsReset As Boolean, VP As Single, VI As Single, VD As Single, Target As Single, Feedback As Single, lngInterval As Long) As Single
    Dim Vout As Single
    
    If IsReset Then
        MultiLoop.sngIntergalSigma(LoopNo) = 0
        Exit Function
    End If
    MultiLoop.sngIntergalSigma(LoopNo) = MultiLoop.sngIntergalSigma(LoopNo) + VI * (Target - Feedback)
    Vout = VP * (Target - Feedback) + MultiLoop.sngIntergalSigma(LoopNo) + VD * (Target - Feedback) / lngInterval / 1000
    
    If Vout > 10 Then
        Vout = 10
    End If
    If Vout < 0 Then
        Vout = 0
    End If
    CalMultiPID = Vout
    
End Function
Public Function CalVacuumPID(IsReset As Boolean, VP As Single, VI As Single, Target As Single, Feedback As Single) As Single
    Dim Vout As Single
    
    If IsReset Then
        VacuumSigma = 0
        Exit Function
    End If
    VacuumSigma = VacuumSigma + VI * (Target - Feedback)
    Vout = VP * (Target - Feedback) + VacuumSigma
    
    If Vout > 100 Then
        Vout = 100
    End If
    If Vout < 0 Then
        Vout = 0
    End If
    CalVacuumPID = Vout
    
    
End Function

Public Function Percent2Volt(percent As Single, limit1 As Integer, limit2 As Integer)
    Dim Value As Single
    
    Value = (limit2 - limit1) * percent / 100
    Percent2Volt = Value

End Function

Public Function Volt2MFC(Index As Integer, Volt As Single) As Single
    Dim Value As Single
    
    If gbintGasEnable(Index) > 0 And gbsngMaxGasSLMP(Index) > 0 Then Value = Volt / 5 * gbsngMaxGasSLMP(Index)
    Volt2MFC = Value
End Function

