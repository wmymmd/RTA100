Attribute VB_Name = "mdlQueryData"
Option Explicit

'Rev8.0.1.5 Integrated the temperature input
Public gbsngCurrTemperature(20)  As Single
Public gbsngCurrTempRet(20)  As Single
Public gbsngFakeValue(10) As Single
Public PlayFakeBall(10) As Single




'Rev8.0.1.5
Public Sub ReadAI_ADT()
    Dim i As Integer
    Dim offset As Integer
    Dim rand As Single
    Dim fTemp As Single
    
    
    On Error GoTo ERRLINE
    If advThermo.IsActive = False Then Exit Sub
    Call advThermo.ReadTemperatureAllChannelFT(Kernel.sngTC)
    For i = 0 To 7
        Kernel.sngTC(i) = Kernel.sngTC(i) * gbsngRatioTC(i) * gbsngRatioEX(i) + gbsngPowerTC(i) * sngTemp_1(i) * sngTemp_1(i) + gbsngErrorTC(i)
        
    Next i
    
    If gbblnPlayFakeBall = True And gbintCurrProcessStep = GB_ACTION_INDEX_HOLD Then
        offset = Kernel.sngTC(0) / 100 - 1
        For i = 1 To 7
            If (Kernel.sngTC(i) + PlayFakeBall(i)) < (Kernel.sngTC(0) - offset) Then
                PlayFakeBall(i) = PlayFakeBall(i) + 0.1
                
            ElseIf (Kernel.sngTC(i) + PlayFakeBall(i)) > (Kernel.sngTC(0) + offset) Then
                PlayFakeBall(i) = PlayFakeBall(i) - 0.1
                
            Else
                Randomize
                fTemp = Rnd * 0.3 - Rnd * 0.3
                PlayFakeBall(i) = PlayFakeBall(i) + fTemp
                
            End If
            
            
            Kernel.sngTC(i) = Kernel.sngTC(i) + PlayFakeBall(i)
        Next i
    Else
        
        For i = 0 To 7
            PlayFakeBall(i) = 0
        Next i
        gbblnPlayFakeBall = False
    End If

    Exit Sub
ERRLINE:
    MsgBox ("Error in Get Temperature!!")
    
End Sub


Public Sub ReadAI_HGU()
    Dim i As Integer
    Dim offset As Integer
    Dim rand As Single
    Dim fTemp As Single
    
    
    On Error GoTo ERRLINE
    If advThermo.IsActive = False Then Exit Sub
    Call advThermo.ReadTemperatureAllChannelFT(Kernel.sngTC)
    For i = 0 To 7
        Kernel.sngTC(i) = Kernel.sngTC(i) * gbsngRatioTC(i) * gbsngRatioEX(i) + gbsngPowerTC(i) * sngTemp_1(i) * sngTemp_1(i) + gbsngErrorTC(i)
        
    Next i
    
    If gbblnPlayFakeBall = True And gbintCurrProcessStep = GB_ACTION_INDEX_HOLD Then
        offset = Kernel.sngTC(0) / 100 - 1
        For i = 1 To 7
            If (Kernel.sngTC(i) + PlayFakeBall(i)) < (Kernel.sngTC(0) - offset) Then
                PlayFakeBall(i) = PlayFakeBall(i) + 0.1
                
            ElseIf (Kernel.sngTC(i) + PlayFakeBall(i)) > (Kernel.sngTC(0) + offset) Then
                PlayFakeBall(i) = PlayFakeBall(i) - 0.1
                
            Else
                Randomize
                fTemp = Rnd * 0.3 - Rnd * 0.3
                PlayFakeBall(i) = PlayFakeBall(i) + fTemp
                
            End If
            
            
            Kernel.sngTC(i) = Kernel.sngTC(i) + PlayFakeBall(i)
        Next i
    Else
        
        For i = 0 To 7
            PlayFakeBall(i) = 0
        Next i
        gbblnPlayFakeBall = False
    End If

    Exit Sub
ERRLINE:
    MsgBox ("Error in Get Temperature!!")
    
End Sub
