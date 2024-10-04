Attribute VB_Name = "commonIOLightModule"
Option Explicit

Public m_bLightExist As Boolean             '조명 컨트롤러 유/무 변수

Public Function InitLightIO() As Boolean
On Error GoTo ErrorHandelr

    Dim tempInitRight As Long
    
    tempInitRight = OpenDAQDevice
    Set_Pwm 0, 255
    Set_Pwm 1, 255
    Set_Pwm 2, 255
    Set_Pwm 3, 255
    
    If (tempInitRight = -1) Then
        InitLightIO = False
    Else
        InitLightIO = True
    End If
    
    Exit Function
ErrorHandelr:
    InitLightIO = False
  '  WriteLog ("InitLightIO (" & erR.Number & ") : " & erR.Description)
End Function

Public Function LightControl(ByVal channel As Integer, ByVal OnOff As Boolean, Optional ByVal Delay As Long = 0)
On Error GoTo ErrorHandelr
    If m_bLightExist = False Then Exit Function

    If (OnOff = True) Then
        '조명 켜기
        Pwm_Enable channel
        
        If (Delay > 0) Then
            Call Delaytime(Delay)
        End If
    Else
        If (Delay > 0) Then
            Call Delaytime(Delay)
        End If
        
        '조명 끄기
        Pwm_Disable channel
    End If
    Exit Function
ErrorHandelr:
    m_bLightExist = False
  '  WriteLog ("LightControl (" & erR.Number & ") : " & erR.Description)
End Function

Public Function CloseLightIO() As Boolean
On Error GoTo ErrorHandelr
    
    Dim bRet        As Boolean
    
    bRet = CloseDAQDevice
    
    CloseLightIO = True
    
    Exit Function
ErrorHandelr:
    CloseLightIO = False
  '  WriteLog ("CloseLightIO (" & erR.Number & ") : " & erR.Description)
End Function

