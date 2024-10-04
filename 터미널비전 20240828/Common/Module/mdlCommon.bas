Attribute VB_Name = "commonGlobal"
Public Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)

Public g_LastMesMsgId As String
Public g_LastMesMsgItem1 As String
Public g_LastMesMsgItem2 As String
Public g_LastMesSystemByte As String
Public g_LastMesMsg As String

Public g_Timeout As Long
Public g_TTemp As Long
Public g_TimeoutInterval As Long
Public g_TimeoutRetry As Long

Public g_bCameraInitialized As Boolean
Public g_btLightBrightness(0 To kMaxLight - 1) As Byte

Public g_ModelNumber As Integer
Public g_ModelChangedDate As String

'조명 On/Off
Public g_UseLightTimer As Long
Public g_LightTimerInterval As Long
Public g_LightTimerCount As Long

'재검사
Public g_UseRetry As Long  '사용여부
Public g_RetryBase As Long
Public g_RetryROI As Long

Public g_CamExposureTime(0 To kMaxCamera - 1) As Double


'검사결과 여부
Public g_SaveResultImage As Integer

'검사결과 여부
Public Sub SaveResultSaving(ByVal ModelName As String)

    Dim FileName As String
    Dim Section As String
    
    FileName = App.Path & "\model\" & ModelName & "\model.ini"
    Section = "OPTION"
    
    Call WriteINI(FileName, Section, "ResultImageSaving", CStr(g_SaveResultImage))

End Sub

Public Sub LoadResultSaving(ByVal ModelName As String)

    Dim FileName As String
    Dim Section As String
    
    FileName = App.Path & "\model\" & ModelName & "\model.ini"
    Section = "OPTION"
    
    g_SaveResultImage = CInt(ReadINI(FileName, Section, "ResultImageSaving", "0"))

End Sub

'재검사 관련 저장
Public Sub SaveRetryParameters(ByVal ModelName As String)
Dim FileName As String
Dim Section As String
    
    FileName = App.Path & "\model\" & ModelName & "\model.ini"
    Section = "Retry_Param"
 
    
    Call WriteINI(FileName, Section, "UseRetry", CStr(g_UseRetry))
    Call WriteINI(FileName, Section, "RetryBase", CStr(g_RetryBase))
    Call WriteINI(FileName, Section, "RetryROI", CStr(g_RetryROI))

End Sub

Public Sub LoadRetryParameters(ByVal ModelName As String)
Dim FileName As String
Dim Section As String
    
    FileName = App.Path & "\model\" & ModelName & "\model.ini"
    Section = "Retry_Param"

    On Error GoTo ErrorHandle
    g_UseRetry = CLng(ReadINI(FileName, Section, "UseRetry", "1"))
    g_RetryBase = CLng(ReadINI(FileName, Section, "RetryBase", "0"))
    g_RetryROI = CLng(ReadINI(FileName, Section, "RetryROI", "1"))
    
    Exit Sub
    
ErrorHandle:
    g_UseRetry = 1
    g_RetryBase = 0
    g_RetryROI = 1
    Call SaveRetryParameters(ModelName)

End Sub


'재검사 관련 불러오기


'자동조명 관련 저장
Public Sub SaveAutoLightParameters()
Dim FileName As String
Dim Section As String
    
    FileName = App.Path & "\" & App.title & ".ini"
    Section = "AutoLight_Param"
    
    Call WriteINI(FileName, Section, "UseAutoLight", CStr(g_UseLightTimer))
    Call WriteINI(FileName, Section, "LightTimerInterval", CStr(g_LightTimerInterval))
    
End Sub

'자동조명 관련 불러오기
Public Sub LoadAutoLightParameters()
Dim FileName As String
Dim Section As String
    
    FileName = App.Path & "\" & App.title & ".ini"
    Section = "AutoLight_Param"
    
    
    On Error GoTo ErrorHandle
    
    g_UseLightTimer = CLng(ReadINI(FileName, Section, "UseAutoLight", "1"))
    g_LightTimerInterval = CLng(ReadINI(FileName, Section, "LightTimerInterval", "60"))
    
    Exit Sub
    
ErrorHandle:
    g_UseLightTimer = 1
    g_LightTimerInterval = 60
    Call SaveAutoLightParameters
    
End Sub


Public Sub SendSignalToMelsec(ByVal Bit As Integer, ByVal Value As Integer)
    
    m_Snd_Bit_1(Bit) = Value
    Call Write_Send_Word(addVisionInspect1, Make_Send_Word(addVisionInspect1, True))

End Sub

Public Sub SaveTimeoutParam()

    Dim FileName As String
    Dim Section As String
    
    FileName = App.Path & "\" & App.title & ".ini"
    Section = "Timeout_Param"
    
    Call WriteINI(FileName, Section, "Interval", CStr(g_TimeoutInterval))
    Call WriteINI(FileName, Section, "Retry", CStr(g_TimeoutRetry))
    
End Sub

Public Sub LoadTimeoutParam()
    
    Dim FileName As String
    Dim Section As String
    
    FileName = App.Path & "\" & App.title & ".ini"
    Section = "Timeout_Param"
    
    g_TimeoutInterval = CLng(ReadINI(FileName, Section, "Interval"))
    g_TimeoutRetry = CLng(ReadINI(FileName, Section, "Retry"))
    
End Sub

Public Sub LoadSystemData()
On Error Resume Next

    Dim FileName As String
    Dim Section As String
    
    Dim i As Integer
    
    FileName = App.Path & "\" & App.title & ".ini"
    
    'System
    frmMain.lblProgramTitle.Caption = ReadINI(FileName, "SYSTEM", "PROGRAM_TITLE")
    For i = 0 To kMaxCamera - 1
        Dim title As String
        title = ReadINI(FileName, "SYSTEM", "CAMERA_TITLE" & CStr(i + 1))
        If title <> "" Then
            frmMain.lblCamBaseCaption(i).Caption = title
        End If
    Next i
    
    '마지막 작업 모델
    sModelName = ReadINI(FileName, "SYSTEM", "LAST_MODEL")
    
    '켈리브레이션
'    Section = "Calibration"
'    For i = 0 To kMaxCamera - 1
'        dCaliMM(i) = CDbl(ReadINI(FileName, Section, "CAMERA" & CStr(i + 1) & "_MM"))
'        dCaliPX(i) = CDbl(ReadINI(FileName, Section, "CAMERA" & CStr(i + 1) & "_MM/PIXEL"))
'    Next i
    
    '조명
    
    Section = "Light_Brightness"
    For i = 0 To kMaxLight - 1
        g_btLightBrightness(i) = CByte(ReadINI(FileName, Section, "CAMERA" & CStr(i + 1)))
        Call PWM_SetLight(CLng(i), CLng(g_btLightBrightness(i)))
    Next i
    
    Section = "Camera_ExposureTime"
    For i = 0 To kMaxCamera - 1
        g_CamExposureTime(i) = CDbl(ReadINI(FileName, Section, "CAMERA" & CStr(i + 1)))
'        If g_bCameraInitialized = True Then
'            Call frmMain.uEyeCam1(i).SetExposureTime(g_CamExposureTime(i))
'        End If
    Next i
    
    Call LoadTimeoutParam
    
End Sub

Public Sub SaveSystemData()
On Error Resume Next

    Dim FileName As String
    Dim Section As String
    
    Dim i As Integer
    
    FileName = App.Path & "\" & App.title & ".ini"
    
    '켈리브레이션
    Section = "Calibration"
    For i = 0 To kMaxCamera - 1
        Call WriteINI(FileName, Section, "CAMERA" & CStr(i + 1) & "_MM", CStr(dCaliMM(i)))
        Call WriteINI(FileName, Section, "CAMERA" & CStr(i + 1) & "_MM/PIXEL", CStr(dCaliPX(i)))
    Next i
    
    '조명
    Section = "Light_Brightness"
    For i = 0 To kMaxLight - 1
        Call WriteINI(FileName, Section, "CAMERA" & CStr(i + 1), CStr(g_btLightBrightness(i)))
    Next i
    
    Section = "Camera_ExposureTime"
    For i = 0 To kMaxCamera - 1
        Call WriteINI(FileName, Section, "CAMERA" & CStr(i + 1), CStr(g_CamExposureTime(i)))
    Next i
    
    Call SaveTimeoutParam
    
End Sub

Public Sub LoadCount()
On Error Resume Next

    Dim FileName As String
    Dim Section As String
    
    Dim i As Integer
    
    FileName = App.Path & "\" & App.title & ".ini"
    Section = "COUNT"
    
    lToTalCount = CLng(ReadINI(FileName, Section, "TOTAL"))
    lOKCount = CLng(ReadINI(FileName, Section, "OK"))
    lNGCount = CLng(ReadINI(FileName, Section, "NG"))
    
End Sub

Public Sub SaveCount()
On Error Resume Next

    Dim FileName As String
    Dim Section As String
    
    Dim i As Integer
    
    FileName = App.Path & "\" & App.title & ".ini"
    Section = "COUNT"
    
    Call WriteINI(FileName, Section, "TOTAL", CStr(lToTalCount))
    Call WriteINI(FileName, Section, "OK", CStr(lOKCount))
    Call WriteINI(FileName, Section, "NG", CStr(lNGCount))
    
End Sub

Public Function CheckTextBox(ByRef TB As TextBox, RangeStart As Long, RangeEnd As Long, Optional ValidColor As ColorConstants = vbWhite, Optional InvalidColor As ColorConstants = vbRed) As Boolean

    If TB Is Nothing Then
        CheckTextBox = False
        Exit Function
    End If
    
    Dim Number As Variant
    
    If IsNumeric(TB.Text) = False Then
        TB.BackColor = InvalidColor
        CheckTextBox = False
        Exit Function
    Else
        TB.BackColor = ValidColor
    End If
    
    Number = CLng(TB.Text)
    
    If Number < RangeStart Then
        TB.BackColor = InvalidColor
        TB.Text = CStr(RangeStart)
        Exit Function
    End If
    
    If Number > RangeEnd Then
        TB.BackColor = InvalidColor
        TB.Text = CStr(RangeEnd)
        Exit Function
    End If
    
    CheckTextBox = True

End Function

Public Function CheckLabel(ByRef TB As Label, RangeStart As Double, RangeEnd As Double, Optional ValidColor As ColorConstants = vbWhite, Optional InvalidColor As ColorConstants = vbRed) As Boolean

    If TB Is Nothing Then
        CheckLabel = False
        Exit Function
    End If
    
    Dim Number As Variant
    
    If IsNumeric(TB.Caption) = False Then
        TB.ForeColor = InvalidColor
        CheckLabel = False
        Exit Function
    Else
        TB.ForeColor = ValidColor
    End If
    
    Number = CDbl(TB.Caption)
    
    If Number < RangeStart Then
        TB.ForeColor = InvalidColor
        Exit Function
    End If
    
    If Number > RangeEnd Then
        TB.ForeColor = InvalidColor
        Exit Function
    End If
    
    CheckLabel = True

End Function

Public Function WriteMelsec(ByRef Melsec As ActEasyIF, Address As String, Data As Long) As Long
On Error GoTo ErrorHandle
    
    Dim nRet As Long
    
    WriteMelsec = Melsec.WriteDeviceRandom(Address, 1, Data)
    
    Exit Function
ErrorHandle:
    WriteMelsec = -1
    
End Function

Public Function ReadMelsec(ByRef Melsec As ActEasyIF, Address As String, Optional Display As Boolean = False) As Long
On Error GoTo ErrorHandle

    Dim Data As Integer
    Dim nRet As Long
    Dim i As Integer
    
    nRet = Melsec.ReadDeviceRandom2(Address, 1, Data)
    
    If nRet <> 0 Then
        ReadMelsec = -1
        If Display = True Then
            For i = 0 To 3
                frmMain.shpInput(i).FillColor = vbGrayText
            Next i
        End If
        Exit Function
    End If
    
    If Display = True Then
        For i = 0 To 3
            frmMain.shpInput(i).FillColor = IIf(GetBit(Data, i) = 1, vbRed, &H80000005)
        Next i
    End If
    
    ReadMelsec = Data
    
    Exit Function
ErrorHandle:

    ReadMelsec = -1
    
End Function

Public Function ReadMelsec2(ByRef Melsec As ActEasyIF, Address As String) As Long
On Error GoTo ErrorHandle

    Dim Data(1) As Integer
    Dim nRet As Long
    Dim i As Integer
    
    nRet = Melsec.ReadDeviceRandom2(Address, 2, Data(0))
    
    If nRet <> 0 Then
        ReadMelsec2 = -1
        Exit Function
    End If
    
    ReadMelsec2 = Data(0) + (Data(1) * 65536)
    
    Exit Function
ErrorHandle:

    ReadMelsec2 = -1
    
End Function

Public Function ReadMelsecModel2(ByRef Melsec As ActEasyIF, Address As String) As Long
On Error GoTo ErrorHandle

    Dim Data(1) As Integer
    Dim nRet As Long
    Dim i As Integer
    
    Dim AddressList As String
    
    AddressList = GetAddressString(Address, 2)
    
    nRet = Melsec.ReadDeviceRandom2(AddressList, 2, Data(0))
    If nRet <> 0 Then
        ReadMelsecModel2 = 0
        Exit Function
    End If
    
    Data(0) = Data(0) + (Data(1) * 65536)
    
    ReadMelsecModel2 = BitPos(Data(0)) + 1
    
    Exit Function
ErrorHandle:
    ReadMelsecModel2 = 0
    
End Function

Public Function Range(ByVal Data As Variant, ByVal LOW As Variant, HIGH As Variant) As Boolean

    Range = (Data >= LOW) And (Data <= HIGH)

End Function

Public Function BitPos(ByVal Value As Long) As Integer
    
    Dim temp As Integer
    Dim i As Integer
    
    For i = 0 To 20
        temp = 2 ^ (i + 1)
        If Value Mod temp >= 2 ^ i Then
            BitPos = i
            Exit Function
        End If
    Next i
    
    BitPos = -1
    
End Function

Public Function GetBit(ByVal Value As Long, Pos As Integer) As Integer

    Dim temp As Integer
    temp = 2 ^ (Pos + 1)
    GetBit = IIf(Value Mod temp >= 2 ^ Pos, 1, 0)
    
End Function

Public Function ModelChange() As Boolean
On Error GoTo ErrorHandle

    Dim ModelNumber As Long
    Dim Room As Integer
    
    ModelNumber = ReadMelsecModel2(frmMain.ActEasyIF, sMelsecAddrModelNumber)
    
    If ModelNumber = g_ModelNumber Or ModelNumber <= 0 Then
        ModelChange = False
        Exit Function
    End If
    
    ModelChange = True
    
    frmMain.txtModelNumber.Text = CStr(ModelNumber)
    
    frmMain.shpModel(0).FillColor = vbRed
    frmMain.shpModel(1).FillColor = vbWhite
    Sleep 200
    
    g_ModelNumber = ModelNumber
    sModelName = sModelRoom(ModelNumber)
    g_ModelChangedDate = Format(Date, "yyyy.mm.dd") & " " & Format(Time, "hh:mm:ss")
    Call LoadModel(sModelName)
    Call LastModelWrite
    
    frmMain.lblModelNameMain.Caption = sModelName
    
    frmMain.shpModel(0).FillColor = vbWhite
    frmMain.shpModel(1).FillColor = vbRed
    frmMain.lblChangedModel.Caption = g_ModelChangedDate
    Sleep 200
    
    Exit Function
ErrorHandle:
    ModelChange = False
    frmMain.shpModel(0).FillColor = vbYellow
    frmMain.shpModel(1).FillColor = vbYellow
    
End Function


Public Function AlarmToPLC()
On Error GoTo err:

    Dim strDeviceList As String
    Dim nSize As Long
    Dim nData() As Long
    Dim nResult As Long
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim Count As Integer
    
    For k = 0 To 3
        For i = 0 To 9
            If i = 0 And k = 0 Then
            Else
                strDeviceList = strDeviceList & vbLf
            End If
            strDeviceList = strDeviceList & GetAddressString(lMelsecAddrInspection(k, i), 2)
            
        Next i
    Next k
    
    For i = 0 To Len(strDeviceList)
        If Right$(Left$(strDeviceList, i + 1), 1) = vbLf Then
            Count = Count + 1
        End If
    Next i
    
    nSize = Count + 1
    
    Debug.Print nSize, strDeviceList
    
    ReDim nData(nSize)
    
    For i = 0 To nSize
        nData(i) = 0
    Next i
    
    Count = 0
    j = 0
    For i = 0 To 79 Step 2
        k = (i / 2) Mod 10
        nData(i) = dInspectResult_mm(j, k) * 100
        Debug.Print i, j, k
        Count = Count + 1
        If Count > 9 Then
            Count = 0
            j = j + 1
        End If
    Next i
    
    nData(80) = 0 '불량코드 만들어서 넣어 주세요!!
    
    nResult = frmMain.ActEasyIF.WriteDeviceRandom(strDeviceList, nSize, nData(0))
    
    
    Exit Function
err:
End Function


