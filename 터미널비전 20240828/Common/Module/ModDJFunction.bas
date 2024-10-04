Attribute VB_Name = "commonDJFunction"
Public Function InitCamera() As Boolean
' 카메라 초기화
    Dim i As Integer
    Dim lRet As Long
    
    
    For i = 0 To kMaxCamera - 1
        frmMain.uEyeCam1(i).EnableErrorReport = False
        lRet = frmMain.uEyeCam1(i).InitCamera(i + 1)
        
        
        If lRet = 0 Then
            Call frmMain.uEyeCam1(i).SetPixelClock(92)
            Call frmMain.uEyeCam1(i).SetFrameRate(13)
            frmMain.uEyeCam1(i).EnableAutoExposure = False
            Call frmMain.uEyeCam1(i).SetExposureTime(10)
            Call frmMain.uEyeCam1(i).SetColorMode(6)
            Call frmMain.uEyeCam1(i).SetExternalTrigger(8)
        Else
            InitCamera = False
            Exit Function
        End If
    Next i
    
    InitCamera = True
 
End Function

Public Sub MasterImage_Copy(sOldModel As String, sNewModel As String)
    
    Dim i As Integer
    
    For i = 0 To kMaxCamera - 1
        Set g_CogImage(i) = LoadCogImage(App.Path & "\Model\" & sOldModel & "\" & "Master" & CStr(i) & ".bmp")
        Call SaveCogImage(App.Path & "\Model\" & sNewModel & "\" & "Master" & CStr(i) & ".bmp", g_CogImage(i))
    Next i

End Sub
Public Sub PN_Select(Index As Integer)
Dim i As Integer
Select Case Index

    Case 0
        'frmMain.lblProgramName.Caption = "J/R 적층 Align 검사"
        'frmMain.ShpCamB(1).Visible = False
        'frmMain.lblCamName(1).Visible = False
       ' frmMain.FavisImageBoxMain(1).Visible = False
        iCamNumber = 0
    Case 1
        'frmMain.lblProgramName.Caption = "Terminal Vision"
        'frmMain.ShpCamB(1).Visible = False
        'frmMain.lblCamName(1).Visible = False
        'frmMain.FavisImageBoxMain(1).Visible = False
        iCamNumber = 3
    Case 2
        'frmMain.lblProgramName.Caption = "물류 Vision"
        'frmMain.ShpCamB(1).Visible = False
        'frmMain.lblCamName(1).Visible = False
        'frmMain.FavisImageBoxMain(1).Visible = False
        iCamNumber = 3
End Select
    iProName = Index
    Call ProgramSelect_Save
End Sub

Public Sub Add_CboList()
    Dim i As Integer
    
'    frmSetting.cboCaliper.Clear
'    frmSetting.cboBlob.Clear
    
'    For i = 1 To iBlobToolCount
'        frmSetting.cboBlob.AddItem CStr(i)
'    Next i
'
'
'    frmSetting.cboFixture.AddItem "사용안함"
'    frmSetting.cboFixture.AddItem "X , Y 사용"
'    frmSetting.cboFixture.AddItem "각도만 사용"
'
'    For i = 0 To (iToolCount / 2) - 1
'        frmSetting.cboCaliper.AddItem sSpecName(i)
'    Next i
    
End Sub
Public Sub FixtureAngle_Run(Index As Integer)
On Error Resume Next
Dim dTemp(0 To 3) As Double
Dim dX1(0 To 1) As Long
Dim dX2(0 To 1) As Long
Dim dY1(0 To 1) As Long
Dim dY2(0 To 1) As Long

    dTemp(Index) = Format(LineAngle(Fix_PointAngleRunX(Index, 0), Fix_PointAngleRunY(Index, 0), Fix_PointAngleRunX(Index, 1), Fix_PointAngleRunY(Index, 1), 0), "00.000")
    Fix_ShiftPointAngle(Index) = dTemp(Index) - Fix_PointAngle(Index)

    dX1(0) = Fix_PointAngleRunX(Index, 0)
    dX2(0) = Fix_PointAngleRunX(Index, 1)
    dY1(0) = Fix_PointAngleRunY(Index, 0)
    dY2(0) = Fix_PointAngleRunY(Index, 1)
    
'    frmMain.FavisImageBoxMain(index).AddStaticLine dX1(0), dY1(0), dX2(0), dY2(0)
'    frmMain.FavisImageBoxMain(index).UpdateDraw
    
End Sub

Public Sub Fixture_Run(Index As Integer)
On Error Resume Next
Dim dX1(0 To 1) As Long
Dim dX2(0 To 1) As Long
Dim dY1(0 To 1) As Long
Dim dY2(0 To 1) As Long

    Fix_ShiftPointX(Index) = Fix_PointRunX(Index) - Fix_PointX(Index)
    Fix_ShiftPointY(Index) = Fix_PointRunY(Index) - Fix_PointY(Index)
    'x
    dX1(0) = Fix_PointRunX(Index) - 20
    dX2(0) = Fix_PointRunX(Index) + 20
    dY1(0) = Fix_PointRunY(Index)
    dY2(0) = Fix_PointRunY(Index)
    'y
    dX1(1) = Fix_PointRunX(Index)
    dX2(1) = Fix_PointRunX(Index)
    dY1(1) = Fix_PointRunY(Index) - 20
    dY2(1) = Fix_PointRunY(Index) + 20
    
'    frmMain.FavisImageBoxMain(index).AddStaticLine dX1(0), dY1(0), dX2(0), dY2(0)
'    frmMain.FavisImageBoxMain(index).AddStaticLine dX1(1), dY1(1), dX2(1), dY2(1)
'    frmMain.FavisImageBoxMain(index).UpdateDraw
    
End Sub

Public Function DJSJ_Point2Line(L_x1 As Double, L_y1 As Double, L_x2 As Double, L_y2 As Double, P_x As Double, P_y As Double) As Double
'-------------------------------------------------------------------------------------------------
'직선과 한점의 거리 구하기                                               20100219 DJSJ
'L_x1 , L_y1 , L_x2 , L_y2 에 선의 좌표 입력받음
'P_x , P_y 에 점의 좌표 입력 받음
'-------------------------------------------------------------------------------------------------
On Error GoTo err:
Dim L_slope As Double '직선의 기울기
Dim L_tempValue As Double '직선의 방정식 y = mx + b 에서    b 의 값

    If L_x1 = 0 And L_y1 = 0 And L_x2 = 0 And L_y2 = 0 Then
        GoTo err:
    ElseIf P_x = 0 And P_y = 0 Then
        GoTo err:
    End If
    
    If (L_x2 - L_x1) = 0 Then
        L_x2 = L_x2 + 0.000000001
        L_slope = (L_y2 - L_y1) / (L_x2 - L_x1)
        L_tempValue = L_y1 - (L_slope * L_x1)
    ElseIf (L_y2 - L_y1) = 0 Then
        L_slope = 0
        L_tempValue = L_y1 - (L_slope * L_x1)
    Else
        L_slope = (L_y2 - L_y1) / (L_x2 - L_x1)
        L_tempValue = L_y1 - (L_slope * L_x1)
    End If
    
    'L_tempValue = L_y1 - (L_slope * L_x1)
    
    DJSJ_Point2Line = Abs((P_y - (L_slope * P_x) - L_tempValue)) / Sqr((L_slope ^ 2) + 1)
Exit Function
err:
DJSJ_Point2Line = 0

End Function
Public Sub Dlay_T(ttm As Single)
Dim tmp_timer As Single
Dim tm As Single

tmp_timer = Timer
tm = ttm + Timer

Do
    DoEvents
    If tmp_timer > Timer Then
        Exit Do
    End If
Loop Until tm <= Timer

End Sub


Public Sub Main_ImageBoxClear(Index As Integer)
'    frmMain.FavisImageBoxMain(Index).DeleteStaticAll
'    frmMain.FavisImageBoxMain(Index).DeleteInteractiveAll
End Sub
Public Sub Setting_ImageBoxClear(Index As Integer)
    'frmSetting.FavisImageBoxSetting.DeleteStaticAll
    'frmSetting.FavisImageBoxSetting.DeleteInteractiveAll
End Sub
Public Sub Main_ImageBoxHide(Index As Integer)
'    frmMain.FavisImageBoxMain(Index).HideAll
End Sub
Public Sub Setting_ImageBoxHide()
'    frmSetting.FavisImageBoxSetting.HideAll
End Sub
Public Sub Live_Image(Index As Integer)

On Error Resume Next
    
'    frmMain.FavisImageBoxMain(Index).DeleteAll
'
'    frmMain.FavisImageBoxMain(Index).Visible = False
    
End Sub
Public Sub Stop_Image(Index As Integer)

On Error Resume Next
    
'    frmMain.FavisImageBoxMain(Index).DeleteAll
'
'    frmMain.FavisImageBoxMain(Index).Visible = True
    
'    Call Acquire_Image(Index, 0)
End Sub

Public Sub CaliperT_Region(Index As Integer, Tnum As Integer, frmindex As Integer)
    Dim CX As Double
    Dim CY As Double
    Dim ret As Boolean
    Select Case frmindex
    Case 0
'        frmMain.FavisImageBoxMain(Index).color = vbCyan
'        If Fix_UseMode(Index) = 2 And Tnum Mod 2 = 1 Then
'            ret = RotationPoint(cx, cy, Modelinfo.favEdgeTCenterX(Index, Tnum - 1), Modelinfo.favEdgeTCenterY(Index, Tnum - 1), Modelinfo.favEdgeTCenterX(Index, Tnum), Modelinfo.favEdgeTCenterY(Index, Tnum), Fix_ShiftPointAngle(Index))
'            frmMain.FavisImageBoxMain(Index).AddInteractiveRotatgeRect cx, cy, _
'                Modelinfo.favEdgeTSideX(Index, Tnum), Modelinfo.favEdgeTSideY(Index, Tnum), Modelinfo.favEdgeTRotation(Index, Tnum) - Fix_ShiftPointAngle(Index)
'        Else
'            frmMain.FavisImageBoxMain(Index).AddInteractiveRotatgeRect Modelinfo.favEdgeTCenterX(Index, Tnum) + Fix_ShiftPointX(Index), Modelinfo.favEdgeTCenterY(Index, Tnum) + Fix_ShiftPointY(Index), _
'                                        Modelinfo.favEdgeTSideX(Index, Tnum), Modelinfo.favEdgeTSideY(Index, Tnum), Modelinfo.favEdgeTRotation(Index, Tnum) - Fix_ShiftPointAngle(Index)
'        End If
'        frmMain.FavisImageBoxMain(Index).UpdateDraw
    Case 1
'        frmSetting.FavisImageBoxSetting.color = vbCyan
'        frmSetting.FavisImageBoxSetting.AddInteractiveRotatgeRect Modelinfo.favEdgeTCenterX(Index, Tnum), Modelinfo.favEdgeTCenterY(Index, Tnum), _
'                                        Modelinfo.favEdgeTSideX(Index, Tnum), Modelinfo.favEdgeTSideY(Index, Tnum), Modelinfo.favEdgeTRotation(Index, Tnum)
'        frmSetting.FavisImageBoxSetting.UpdateDraw
    End Select
End Sub


'=================================================
' 함수명 : ListBox_Append
' 인자 : 리스트 박스에 추가할 문자열
' 인수 :
' 함수설명 : 리스트 박스에 문자를 추가시킨다.
'==================================================
Public Sub ListBox_Append(Str As String, Index As Integer)
On Error GoTo err
    
    Select Case Index
    Case 0
'        If frmMain.lstPLCSocket.ListCount > 50 Then
'
'            frmMain.lstPLCSocket.Clear
'
'        End If
'
'            frmMain.lstPLCSocket.AddItem Str
'            frmMain.lstPLCSocket.ListIndex = frmMain.lstPLCSocket.ListCount - 1
    Case 1
        If frmMain.lstMESSocket.ListCount > 50 Then

            frmMain.lstMESSocket.Clear

        End If

            frmMain.lstMESSocket.AddItem Str
            frmMain.lstMESSocket.ListIndex = frmMain.lstMESSocket.ListCount - 1
    End Select
Exit Sub
    
err:

End Sub
Public Function Delete_Model_Dir(TempName As String) As Boolean
On Error GoTo LPerr
    RmDir App.Path & "\Model\" & TempName
    Delete_Model_Dir = True
Exit Function
    
LPerr:
    Delete_Model_Dir = False
End Function

Public Function Delete_Model_File(TempName As String) As Boolean
On Error GoTo LPerr
    'Kill App.Path & "\model\" & TempName & "\" & TempName & ".dat"
    Kill App.Path & "\Model\" & TempName & "\" & "*.*"
    Delete_Model_File = True
Exit Function
    
LPerr:
    Delete_Model_File = False
End Function

Public Function LOGWrite(LogStr As String)
On Error Resume Next
' 쓰기
    Dim FileNameT As String
    Dim FileNumberT As Integer
    FileNameT = App.Path & "\Model" & "\ModelLog.txt"
    FileNumberT = FreeFile
    Open FileNameT For Append As FileNumberT
        Print #FileNumberT, Date & " - " & Time & "  :  " & LogStr
    Close #FileNumberT
End Function

Public Function KillFile(KillName As String) As Boolean
On Error GoTo LPerr
    Kill App.Path & "\Model\" & Trim(KillName) & ".dat"
    KillFile = True
Exit Function

LPerr:
    KillFile = False
End Function

Public Sub DJ_ModelLoad(mdlname As String)
On Error Resume Next
Dim i As Integer
    Call DJ_ToolCountLoad(mdlname)
    Call ModelData_FileLoad(mdlname)
    Call Calibration_Load(mdlname)
    'Call Calibration_Loady(mdlname)
    Call FixPoint_Load(mdlname)
    Call SpecName_Load(mdlname)
    Call SpecAllValue_Load(mdlname)
    Call FunctionValue_Load(mdlname)
End Sub

Public Sub DJ_ModelSave(mdlname As String)
    Call Modeldata_Filesave(mdlname)
    Call DJ_ToolCountSave(mdlname)
    Call Calibration_Save(mdlname)
    'Call Calibration_Savey(mdlname)
    Call FixPoint_Save(mdlname)
    Call SpecName_Save(mdlname)
    Call SpecAllValue_Save(mdlname)
    Call FunctionValue_Save(mdlname)
End Sub

Public Function DJ_HexToBin(sHexData As String) As String  'Hex 를 binary 로 바꿈
On Error Resume Next

Dim i As Integer
Dim tempstr As String
Dim tempLen As Integer
Dim tempHex(1 To 4) As String
    
    tempLen = Len(sHexData)
    For i = 1 To tempLen
        tempHex(i) = Mid(sHexData, i, 1)
        Select Case tempHex(i)
            Case "0"
                tempstr = tempstr & "0000"
            Case "1"
                tempstr = tempstr & "0001"
            Case "2"
                tempstr = tempstr & "0010"
            Case "3"
                tempstr = tempstr & "0011"
            Case "4"
                tempstr = tempstr & "0100"
            Case "5"
                tempstr = tempstr & "0101"
            Case "6"
                tempstr = tempstr & "0110"
            Case "7"
                tempstr = tempstr & "0111"
            Case "8"
                tempstr = tempstr & "1000"
            Case "9"
                tempstr = tempstr & "1001"
            Case "A"
                tempstr = tempstr & "1010"
            Case "B"
                tempstr = tempstr & "1011"
            Case "C"
                tempstr = tempstr & "1100"
            Case "D"
                tempstr = tempstr & "1101"
            Case "E"
                tempstr = tempstr & "1110"
            Case "F"
                tempstr = tempstr & "1111"
        End Select
    Next i

    DJ_HexToBin = tempstr
    
End Function
Public Function DJ_BintoHex_16bit(sBinary As String) As String  'binary 를 hex 로 바꿈
On Error Resume Next

Dim i As Integer
Dim tempstr As String
Dim tempBin As String
Dim tempLen As Integer
Dim tempBinary(1 To 4) As String
    
    tempBin = Left(sBinary & "0000000000000000", 16) '16비트
    
    For i = 1 To 4
        tempBinary(i) = Mid(tempBin, (i * 4) - 3, 4)
        Select Case tempBinary(i)
            Case "0000"
                tempstr = tempstr & "0"
            Case "0001"
                tempstr = tempstr & "1"
            Case "0010"
                tempstr = tempstr & "2"
            Case "0011"
                tempstr = tempstr & "3"
            Case "0100"
                tempstr = tempstr & "4"
            Case "0101"
                tempstr = tempstr & "5"
            Case "0110"
                tempstr = tempstr & "6"
            Case "0111"
                tempstr = tempstr & "7"
            Case "1000"
                tempstr = tempstr & "8"
            Case "1001"
                tempstr = tempstr & "9"
            Case "1010"
                tempstr = tempstr & "A"
            Case "1011"
                tempstr = tempstr & "B"
            Case "1100"
                tempstr = tempstr & "C"
            Case "1101"
                tempstr = tempstr & "D"
            Case "1110"
                tempstr = tempstr & "E"
            Case "1111"
                tempstr = tempstr & "F"
                
        End Select
    Next i

    DJ_BintoHex_16bit = tempstr
    
End Function
Public Function joe_Dec2Bin_8bit(DecValue As Long) As String
'16비트로 맞추어줌
On Error GoTo err
    
    Dim result As String
        
    result = ""
    Do
        result = CStr(DecValue Mod 2) & result
        DecValue = Fix(DecValue / 2)
    Loop Until DecValue = 0 Or DecValue = 1
    
    result = CStr(DecValue) & result
    
    If Len(result) > 16 Then
        result = Right(result, 8)
    ElseIf Len(result) < 8 Then
        result = "000000000" & result
        result = Right(result, 8)
    End If
    joe_Dec2Bin_8bit = result
Exit Function
err:
    joe_Dec2Bin_8bit = "err"
End Function

Public Function DJ_BintoHex_8bit(sBinary As String) As String  'binary 를 hex 로 바꿈
On Error Resume Next

Dim i As Integer
Dim tempstr As String
Dim tempBin As String
Dim tempLen As Integer
Dim tempBinary(1 To 4) As String
    
    tempBin = Left(sBinary & "0000000000000000", 8) '8비트
    
    For i = 1 To 2
        tempBinary(i) = Mid(tempBin, (i * 4) - 3, 4)
        Select Case tempBinary(i)
            Case "0000"
                tempstr = tempstr & "0"
            Case "0001"
                tempstr = tempstr & "1"
            Case "0010"
                tempstr = tempstr & "2"
            Case "0011"
                tempstr = tempstr & "3"
            Case "0100"
                tempstr = tempstr & "4"
            Case "0101"
                tempstr = tempstr & "5"
            Case "0110"
                tempstr = tempstr & "6"
            Case "0111"
                tempstr = tempstr & "7"
            Case "1000"
                tempstr = tempstr & "8"
            Case "1001"
                tempstr = tempstr & "9"
            Case "1010"
                tempstr = tempstr & "A"
            Case "1011"
                tempstr = tempstr & "B"
            Case "1100"
                tempstr = tempstr & "C"
            Case "1101"
                tempstr = tempstr & "D"
            Case "1110"
                tempstr = tempstr & "E"
            Case "1111"
                tempstr = tempstr & "F"
                
        End Select
    Next i

    DJ_BintoHex_8bit = tempstr
    
End Function

Public Function joe_Dec2Bin_16bit(DecValue As Long) As String
'16비트로 맞추어줌
On Error GoTo err
    
    Dim result As String
        
    result = ""
    Do
        result = CStr(DecValue Mod 2) & result
        DecValue = Fix(DecValue / 2)
    Loop Until DecValue = 0 Or DecValue = 1
    
    result = CStr(DecValue) & result
    
    If Len(result) > 16 Then
        result = Right(result, 16)
    ElseIf Len(result) < 16 Then
        result = "000000000000000" & result
        result = Right(result, 16)
    End If
    joe_Dec2Bin_16bit = result
Exit Function
err:
    joe_Dec2Bin_16bit = "err"
End Function

Public Function joe_Bin2Dec(BinValue As String) As Long
On Error GoTo err

    Dim i As Integer
    Dim result As Long
    
    For i = 1 To Len(BinValue)
        result = result + (Val(Mid(BinValue, Len(BinValue) - i + 1, 1)) * 2 ^ (i - 1))
    Next i
    joe_Bin2Dec = result
Exit Function
err:
    joe_Bin2Dec = -1
End Function
Public Function DJ_DataMake(TotalData As String, idataidx As Integer) As String '데이터를 조합해서 만듬
Dim i As Integer
Dim tempstr(0 To 4) As String
Dim temptest As String
Dim temp(0 To 99) As String
    temptest = "00000000000000000000000000000000000000000000000000000000000000000000000000000000"  '80자리
    TotalData = TotalData & temptest
    Select Case idataidx
    Case 0
        For i = 1 To 10
            temp(i) = Format((Mid(TotalData, (i * 4) - 3, 4)), "@@@@")
            temp(i) = DJ_DataReserve(temp(i))
            tempstr(idataidx) = tempstr(idataidx) & temp(i)
        Next i
    Case 1
        For i = 21 To 30
            temp(i) = Format((Mid(TotalData, (i * 4) - 3, 4)), "@@@@")
            temp(i) = DJ_DataReserve(temp(i))
            tempstr(idataidx) = tempstr(idataidx) & temp(i)
        Next i
    Case 2
        For i = 41 To 50
            temp(i) = Format((Mid(TotalData, (i * 4) - 3, 4)), "@@@@")
            temp(i) = DJ_DataReserve(temp(i))
            tempstr(idataidx) = tempstr(idataidx) & temp(i)
        Next i
    Case 3
        For i = 61 To 70
            temp(i) = Format((Mid(TotalData, (i * 4) - 3, 4)), "@@@@")
            temp(i) = DJ_DataReserve(temp(i))
            tempstr(idataidx) = tempstr(idataidx) & temp(i)
        Next i
    End Select
    DJ_DataMake = tempstr(idataidx)
End Function
Public Function DJ_DataReserve(Datastr As String) As String '한 주소안에 두자리 워드를 서로 뒤바꿈
Dim i As Integer
Dim tempstr As String
Dim temp(1 To 2) As String
    For i = 1 To 2
        temp(i) = Chr("&h" & (Mid(Datastr, (i * 2) - 1, 2)))
        If temp(i) = Chr("&h" & "00") Then
            temp(i) = " "
        End If
    Next i
    tempstr = temp(2) & temp(1)
    
    DJ_DataReserve = tempstr
End Function

Public Sub DJ_FavLineDraw(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double, lineColor As ColorConstants, frmindex As Integer)

Select Case frmindex
Case 0
'    frmMain.FavisImageBoxMain(0).color = lineColor
'    frmMain.FavisImageBoxMain(0).AddStaticLine X1, Y1, X2, Y2
'    frmMain.FavisImageBoxMain(0).UpdateDraw
Case 1
'    frmSetting.FavisImageBoxSetting.color = lineColor
'    frmSetting.FavisImageBoxSetting.AddStaticLine X1, Y1, X2, Y2
'    frmSetting.FavisImageBoxSetting.UpdateDraw
End Select
End Sub

Public Sub FormControlShow() '파일처리 함수에서 빼서 따로 함수로 만듬 (파일처리시 불필요한 폼 로드 발생 방지)
On Error Resume Next
Dim i As Integer

    Call LoadResultSaving(sModelName)
    Call SpecName_Load(sModelName)
    Call SpecAllValue_Load(sModelName)
    Call FunctionValue_Load(sModelName)

    For i = 0 To 9
        frmMain.txtSpecName(i).Text = sSpecName(i)
        frmMain.chkJudgement(i).Caption = sSpecName(i)
    Next i
    
    For i = 0 To 9
        'frmMain.txtOffset(i).Text = Format(dSpecOffset(i), "#0.00")
    Next i
    
    frmMain.chkResultImageSaving.Value = g_SaveResultImage

    frmMain.chkCamPass.Value = IIf(bCamPass, 1, 0)
    frmMain.chkOKImageSave.Value = IIf(bOKimageSave, 1, 0)
    frmMain.chkNGImageSave.Value = IIf(bNGimageSave, 1, 0)
    frmMain.chkWriteDataSave.Value = IIf(bWriteDataSave, 1, 0)
    frmMain.Option1(iImageFileMode - 1).Value = True
    
    frmMain.SSTab1.Tab = 0
    
End Sub
'이선행 대리 함수 참조
Public Sub SH_ScreenSave(Str As String, Optional Str2 As String = "", Optional Str3 As String = "") '스크린샷을 bmp 파일로 저장한다.
  
On Error Resume Next

'    frmMain.Image1.Picture = Nothing
'    frmMain.picScreenShotSave.Picture = Nothing
    'frmSetting.SetFocus
    Call Dlay_T(0.5)
    keybd_event VK_SNAPSHOT, 0, 0&, 0&      '1 이면 활성화면
    'keybd_event VK_SNAPSHOT, 0, 0&, 0&     '0 이면 전체화면
    DoEvents
    
    '여기서 화면 캡쳐
    frmMain.ImageScreen.Picture = Clipboard.GetData(vbCFBitmap)
    frmMain.picScreenShotSave.Picture = frmMain.ImageScreen.Picture
    
    'SavePicture frmMain.picScreenShotSave.Image, Str 'picturebox의 이미지를 저장
    
    Call SH_PictureSaveToJpg(Str) 'jpg 저장
    
    If Len(Dir$(Str)) = 0 Then
        Call SH_PictureSaveToJpg(Str3)
    End If
    
    If Str2 <> "" Then
        Call SH_PictureSaveToJpg(Str2)
    End If
    
End Sub
'이선행 대리 함수 참조
Public Sub SH_PictureSaveToJpg(Str As String) '픽쳐박스의 이미지를 jpg 로 저장한다.

On Error Resume Next

    Dim PicName As String
    Dim A As New cDIBSection
    
    PicName = Str
    
    A.CreateFromPicture frmMain.picScreenShotSave.Picture
    
    If SaveJPG(A, PicName) = False Then
        
        Exit Sub
    
    End If
        

End Sub
'이선행 대리 함수 참조
Public Sub SH_HDDCheking(Index As Integer)

On Error Resume Next

    Dim strDrive As String
    Dim lFreeBytesToCallers As Currency
    Dim lTotalBytes As Currency
    Dim lFreeBytes As Currency
    Dim lRetVal As Long
    
    Dim CapLen As Integer
    Dim ReadLen As Integer
    Dim UsedCap As String
    Dim UsedCapPerS As Double
    
    Select Case Index
    Case 0
        strDrive = "C:\"
    Case 1
        strDrive = "D:\"
    End Select
    lRetVal = GetDiskFreeSpaceEx(strDrive, lFreeBytesToCallers, lTotalBytes, lFreeBytes)
     
     UsedCap = CDbl(lTotalBytes * 10000) - CDbl(lFreeBytesToCallers * 10000)
     CapLen = Len(UsedCap)
     ReadLen = CapLen - 7
     
     UsedCapPerS = Int((UsedCap / CDbl(lTotalBytes * 10000)) * 100)
    
    frmMain.txtUsedCapPerS(Index).Text = UsedCapPerS

    If UsedCapPerS >= 50 And UsedCapPerS < 70 Then '사용용량이 알람 요구치를 넘어서면..
    
        'MsgBox "저장 HDD의 사용량이 경계치를 넘었습니다. 오래된 파일을 삭제하여 주십시요..!!", vbCritical
        'Over` = True
        '=================================================== 80%가 넘으면 체크를 강제로 없애고, 저장안함
        OverHdd = False
        frmMain.lblOverHdd(Index).BackColor = &H40C0&
        frmMain.lblOverHdd(Index).Caption = "경고"
        Call WriteMelsec(frmMain.ActEasyIF, sMelsecAddrAlarmHDD, 0)
    ElseIf UsedCapPerS >= 70 Then
        OverHdd = True
        frmMain.chkOKImageSave.Value = 0
        frmMain.chkNGImageSave.Value = 0
        bOKimageSave = False
        bNGimageSave = False
        frmMain.lblOverHdd(Index).BackColor = vbRed
        frmMain.lblOverHdd(Index).Caption = "하드정리"
        Call WriteMelsec(frmMain.ActEasyIF, sMelsecAddrAlarmHDD, 1)
    Else
        OverHdd = False
        frmMain.lblOverHdd(Index).BackColor = vbGreen
        frmMain.lblOverHdd(Index).Caption = "양호"
        Call WriteMelsec(frmMain.ActEasyIF, sMelsecAddrAlarmHDD, 0)
    End If
End Sub

Property Get RAD(ByVal aDeg As Double) As Double
    RAD = aDeg * (pi / 180#)
End Property

Property Get DEG(ByVal aRad As Double) As Double
    DEG = aRad * (180# / pi)
End Property

'-----------------------------------------------------------------------
'Arcsin으로 각을 구할때 입력값이 +1 , -1일때는 0으로 나누게 되는
'에러가 발생하므로, 이 두값에 대한 보정을 한다.
'-----------------------------------------------------------------------
Property Get Arcsin(ByVal X As Double) As Double
    If X = -1 Or X = 1 Then
        Arcsin = X * pi / 2#
    Else
        Arcsin = Atn(X / Sqr(-X * X + 1))
    End If
End Property

'-----------------------------------------------------------------------
'''두좌표의 각도 구하기 (인터넷 돌아다니는 함수 수정) DEJAY
'-----------------------------------------------------------------------
Function LineAngle(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double, ByRef dist As Double) As Double
    Dim radian As Double, degree As Double

    LineAngle = 0

    dist = Sqr((X1 - X2) * (X1 - X2) + (Y1 - Y2) * (Y1 - Y2)) '두점 사이의 거리

    '점일때는 0도를 리턴한다.
    If dist < 1 Then Exit Function

    radian = (Y1 - Y2) / dist   '라디안
    degree = DEG(Arcsin(radian))

    '2사분면, 3사분면의 처리
    If X1 > X2 Then degree = 180 - degree

    '0~360사이의 각으로 리턴하도록 음수각도 일때 360을 더해준다.
    If degree < 0 Then degree = degree
    LineAngle = degree
End Function

