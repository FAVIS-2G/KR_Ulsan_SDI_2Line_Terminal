Attribute VB_Name = "ModInspection"
Public Sub DistanceCaliper(Index As Integer, Tnum1 As Integer, Tnum2 As Integer)
Dim i As Integer
Dim dTemp As Double
Dim dTemp1 As Double
Dim dTemp2 As Double

Dim j As Integer

    dTemp1 = DJSJ_Point2Line(dCaliperX1(Index, Tnum1), dCaliperY1(Index, Tnum1), dCaliperX2(Index, Tnum1), dCaliperY2(Index, Tnum1), dCaliperCX(Index, Tnum2), dCaliperCY(Index, Tnum2))
    dTemp2 = DJSJ_Point2Line(dCaliperX1(Index, Tnum2), dCaliperY1(Index, Tnum2), dCaliperX2(Index, Tnum2), dCaliperY2(Index, Tnum2), dCaliperCX(Index, Tnum1), dCaliperCY(Index, Tnum1))
    
    dTemp = (dTemp1 + dTemp2) / 2
    
    If Tnum1 Mod 2 = 1 Then
        Exit Sub
    End If
    
    j = Tnum1 / 2
    dInspectResult_Pixel(Index, j) = dTemp
    
    If j < 2 Then
        dInspectResult_mm(Index, j) = Format((dInspectResult_Pixel(Index, j) * dCaliPX(Index)) + dSpecOffset((Index * 10) + j), "#00.00")
    Else
        dInspectResult_mm(Index, j) = Format((dInspectResult_Pixel(Index, j) * dCaliPXY(Index)) + dSpecOffset((Index * 10) + j), "#00.00")
    End If
    dTextPointX(Index, j) = (dCaliperCX(Index, Tnum1) + dCaliperCX(Index, Tnum1 + 1)) / 2 - 300
    
    If dTextPointX(Index, j) < 0 Then
        dTextPointX(Index, j) = 0
    End If
    
    dTextPointY(Index, j) = (dCaliperCY(Index, Tnum1) + dCaliperCY(Index, Tnum1 + 1)) / 2
'
'    Select Case Tnum1
'    Case 0
'        dInspectResult_Pixel(index, 0) = dTemp
'        dInspectResult_mm(index, 0) = Format((dInspectResult_Pixel(index, 0) * dCaliPX(index)) + dSpecOffset((index * 10) + 0), "#00.00")
'        dTextPointX(index, 0) = dCaliperCX(index, 0)
'        dTextPointY(index, 0) = dCaliperCY(index, 0)
'    Case 2
'        dInspectResult_Pixel(index, 1) = dTemp
'        dInspectResult_mm(index, 1) = Format((dInspectResult_Pixel(index, 1) * dCaliPX(index)) + dSpecOffset((index * 10) + 1), "#00.00")
'        dTextPointX(index, 1) = dCaliperCX(index, 2)
'        dTextPointY(index, 1) = dCaliperCY(index, 2)
'    Case 4
'        dInspectResult_Pixel(index, 2) = dTemp
'        dInspectResult_mm(index, 2) = Format((dInspectResult_Pixel(index, 2) * dCaliPX(index)) + dSpecOffset((index * 10) + 2), "#00.00")
'        dTextPointX(index, 2) = dCaliperCX(index, 4)
'        dTextPointY(index, 2) = dCaliperCY(index, 4)
'    Case 6
'        dInspectResult_Pixel(index, 3) = dTemp
'        dInspectResult_mm(index, 3) = Format((dInspectResult_Pixel(index, 3) * dCaliPX(index)) + dSpecOffset((index * 10) + 3), "#00.00")
'        dTextPointX(index, 3) = dCaliperCX(index, 6)
'        dTextPointY(index, 3) = dCaliperCY(index, 6)
'    Case 8
'        dInspectResult_Pixel(index, 4) = dTemp
'        dInspectResult_mm(index, 4) = Format((dInspectResult_Pixel(index, 4) * dCaliPX(index)) + dSpecOffset((index * 10) + 4), "#00.00")
'        dTextPointX(index, 4) = dCaliperCX(index, 8)
'        dTextPointY(index, 4) = dCaliperCY(index, 8)
'    Case 10
'        dInspectResult_Pixel(index, 5) = dTemp
'        dInspectResult_mm(index, 5) = Format((dInspectResult_Pixel(index, 5) * dCaliPX(index)) + dSpecOffset((index * 10) + 5), "#00.00")
'        dTextPointX(index, 5) = dCaliperCX(index, 10)
'        dTextPointY(index, 5) = dCaliperCY(index, 10)
'    Case 12
'        dInspectResult_Pixel(index, 6) = dTemp
'        dInspectResult_mm(index, 6) = Format((dInspectResult_Pixel(index, 6) * dCaliPX(index)) + dSpecOffset((index * 10) + 6), "#00.00")
'        dTextPointX(index, 6) = dCaliperCX(index, 12)
'        dTextPointY(index, 6) = dCaliperCY(index, 12)
'    Case 14
'        dInspectResult_Pixel(index, 7) = dTemp
'        dInspectResult_mm(index, 7) = Format((dInspectResult_Pixel(index, 7) * dCaliPX(index)) + dSpecOffset((index * 10) + 7), "#00.00")
'        dTextPointX(index, 7) = dCaliperCX(index, 14)
'        dTextPointY(index, 7) = dCaliperCY(index, 14)
'    Case 16
'        dInspectResult_Pixel(index, 8) = dTemp
'        dInspectResult_mm(index, 8) = Format((dInspectResult_Pixel(index, 8) * dCaliPX(index)) + dSpecOffset((index * 10) + 8), "#00.00")
'        dTextPointX(index, 8) = dCaliperCX(index, 16)
'        dTextPointY(index, 8) = dCaliperCY(index, 16)
'    Case 18
'        dInspectResult_Pixel(index, 9) = dTemp
'        dInspectResult_mm(index, 9) = Format((dInspectResult_Pixel(index, 9) * dCaliPX(index)) + dSpecOffset((index * 10) + 9), "#00.00")
'        dTextPointX(index, 9) = dCaliperCX(index, 18)
'        dTextPointY(index, 9) = dCaliperCY(index, 18)
'    End Select
    
            
End Sub
Public Sub testDistanceCaliper(Index As Integer)
Dim i As Integer
Dim dTemp(0 To 99) As Double
    For i = 1 To 5 Step 2
        dTemp(i - 1) = DJSJ_Point2Line(dCaliperCX(Index, 0), dCaliperCY(Index, 0), dCaliperCX(Index, 2), dCaliperCY(Index, 2), dCaliperCX(Index, i), dCaliperCY(Index, i))
    Next i
    For i = 7 To 11 Step 2
        dTemp(i - 1) = DJSJ_Point2Line(dCaliperCX(Index, 6), dCaliperCY(Index, 6), dCaliperCX(Index, 8), dCaliperCY(Index, 8), dCaliperCX(Index, i), dCaliperCY(Index, i))
    Next i

        dInspectResult_Pixel(Index, 0) = dTemp(0)
        dInspectResult_mm(Index, 0) = Format((dInspectResult_Pixel(Index, 0) * dCaliPX(Index)) + dSpecOffset((Index * 10) + 0), "#00.00")
        dTextPointX(Index, 0) = dCaliperCX(Index, 0)
        dTextPointY(Index, 0) = dCaliperCY(Index, 0)

        dInspectResult_Pixel(Index, 1) = dTemp(2)
        dInspectResult_mm(Index, 1) = Format((dInspectResult_Pixel(Index, 1) * dCaliPX(Index)) + dSpecOffset((Index * 10) + 1), "#00.00")
        dTextPointX(Index, 1) = dCaliperCX(Index, 2)
        dTextPointY(Index, 1) = dCaliperCY(Index, 2)

        dInspectResult_Pixel(Index, 2) = dTemp(4)
        dInspectResult_mm(Index, 2) = Format((dInspectResult_Pixel(Index, 2) * dCaliPX(Index)) + dSpecOffset((Index * 10) + 2), "#00.00")
        dTextPointX(Index, 2) = dCaliperCX(Index, 4)
        dTextPointY(Index, 2) = dCaliperCY(Index, 4)

        dInspectResult_Pixel(Index, 3) = dTemp(6)
        dInspectResult_mm(Index, 3) = Format((dInspectResult_Pixel(Index, 3) * dCaliPX(Index)) + dSpecOffset((Index * 10) + 3), "#00.00")
        dTextPointX(Index, 3) = dCaliperCX(Index, 7)
        dTextPointY(Index, 3) = dCaliperCY(Index, 7)

        dInspectResult_Pixel(Index, 4) = dTemp(8)
        dInspectResult_mm(Index, 4) = Format((dInspectResult_Pixel(Index, 4) * dCaliPX(Index)) + dSpecOffset((Index * 10) + 4), "#00.00")
        dTextPointX(Index, 4) = dCaliperCX(Index, 9)
        dTextPointY(Index, 4) = dCaliperCY(Index, 9)

        dInspectResult_Pixel(Index, 5) = dTemp(10)
        dInspectResult_mm(Index, 5) = Format((dInspectResult_Pixel(Index, 5) * dCaliPX(Index)) + dSpecOffset((Index * 10) + 5), "#00.00")
        dTextPointX(Index, 5) = dCaliperCX(Index, 11)
        dTextPointY(Index, 5) = dCaliperCY(Index, 11)

    
            
End Sub
Public Sub SpecCompare_Distance(Index As Integer, rstindex As Integer, frmindex As Integer)

    
    Select Case frmindex
        Case 0
            If dInspectResult_mm(Index, rstindex) >= dSpecOriMin(rstindex) And dInspectResult_mm(Index, rstindex) <= dSpecOriMax(rstindex) Then
                bResultJudge_Spec(Index, rstindex) = True
                frmMain.lblResultData(rstindex + (Index * 10)).BackColor = vbWhite
                frmMain.lblResultData(rstindex + (Index * 10)).Caption = dInspectResult_mm(Index, rstindex)
'                frmMain.FavisImageBoxMain(Index).color = vbGreen
'                frmMain.FavisImageBoxMain(Index).AddStaticText dTextPointX(Index, rstindex), dTextPointY(Index, rstindex), sSpecName(rstindex)
'                frmMain.FavisImageBoxMain(Index).AddStaticText dTextPointX(Index, rstindex), dTextPointY(Index, rstindex) + 100, "(" & dInspectResult_mm(Index, rstindex) & ")"
            Else
                If bSpecPass(rstindex) = False Then
                    bResultJudge_Spec(Index, rstindex) = False
                    bResultjudge_cnt = bResultjudge_cnt + 1
                    ispecFalse = 1
                    frmMain.lblResultData(rstindex + (Index * 10)).BackColor = vbRed
                    frmMain.lblResultData(rstindex + (Index * 10)).Caption = dInspectResult_mm(Index, rstindex)
                    If dTextPointX(Index, rstindex) <> 0 Then
'                        frmMain.FavisImageBoxMain(Index).color = vbRed
'                        frmMain.FavisImageBoxMain(Index).AddStaticText dTextPointX(Index, rstindex), dTextPointY(Index, rstindex), sSpecName(rstindex)
'                        frmMain.FavisImageBoxMain(Index).AddStaticText dTextPointX(Index, rstindex), dTextPointY(Index, rstindex) + 100, "(" & dInspectResult_mm(Index, rstindex) & ")"
                    End If
                Else
                    bResultJudge_Spec(Index, rstindex) = True
                    frmMain.lblResultData(rstindex + (Index * 10)).BackColor = vbWhite
                    frmMain.lblResultData(rstindex + (Index * 10)).Caption = dInspectResult_mm(Index, rstindex)
                    If dTextPointX(Index, rstindex) <> 0 Then
'                        frmMain.FavisImageBoxMain(Index).color = vbBlue
'                        frmMain.FavisImageBoxMain(Index).AddStaticText dTextPointX(Index, rstindex), dTextPointY(Index, rstindex), sSpecName(rstindex)
'                        frmMain.FavisImageBoxMain(Index).AddStaticText dTextPointX(Index, rstindex), dTextPointY(Index, rstindex) + 100, "(" & dInspectResult_mm(Index, rstindex) & ")"
                    End If
                End If
            End If
        Case 1
'            If dInspectResult_mm(Index, rstindex) >= dSpecOriMin(rstindex) And dInspectResult_mm(Index, rstindex) <= dSpecOriMax(rstindex) Then
'                bResultJudge_Spec(Index, rstindex) = True
'                frmSetting.lblResultData(rstindex).BackColor = vbWhite
'                frmSetting.lblResultData(rstindex).Caption = dInspectResult_mm(Index, rstindex)
'                If dTextPointX(Index, rstindex) <> 0 Then
'                    frmSetting.FavisImageBoxSetting.color = vbGreen
'                    frmSetting.FavisImageBoxSetting.AddStaticText dTextPointX(Index, rstindex), dTextPointY(Index, rstindex), sSpecName(rstindex)
'                    frmSetting.FavisImageBoxSetting.AddStaticText dTextPointX(Index, rstindex), dTextPointY(Index, rstindex) + 100, "(" & dInspectResult_mm(Index, rstindex) & ")"
'                End If
'            Else
'                bResultJudge_Spec(Index, rstindex) = False
'
'                frmSetting.lblResultData(rstindex).BackColor = vbRed
'                frmSetting.lblResultData(rstindex).Caption = dInspectResult_mm(Index, rstindex)
'                ispecFalse = 1
'                If dTextPointX(Index, rstindex) <> 0 Then
'                    frmSetting.FavisImageBoxSetting.color = vbRed
'                    frmSetting.FavisImageBoxSetting.AddStaticText dTextPointX(Index, rstindex), dTextPointY(Index, rstindex), sSpecName(rstindex)
'                    frmSetting.FavisImageBoxSetting.AddStaticText dTextPointX(Index, rstindex), dTextPointY(Index, rstindex) + 100, "(" & dInspectResult_mm(Index, rstindex) & ")"
'                End If
'            End If
    End Select
    
End Sub
Public Sub DJ_ManualRun(Index As Integer, frmindex As Integer)
Dim i As Integer
    
    '초기화---------------------------------
    ispecFalse = 0
    iResultJudge_BlobCnt = 0
    If frmindex = 0 Then
        Call Main_ImageBoxClear(Index)
    Else
        Call Setting_ImageBoxClear(Index)
    End If
    For i = 0 To 14
        dInspectResult_mm(Index, i) = 0
        dInspectResult_Pixel(Index, i) = 0
        bResultJudge_Spec(Index, i) = False
    Next i
    For i = 0 To 1
        Fix_ShiftPointX(Index) = 0
        Fix_ShiftPointY(Index) = 0
        Fix_ShiftPointAngle(Index) = 0
    Next i
    bResultJudge(Index) = False
    '---------------------------------------
    
    'Fixture Find --------------------------
    Select Case Fix_UseMode(Index)
    Case 0
    Case 1
        For i = 0 To 1
            Call CaliperFix_Region(Index, i, frmindex)
            Call CaliperFix_Find(Index, i, frmindex)
        Next i
        If frmindex = 0 Then
            Call Fixture_Run(Index)
        End If
    Case 2
        For i = 0 To 1
            Call CaliperFix_Region(Index, i, frmindex)
            Call CaliperFix_Find(Index, i, frmindex)
        Next i
        If frmindex = 0 Then
            Call FixtureAngle_Run(Index)
        End If
    End Select
    '---------------------------------------
    
    'Caliper Find---------------------------
    For i = 0 To iToolCount - 1
        Call CaliperT_Region(Index, i, frmindex)
        Call CaliperT_Find(Index, i, frmindex)
    Next i
    '---------------------------------------
    
    'Distance Caliper & SpecCompare ---------------------
    For i = 0 To iToolCount - 2 Step 2
        Call DistanceCaliper(Index, i, i + 1)
    Next i
'    Call testDistanceCaliper(Index)
    For i = 0 To (iToolCount / 2) - 1
        Call SpecCompare_Distance(Index, i, frmindex)
    Next i
    '----------------------------------------------------
    'Blob Find---------------------------
    For i = 0 To iBlobToolCount - 1
        iResultJudge_BlobCnt = iResultJudge_BlobCnt + BlobT_Find(Index, i, frmindex)
    Next i
    '---------------------------------------
    If ispecFalse = 0 And iResultJudge_BlobCnt = 0 Then
        bResultJudge(Index) = True
    Else
        bResultJudge(Index) = False
    End If
    
End Sub

Public Sub JobAndSpecChange()
    '양승조 추가(잡 체인지 //////////////////////
    If Communication_Interface("read", 13530, "JobChange") = True Then
        Call Dlay_T(0.2)
        If bJobChangeOn = True Then
            '모델체인지
            frmModelAuto.Show
            TempClickModelNo = nPreJobNum
            TempClickModelName = frmModelAuto.cmdModelRoom(nPreJobNum).Caption
            If TempClickModelName = "" Or TempClickModelName = "Empty" Or TempClickModelNo = 0 Then
                MsgBox "없는 모델입니다. 로드할 수 없습니다.", vbCritical, "모델 불러오기 실패"
            Else
                loadsw = ModelData_FileLoad(Trim(TempClickModelName))     '모델로딩(리턴값 불리언)
                If loadsw = True Then           '모델로딩이 정상적으로 되었다면 모델이 존재 하므로
                    sModelName = Trim(Modelinfo.ModelName)
                    'Call SealPin_ModelLoad(sModelName)
                    Call FormControlShow
                    Call LastModelWrite             '마지막 작업명으로 저장
                    frmMain.lblModelNameMain.Caption = sModelName
                End If
                If loadsw = False Then
                    MsgBox "모델 로드에 실패 했습니다.", vbCritical, "모델 불러오기 실패"
                    Call LOGWrite("번호[" & TempClickModelNo & "]" & TempClickModelName & "모델 로드중 자동검사에 필요한 정보를 로드하지 못하였습니다. 모델 로드 작업이 중단 되었습니다.")
                    Exit Sub
                End If
                iNowModelNo = TempClickModelNo
                sModelName = Trim(Modelinfo.ModelName)
                Unload frmModelAuto
                Call LOGWrite("번호[" & TempClickModelNo & "]" & TempClickModelName & "모델 로드작업을 정상적으로 수행 하였습니다. 현재 부터 번호[" & TempClickModelNo & "]" & TempClickModelName & "로 적용 됩니다.")
            End If
            If Communication_Interface("write", 4092, "0001") = True Then
                Dlay_T (0.2)
            End If
        End If
    End If
    
    '양승조 추가(잡 체인지 //////////////////////end
    
    
    '양승조 추가(스펙받아오기//////////////////////
    If Communication_Interface("read", 4092, "SpecChange") = True Then
        Call Dlay_T(0.2)
        If bSpecChangeOn = True Then
            If iProName = 0 Then
                If Communication_Interface("readData", 13660, "ReadSpecAlg") = True Then
                    Call Dlay_T(0.5)
                    Call FormControlShow
                End If
                
                '스펙이 변경되었다는 것을 알리는 주소 0으로 초기화
                If Communication_Interface("write", 4092, "0000") = True Then
                    Dlay_T (0.2)
                End If
            ElseIf iProName = 1 Then
                If Communication_Interface("readData", 13600, "ReadSpecTer") = True Then
                    Call Dlay_T(0.5)
                    Call FormControlShow
                End If
                
                '스펙이 변경되었다는 것을 알리는 주소 0으로 초기화
                If Communication_Interface("write", 4092, "0000") = True Then
                    Dlay_T (0.2)
                End If
            End If
        End If
    End If
    '양승조 추가(스펙받아오기//////////////////////end
End Sub

Public Sub JobAndSpecChangeQ71()
    '양승조 추가(잡체인지)///////////////////////////////////
'     Dim ln_tempJobNum As Integer
'     ln_tempJobNum = QJ71E71ReadData("20307", CLng(1), 0)
'     If nPreJobNum <> ln_tempJobNum Then
'         '모델체인지
'         frmModelAuto.Show
'         TempClickModelName = frmModelAuto.cmdModelRoom(ln_tempJobNum).Caption
'         TempClickModelNo = ln_tempJobNum
'         If TempClickModelName = "" Or TempClickModelName = "Empty" Or TempClickModelNo = 0 Then
'             MsgBox "없는 모델입니다. 로드할 수 없습니다.", vbCritical, "모델 불러오기 실패"
'         Else
'             loadsw = ModelData_FileLoad(Trim(TempClickModelName))     '모델로딩(리턴값 불리언)
'             If loadsw = True Then           '모델로딩이 정상적으로 되었다면 모델이 존재 하므로
'                 sModelName = Trim(Modelinfo.ModelName)
'                 Call DJ_ModelLoad(sModelName)
'                 Call FormControlShow
'                 Call LastModelWrite             '마지막 작업명으로 저장
'                 frmMain.lblModelNameMain.Caption = sModelName
'             End If
'             If loadsw = False Then
'                 MsgBox "모델 로드에 실패 했습니다.", vbCritical, "모델 불러오기 실패"
'                 Call LOGWrite("번호[" & TempClickModelNo & "]" & TempClickModelName & "모델 로드중 자동검사에 필요한 정보를 로드하지 못하였습니다. 모델 로드 작업이 중단 되었습니다.")
'                 Exit Sub
'             End If
'             iNowModelNo = TempClickModelNo
'             sModelName = Trim(Modelinfo.ModelName)
'             Unload frmModelAuto
'             Call LOGWrite("번호[" & TempClickModelNo & "]" & TempClickModelName & "모델 로드작업을 정상적으로 수행 하였습니다. 현재 부터 번호[" & TempClickModelNo & "]" & TempClickModelName & "로 적용 됩니다.")
'         End If
'         nPreJobNum = ln_tempJobNum
'         Call QJ71E71WriteData("30004", 1)
'         Unload frmModelAuto
'     End If
'     '양승조 추가(잡체인지)///////////////////////////////////end
'
'     '==========================스펙 받기===========================
'     If QJ71E71ReadData("30004", CLng(1), 0) = 1 Then
'         Dim specAddr As Integer
'
'         For specAddr = 0 To (iToolCount / 2) - 1
'             Dim tmpSpecData As Long
'
'             tmpSpecData = QJ71E71ReadData(CStr(13650 + (specAddr * 3)), 1, 0)    '기준값
'             dSpecOri(specAddr) = tmpSpecData * 0.01
'             tmpSpecData = QJ71E71ReadData(CStr(13650 + (specAddr * 3) + 1), 1, 0)  '공차 +
'             dSpecMax(specAddr) = tmpSpecData * 0.01
'             tmpSpecData = QJ71E71ReadData(CStr(13650 + (specAddr * 3) + 2), 1, 0)   '공차 -
'             dSpecMin(specAddr) = tmpSpecData * 0.01
'
'             dSpecOriMin(specAddr) = dSpecOri(specAddr) - dSpecMin(specAddr)
'             dSpecOriMax(specAddr) = dSpecOri(specAddr) + dSpecMax(specAddr)
'         Next specAddr
'
'
'         Call Grid_Init
'         Call FormControlShow
'         Call QJ71E71WriteData("30004", CLng(0))
'     End If

     '==========================스펙 받기===========================END
End Sub
'---------------------------------------------------------------------------------------
' Procedure : DJ_AutoRun_Q71E71
' DateTime  : 2013-01-11 18:36
' Author    : Administrator
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub DJ_AutoRun_Q71E71()
On Error GoTo CommErr:
Dim bRet As Boolean
Dim tempstr As String
Dim ImageFolderName As String
Dim ImageFolderName2 As String
Dim sMesSendJPGPath As String
Dim i As Integer
Dim sDate As String
Dim stime As String
Dim sMESDate As String
Dim sMesTime As String
Dim sTempCode As String
Dim starttime As Long
Dim endtime As Long
Dim sDataTemp As String
Dim ltemp(0 To 99) As Long

    Do While frmMain.ActEasyIF.Open = 0 And bAutoRunOn = True
        Sleep 1000
    Loop
        
    Do
        DoEvents
                    
        Do While m_Rcv_Bit(1) = 0 And m_Rcv_Bit(0) = 0
            DoEvents
            If bAutoRunOn = False Then
                Exit Sub
            End If
            Call ClearMelsecResult(addVisionInspect1)
            m_Snd_Bit_1(outReadyVision) = 1
            Call Write_Send_Word(addVisionInspect1, Make_Send_Word(addVisionInspect1, True))
            Sleep 10
        Loop
        
        starttime = GetTickCount
        
        '검사중 신호 전송
        m_Snd_Bit_1(outBusyVision) = 1
        Call Write_Send_Word(addVisionInspect1, Make_Send_Word(addVisionInspect1, True))
        
        
        For i = 0 To kMaxCamera - 1
           Set g_CogImage(i) = IDS_AcquireCognex(frmMain.uEyeCam1(i), frmMain.CogDisplay(i))
           'Call SealPin_ManualRun(i, frmMain.CogDisplay(i))
        Next i
        
        
        '검사종료 및 OK NG 신호 전송
        m_Snd_Bit_1(outBusyVision) = 0
        m_Snd_Bit_1(outEndVision) = 1
        Call Write_Send_Word(addVisionInspect1, Make_Send_Word(addVisionInspect1, True))
        
        endtime = GetTickCount
        
        frmMain.lblInspecTime.Caption = CStr(endtime - starttime)
        Sleep 10
    
    Loop Until bAutoRunOn = False
    
    If 0 Then
    'If m_Snd_Bit_1(outReadyVision) = 1 Then
        
        ListBox_Append Time & "  Vision Ready 신호 전송완료", 0
        Do
            DoEvents
            
            If m_Rcv_Bit(1) = 1 Then ' 모델 체인지 트리거 들어옴

                    strDeviceList = "D36010"
                    nSize = 1
                    ReDim nData_d(nSize)
                    nResult = frmMain.ActEasyIF.ReadDeviceRandom(strDeviceList, nSize, nData_d(0))
                    If nResult = 0 Then
                        ModelIndex = nData_d(0)
                    End If

                    TempClickModelNo = ModelIndex                                    '모델창에서 버튼에의해 선택된 모델 번호
                    TempClickModelName = Trim(frmModelAuto.cmdModelRoom(ModelIndex).Caption)       '모델창에서 버튼에의해 선택된 모델 이름
                    For i = 1 To 100
                        frmModelAuto.cmdModelRoom(i).BackColor = vbWhite
                    Next i
                     frmModelAuto.cmdModelRoom(ModelIndex).BackColor = vbGreen

                     loadsw = ModelData_FileLoad(Trim(TempClickModelName))     '모델로딩(리턴값 불리언)
                    If loadsw = True Then           '모델로딩이 정상적으로 되었다면 모델이 존재 하므로
                        sModelName = Trim(Modelinfo.ModelName)
                        'Call SealPin_ModelLoad(sModelName)
                        Call FormControlShow
                        Call LastModelWrite             '마지막 작업명으로 저장
                        frmMain.lblModelNameMain.Caption = sModelName + "[" + CStr(TempClickModelNo) + "]"
                    End If
                    If loadsw = False Then
                        MsgBox "모델 로드에 실패 했습니다.", vbCritical, "모델 불러오기 실패"
                        Call LOGWrite("번호[" & TempClickModelNo & "]" & TempClickModelName & "모델 로드중 자동검사에 필요한 정보를 로드하지 못하였습니다. 모델 로드 작업이 중단 되었습니다.")
                        Exit Sub
                    End If
                    iNowModelNo = TempClickModelNo
                    sModelName = Trim(Modelinfo.ModelName)
                   ' MsgBox "선택하신 [" & TempClickModelNo & "]" & TempClickModelName & " 모델을 정상적으로 로드 했습니다.", vbInformation, "모델 로드 성공"
                    Call LOGWrite("번호[" & TempClickModelNo & "]" & TempClickModelName & "모델 로드작업을 정상적으로 수행 하였습니다. 현재 부터 번호[" & TempClickModelNo & "]" & TempClickModelName & "로 적용 됩니다.")
                    nPreJobNum = TempClickModelNo

            End If
            
            If m_Rcv_Bit(0) = 1 Then ' 트리거 들어옴
                    '초기화 후 비지 신호
                    'PLC -> PC ======================================================
                    Call ReadDataFromPLC
                    'PLC -> PC ================================================== END
                    
                    Call ClearMelsecResult(addVisionInspect1)
                    m_Snd_Bit_1(outBusyVision) = 1
                    Call Write_Send_Word(addVisionInspect1, Make_Send_Word(addVisionInspect1, True))
                

                    sDate = Format(Date, "yy-mm-dd")
                    stime = Format(Time, "hh-mm-ss")
                    sMESDate = Format(Date, "YYYYMMDD")
                    sMesTime = Format(Time, "HHMMSS")
                    sDateTimeCheck = sMESDate & sMesTime
                    ImageFolderName = "D:\Imagesave\"
                    Call Create_DIR(ImageFolderName)
                    ImageFolderName = "D:\Imagesave\" & sDate & "\"
                    Call Create_DIR(ImageFolderName)
                    ImageFolderName = "D:\Imagesave\" & sDate & "\" & sModelName & "\"
                    Call Create_DIR(ImageFolderName)
                    Call Create_DIR("D:\VisionImage\")
    
                    '조명켜기
                    
                    Call LightControl(0, True)
                    Call LightControl(1, True)
                    Call LightControl(2, True)
                    Call LightControl(3, True)
                    starttime = GetTickCount
                    '바코드 아이디 리드
                    
                    '''''''''''''''''''''''''''''
    
                    '아이디 코드에 쓰레기 값을 지우기 위해 라벨을 이용했다
                    For i = 0 To 3
'                      frmMain.lblIDCodeNum(i).Caption = sIDCode(i)
                      If Asc(Left(sIDCode(i), 1)) = 0 Then
                          sIDCode(i) = "NOID"
                      Else
                         ' sIDCode(i) = frmMain.lbl_IDcodeCleaner.Caption
                      End If
                    
                      frmMain.lblIDCodeNum(i).Caption = sIDCode(i)
                      frmMain.lblIDCode.Caption = sIDCode(i)
                    Next i
                    '양승조 추가///////////////////////////////////////////////////////////
                    '기종이 변경되었거나 스펙이 변경되었을 때 plc의 지정된 주소에서 받아온다
                   ' Call JobAndSpecChangeQ71
                    
                    For i = 0 To 3
                        If dCelluse(i) = 1 Then
                            '검사부분 =============================================================================
                            If bCamPass = False Then
                                Set g_CogImage(i) = IDS_AcquireCognex(frmMain.uEyeCam1(i), frmMain.CogDisplay(i))
                                'Call SealPin_ManualRun(i, frmMain.CogDisplay(i))
                                If bResultJudge(i) = True And sIDCode(i) <> "NOID" Then
                                    'ok 신호
                                    m_Snd_Bit_1(outOk1cam + (i * 2)) = 1
                                    tempstr = "OK"
                                    lOKCount = lOKCount + 1
                                    lToTalCount = lToTalCount + 1
                                    frmMain.lblCountOK.Caption = lOKCount
                                    frmMain.lblCountTotal.Caption = lToTalCount
                                    Call Dlay_T(0.02)
                                Else
                                    'NG' 전송
                                    m_Snd_Bit_1(outNG1cam + (i * 2)) = 1
                                    tempstr = "NG"
                                    lNGCount = lNGCount + 1
                                    lToTalCount = lToTalCount + 1
                                    frmMain.lblCountNG.Caption = lNGCount
                                    frmMain.lblCountTotal.Caption = lToTalCount
                                    Call Dlay_T(0.02)
                                End If
                            Else
                                Set g_CogImage(i) = IDS_AcquireCognex(frmMain.uEyeCam1(i), frmMain.CogDisplay(i))
                                m_Snd_Bit_1(outOk1cam + (i * 2)) = 1
                               '카메라 패스면 OK 아니면 NG
                            End If
                        Else
                            Set g_CogImage(i) = IDS_AcquireCognex(frmMain.uEyeCam1(i), frmMain.CogDisplay(i))
                        End If
                            '======================================================================================
                   Next i
                  
                    Call LightControl(0, False)
                    Call LightControl(1, False)
                    Call LightControl(2, False)
                    Call LightControl(3, False)
                    
                     ' 결과 데이터 ==========================================================
                     Call WriteDataToPLC
                     ' 결과 데이터 ====================================================== END
                    
                    '엔드 및 OK NG 신호 전송
                    m_Snd_Bit_1(outBusyVision) = 0
                    m_Snd_Bit_1(outEndVision) = 1
                    Call Write_Send_Word(addVisionInspect1, Make_Send_Word(addVisionInspect1, True))
                    
                    Call DJ_CountSave '카운트 저장
                    Call DJ_CountLoad '카운트 저장
                    
                For i = 0 To 3
                    'Write Grid Data------------------------
                    If dCelluse(i) = 1 Then
                        Call WriteData_Grid(i)
                    Else
                    End If
                    '---------------------------------------
                Next i
                For i = 0 To 3
                    If dCelluse(i) = 1 Then
                        '판정 후 이미지 저장 및 데이터 저장 ========================================덕화
                        If bResultJudge(i) = True And sIDCode(i) <> "NOID" Then
                            'frmMain.FavisImageBoxMain(Index).color = vbGreen
                            frmMain.lblResults.Caption = "O.K"
                            'frmMain.ShpResult.BackColor = &H8000&
'                            frmMain.camresult(i * 2).BackColor = &H8000&
'                            frmMain.camresult((i * 2) + 1).BackColor = &H8000&
'                            frmMain.FavisImageBoxMain(i).AddStaticText 2200, 50, "OK"
                            If bOKimageSave = True Then
                                Call DJ_ImageSave(i, "D:\VisionImage\", "OK", iImageFileMode)
                            End If
                            
                        Else
'                            frmMain.FavisImageBoxMain(i).color = vbRed
                            frmMain.lblResults.Caption = "N.G"
'                            frmMain.ShpResult.BackColor = vbRed
'                            frmMain.camresult(i * 2).BackColor = vbRed
'                            frmMain.camresult((i * 2) + 1).BackColor = vbRed
'                            frmMain.FavisImageBoxMain(i).AddStaticText 200, 50, "NG"
                            If bNGimageSave = True Then
                                Call DJ_ImageSave(i, "D:\VisionImage\", "NG", iImageFileMode)
                            End If
                            Call MES_DATASEND_FUNC("NG_PRODUCT_EVENT", "", "")          '불량코드 전송
                        End If
                        
                        If bWriteDataSave = True Then
                            Call ResultWriteOpen(i, tempstr, ImageFolderName)
                        Else
                            'lInspectionNum = lInspectionNum + 1
                        End If
                    Else
                    End If
                Next i
                
                    '''''''''
                Call Dlay_T(0.01)
                sMesSendJPGPath = "d:\MES\SEND\" & sIDCode(0) & "^" & sIDCode(1) & "^" & sIDCode(2) & "^" & sIDCode(3) & "_" & sMESEquipCode & "_" & 1 & "_" & sDateTimeCheck & "_001" & ".JPG"
                Call Dlay_T(0.1)
                Call SH_ScreenSave(sMesSendJPGPath)
                Call SH_ScreenSave(ImageFolderName & sIDCode(0) & "^" & sIDCode(1) & "^" & sIDCode(2) & "^" & sIDCode(3) & "_" & sMESEquipCode & "_" & 1 & "_" & sDateTimeCheck & "_001" & ".JPG")
                Call Dlay_T(0.01)
               
    

                    
                    '===========================================================================
                For i = 0 To 3
                
                    sDataTemp = DJ_DataFileADD(i)
                    Call DataFileSave(i, sDataTemp, "d:\MES\SEND\" & sIDCode(i) & "_" & sMESEquipCode & "_" & i + 1 & "_" & sDateTimeCheck & ".QCP")     'MES 에 전송할 데이터 생성
                    Call DataFileSave(i, sDataTemp, ImageFolderName & sIDCode(i) & "_" & sMESEquipCode & "_" & i + 1 & "_" & sDateTimeCheck & ".QCP")     '저장되는 데이터 생성
                Next i
                
                Call MES_NetDriveConnect
                
                    
                    For i = 0 To 1
                        Call SH_HDDCheking(i)
                    Next i
                endtime = GetTickCount
                frmMain.lblInspecTime.Caption = endtime - starttime
            Else
                bTriggerOn = False
                
            End If
            
        Loop Until bAutoRunOn = False
    End If
    Call Dlay_T(0.15) '딜레이 안주면 Melsec read 에서 Data 잘 못가져 옴
    
Exit Sub

CommErr:
    
    
    MsgBox "PLC 와 통신이 끊어졌습니다. 잠시후 재접속 하십시오.", vbCritical, "PLC 통신 확인"
    frmMain.shpAutoStop.BackColor = &H40C0&
    frmMain.lblAutoStop.Caption = "정지상태"
    frmMain.BHBLive.Enabled = True
    frmMain.BHBManualRun.Enabled = True
    frmMain.BHBModel.Enabled = True
    frmMain.BHBSetting.Enabled = True
    frmMain.BHBEnd.Enabled = True
    frmMain.BHBAutoRun.Enabled = True
    frmMain.BHBStop.Enabled = True
    bAutoRunOn = False
End Sub

Public Sub WriteData_Grid(Index As Integer)
Dim i As Integer
Dim Rownum As Long
Dim tempstr As String

    If bResultJudge(Index) = True And sIDCode(Index) <> "NOID" Then
        tempstr = "OK"
    Else
        tempstr = "NG"
    End If
    
    Rownum = frmMain.MSFlexGrid1.Rows
    If frmMain.MSFlexGrid1.Rows >= 3001 Then
        frmMain.MSFlexGrid1.Clear
        frmMain.MSFlexGrid1.Rows = 1
        Rownum = 1
        Call Grid_Init
        Rownum = Rownum + 1
    Else
        Rownum = Rownum + 1
    End If
        
    frmMain.MSFlexGrid1.Rows = Rownum

''''    For i = 0 To 3
''''        sIDCode(i) = "AGBDFSDE"
''''    Next i
    Dim tmpColN As Integer
    Dim tmpColN2 As Integer
   ' Select Case Index
        'Case 0
            frmMain.MSFlexGrid1.TextMatrix(Rownum - 1, 0) = lInspectionNum
            frmMain.MSFlexGrid1.TextMatrix(Rownum - 1, 1) = Time
            frmMain.MSFlexGrid1.TextMatrix(Rownum - 1, 2) = sIDCode(Index)
            If iProName = 0 Then
                frmMain.MSFlexGrid1.TextMatrix(Rownum - 1, 3) = sIDCode(1)
                frmMain.MSFlexGrid1.TextMatrix(Rownum - 1, 4) = sIDCode(2)
                frmMain.MSFlexGrid1.TextMatrix(Rownum - 1, 5) = sIDCode(3)
                frmMain.MSFlexGrid1.TextMatrix(Rownum - 1, 6) = tempstr
                tmpColN = 6
                
            Else
                frmMain.MSFlexGrid1.TextMatrix(Rownum - 1, 3) = tempstr
                tmpColN = 3
            End If
            For i = 1 To iToolCount / 2
                frmMain.MSFlexGrid1.TextMatrix(Rownum - 1, i + tmpColN) = dInspectResult_mm(Index, i - 1)
                tmpColN2 = tmpColN + i
            Next i
            
            For i = 1 To iBlobToolCount
                If iResultJudge_BlobCnt > 0 Then
                    frmMain.MSFlexGrid1.TextMatrix(Rownum - 1, i + tmpColN2) = "NG"
                Else
                    frmMain.MSFlexGrid1.TextMatrix(Rownum - 1, i + tmpColN2) = "OK"
                End If
            Next i
            
            frmMain.MSFlexGrid1.Col = 0
            frmMain.MSFlexGrid1.Row = Rownum - 1
            frmMain.MSFlexGrid1.CellForeColor = vbWhite
            lInspectionNum = lInspectionNum + 1
        'Case 1
   ' End Select
    
    For i = 0 To (iToolCount / 2) - 1
        If bResultJudge_Spec(Index, i) = False Then
            frmMain.MSFlexGrid1.Col = i + tmpColN + 1
            frmMain.MSFlexGrid1.Row = Rownum - 1
            frmMain.MSFlexGrid1.CellForeColor = vbRed
        End If
    Next i
    If bResultJudge(Index) = False Then
        frmMain.MSFlexGrid1.Col = tmpColN
        frmMain.MSFlexGrid1.Row = Rownum - 1
        frmMain.MSFlexGrid1.CellForeColor = vbRed
    End If
    
    For i = 0 To iBlobToolCount - 1
        If iResultJudge_BlobCnt > 0 Then
            frmMain.MSFlexGrid1.Col = i + tmpColN2 + 1
            frmMain.MSFlexGrid1.Row = Rownum - 1
            frmMain.MSFlexGrid1.CellForeColor = vbRed
        End If
    Next i
    frmMain.MSFlexGrid1.Row = 1
    frmMain.MSFlexGrid1.Col = 0
    frmMain.MSFlexGrid1.Sort = 4
End Sub

Public Sub DJ_ImageSave(Index As Integer, dirPath As String, rstJudge As String, filemode As Integer)
Dim i As Integer
Dim stime As String
Dim sDate As String
    stime = Format(Time, "hh-mm-ss")
    sDate = Format(Date, "yy-mm-dd")
    frmMain.mstr_PathVisionImg = dirPath
    
    Select Case filemode
    Case 1            'bmp
        favImageFileT(Index).FileType = 1
        favImageFileT(Index).ImageWidth = XRES
        favImageFileT(Index).ImageHeight = YRES
        favImageFileT(Index).Write frmMain.mstr_PathVisionImg & stime & "_" & sIDCode(Index) & ".bmp", fvImageBuf(iCamNumberS)
    Case 2            'jpg
        favImageFileT(Index).FileType = 2
        favImageFileT(Index).ImageWidth = XRES
        favImageFileT(Index).ImageHeight = YRES
        favImageFileT(Index).Write frmMain.mstr_PathVisionImg & stime & "_" & sIDCode(Index) & ".jpg", fvImageBuf(iCamNumberS)
        favImageFileT(Index).FileType = 1
    End Select
End Sub

Public Function ResultWriteOpen(Index As Integer, Str As String, modelpath As String)

On Error GoTo err

Dim i As Integer
Dim SHDate As String
Dim SHTime As String
Dim cam1(0 To 10) As String
Dim sSpecTemp As String
Dim sSpecDTemp As String
Dim sDataTemp As String
Dim tempLen As Integer

    SHDate = Format(Date, "yy-mm-dd")
    SHTime = Format(Time, "hh:mm:ss")
    
    FileName_Result = modelpath & SHDate & ".csv"
    FileNumber_Result = FreeFile

    Open FileName_Result For Append As FileNumber_Result
        For i = 0 To (iToolCount / 2) - 1
            If bResultJudge_Spec(Index, i) = True Then
                cam1(i) = Format(dInspectResult_mm(Index, i), "#00.00")
            Else
                cam1(i) = "<" & Format(dInspectResult_mm(Index, i), "#00.00") & ">"
            End If
            sSpecTemp = sSpecTemp & sSpecName(i) & ","
            sSpecDTemp = sSpecDTemp & dSpecOriMin(i) & " ~ " & dSpecOriMax(i) & ","
            sDataTemp = sDataTemp & cam1(i) & ","
        Next i

        If lInspectionNum = 0 Then
            Print #FileNumber_Result, "검사개수"; ","; "시간"; ","; "ID_CODE"; ","; "판정"; ","; sSpecTemp; "진행모델"
                                    
            Print #FileNumber_Result, "    "; ","; "        "; ","; "        "; ","; "        "; ","; sSpecDTemp; ""
            lInspectionNum = 1
        End If
            
        Print #FileNumber_Result, lInspectionNum; ","; SHTime; ","; sIDCode(Index); ","; Str; ","; sDataTemp; sModelName; ""
            'lInspectionNum = lInspectionNum + 1

    Close #FileNumber_Result
    

Exit Function

err:

Close #FileNumber_Result

End Function

