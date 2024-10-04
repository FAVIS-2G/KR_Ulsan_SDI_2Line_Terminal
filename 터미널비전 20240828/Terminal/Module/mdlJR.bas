Attribute VB_Name = "mdlJR"
Public Const kMaxCamera As Integer = 4
Public Const kMaxTool As Integer = 16

Public g_CogImage(0 To kMaxCamera - 1) As New CogImage8Grey

Public g_CogCalibrationTool(0 To 1) As CogCaliperTool
Public g_CogCalibrationRegion(0 To 1) As CogRectangleAffine

Public g_CogCaliperTool(0 To kMaxCamera - 1, 0 To kMaxTool - 1) As CogCaliperTool
Public g_CogCaliperRegion(0 To kMaxCamera - 1, 0 To kMaxTool - 1) As CogRectangleAffine
Public g_CogCaliperScorer As CogCaliperScorerPositionNeg
Public g_CogCaliperScorerPosition As CogCaliperScorerPosition

Public g_CogFindLineTool(0 To kMaxCamera - 1, 0 To kMaxTool - 1) As New CogFindLineTool
Public g_CogFindLineSegment(0 To kMaxCamera - 1, 0 To kMaxTool - 1) As New CogLineSegment

Public g_Distance(0 To kMaxCamera - 1, 0 To 3) As Double

Public Function InitCogTool() As Boolean
On Error Resume Next

    Dim CaliperTool As CogCaliperTool
    Dim CaliperRegion As CogRectangleAffine
    Dim CaliperScorer As CogCaliperScorerPositionNeg
    
    Dim i, j As Integer
    
    Set g_CogCaliperScorer = New CogCaliperScorerPositionNeg
    g_CogCaliperScorer.SetXYParameters -100, 100, 10000, 1, 0
    
    Set g_CogCaliperScorerPosition = New CogCaliperScorerPosition
    g_CogCaliperScorerPosition.SetXYParameters 0, 100, 10000, 1, 0
    
    For i = 0 To 1
        Set g_CogCalibrationTool(i) = New CogCaliperTool
        Set g_CogCalibrationRegion(i) = New CogRectangleAffine
        Set g_CogCalibrationTool(i).Region = g_CogCalibrationRegion(i)
        
        With g_CogCalibrationTool(i).RunParams
            .ContrastThreshold = 10
            .Edge0Polarity = cogCaliperPolarityDontCare
            .EdgeMode = cogCaliperEdgeModeSingle
            .FilterHalfSizeInPixels = 3
            .MaxResults = 1
            .SingleEdgeScorers.Clear
            .SingleEdgeScorers.Add g_CogCaliperScorer
        End With
        
        With g_CogCalibrationRegion(i)
            .SetCenterLengthsRotationSkew XRES / 4 * (1 + 2 * i), YRES / 4 * 1, XRES / 4, 100, CogMisc.DegToRad(i * 180), 0
            .GraphicDOFEnable = cogRectangleAffineDOFAll
            .XDirectionAdornment = cogRectangleAffineDirectionAdornmentSolidArrow
            .YDirectionAdornment = cogRectangleAffineDirectionAdornmentArrow
            .Color = cogColorGreen
            .Interactive = True
        End With
    Next i

    For i = 0 To kMaxCamera - 1
    
        For j = 0 To kMaxTool - 1
            Set g_CogCaliperTool(i, j) = New CogCaliperTool
            Set g_CogCaliperRegion(i, j) = New CogRectangleAffine
            Set g_CogCaliperTool(i, j).Region = g_CogCaliperRegion(i, j)
            
            With g_CogCaliperTool(i, j).RunParams
                .ContrastThreshold = 10
                .Edge0Polarity = cogCaliperPolarityDarkToLight
                .EdgeMode = cogCaliperEdgeModeSingle
                .FilterHalfSizeInPixels = 3
                .MaxResults = 1
                .SingleEdgeScorers.Clear
                If j = 2 Or j = 3 Or j = 6 Or j = 7 Then
                    .SingleEdgeScorers.Add g_CogCaliperScorerPosition
                Else
                    .SingleEdgeScorers.Add g_CogCaliperScorer
                End If
            End With
            
            With g_CogCaliperRegion(i, j)
                .SetCenterLengthsRotationSkew XRES / 4 * (1 + 2 * i), YRES / 4 * 1, XRES / 4, 100, CogMisc.DegToRad(i * 180), 0
                .GraphicDOFEnable = cogRectangleAffineDOFAll
                .XDirectionAdornment = cogRectangleAffineDirectionAdornmentSolidArrow
                .YDirectionAdornment = cogRectangleAffineDirectionAdornmentArrow
                .Color = cogColorGreen
                .Interactive = True
            End With
            
        Next j
        
        Call g_CogCaliperRegion(i, 0).SetCenterLengthsRotationSkew(XRES / 4 * 1, YRES / 4 * 1, XRES / 4, 100, CogMisc.DegToRad(0), 0)
        Call g_CogCaliperRegion(i, 1).SetCenterLengthsRotationSkew(XRES / 4 * 1, YRES / 4 * 1, XRES / 4, 100, CogMisc.DegToRad(0), 0)
        Call g_CogCaliperRegion(i, 2).SetCenterLengthsRotationSkew(XRES / 4 * 3, YRES / 4 * 1, XRES / 4, 100, CogMisc.DegToRad(180), 0)
        Call g_CogCaliperRegion(i, 3).SetCenterLengthsRotationSkew(XRES / 4 * 3, YRES / 4 * 1, XRES / 4, 100, CogMisc.DegToRad(180), 0)
        Call g_CogCaliperRegion(i, 4).SetCenterLengthsRotationSkew(XRES / 4 * 1, YRES / 4 * 3, XRES / 4, 100, CogMisc.DegToRad(0), 0)
        Call g_CogCaliperRegion(i, 5).SetCenterLengthsRotationSkew(XRES / 4 * 1, YRES / 4 * 3, XRES / 4, 100, CogMisc.DegToRad(0), 0)
        Call g_CogCaliperRegion(i, 6).SetCenterLengthsRotationSkew(XRES / 4 * 3, YRES / 4 * 3, XRES / 4, 100, CogMisc.DegToRad(180), 0)
        Call g_CogCaliperRegion(i, 7).SetCenterLengthsRotationSkew(XRES / 4 * 3, YRES / 4 * 3, XRES / 4, 100, CogMisc.DegToRad(180), 0)
        Call g_CogCaliperRegion(i, 8).SetCenterLengthsRotationSkew(XRES / 4 * 1, YRES / 4 * 1, XRES / 4, 100, CogMisc.DegToRad(90), 0)
        Call g_CogCaliperRegion(i, 9).SetCenterLengthsRotationSkew(XRES / 4 * 1, YRES / 4 * 1, XRES / 4, 100, CogMisc.DegToRad(90), 0)
        Call g_CogCaliperRegion(i, 10).SetCenterLengthsRotationSkew(XRES / 4 * 1, YRES / 4 * 3, XRES / 4, 100, CogMisc.DegToRad(-90), 0)
        Call g_CogCaliperRegion(i, 11).SetCenterLengthsRotationSkew(XRES / 4 * 1, YRES / 4 * 3, XRES / 4, 100, CogMisc.DegToRad(-90), 0)
        Call g_CogCaliperRegion(i, 12).SetCenterLengthsRotationSkew(XRES / 4 * 3, YRES / 4 * 1, XRES / 4, 100, CogMisc.DegToRad(90), 0)
        Call g_CogCaliperRegion(i, 13).SetCenterLengthsRotationSkew(XRES / 4 * 3, YRES / 4 * 1, XRES / 4, 100, CogMisc.DegToRad(90), 0)
        Call g_CogCaliperRegion(i, 14).SetCenterLengthsRotationSkew(XRES / 4 * 3, YRES / 4 * 3, XRES / 4, 100, CogMisc.DegToRad(-90), 0)
        Call g_CogCaliperRegion(i, 15).SetCenterLengthsRotationSkew(XRES / 4 * 3, YRES / 4 * 3, XRES / 4, 100, CogMisc.DegToRad(-90), 0)
    Next i
    
    
End Function

Public Function LoadCogTool(ModelName As String) As Boolean
On Error Resume Next

    Dim FileName As String
    Dim Section As String
    Dim i, j As Integer
    
    FileName = App.Path & "\model\" & ModelName & "\model.ini"
    
    For i = 0 To kMaxCamera - 1
        For j = 0 To kMaxTool - 1
            Dim CaliperTool As CogCaliperTool
            
            Section = "CogCaliperTool" & CStr(i) & "_" & CStr(j)
            Set CaliperTool = g_CogCaliperTool(i, j)
            CaliperTool.RunParams.ContrastThreshold = CDbl(ReadINI(FileName, Section, "ContrastThreshold"))
            
            Dim RectangleAffine As CogRectangleAffine
            
            Section = "CogCaliperRegion" & CStr(i) & "_" & CStr(j)
            Set RectangleAffine = g_CogCaliperRegion(i, j)
            RectangleAffine.CenterX = CDbl(ReadINI(FileName, Section, "CenterX"))
            RectangleAffine.CenterY = CDbl(ReadINI(FileName, Section, "CenterY"))
            RectangleAffine.SideXLength = CDbl(ReadINI(FileName, Section, "SideXLength"))
            RectangleAffine.SideYLength = CDbl(ReadINI(FileName, Section, "SideYLength"))
            RectangleAffine.Rotation = CogMisc.DegToRad(CDbl(ReadINI(FileName, Section, "Rotation")))
            RectangleAffine.Skew = CogMisc.DegToRad(CDbl(ReadINI(FileName, Section, "Skew")))
        Next j
    Next i
    
    
End Function

Public Function SaveCogTool(ModelName As String) As Boolean

    Dim FileName As String
    Dim Section As String
    Dim i, j As Integer
    
    FileName = App.Path & "\model\" & ModelName & "\model.ini"
    
    For i = 0 To kMaxCamera - 1
        For j = 0 To kMaxTool - 1
            Dim CaliperTool As CogCaliperTool
            
            Section = "CogCaliperTool" & CStr(i) & "_" & CStr(j)
            Set CaliperTool = g_CogCaliperTool(i, j)
            Call WriteINI(FileName, Section, "", CaliperTool.RunParams.ContrastThreshold)
            
            Dim RectangleAffine As CogRectangleAffine
            
            Section = "CogCaliperRegion" & CStr(i) & "_" & CStr(j)
            Set RectangleAffine = g_CogCaliperRegion(i, j)
            Call WriteINI(FileName, Section, "CenterX", CStr(RectangleAffine.CenterX))
            Call WriteINI(FileName, Section, "CenterY", CStr(RectangleAffine.CenterY))
            Call WriteINI(FileName, Section, "SideXLength", CStr(RectangleAffine.SideXLength))
            Call WriteINI(FileName, Section, "SideYLength", CStr(RectangleAffine.SideYLength))
            Call WriteINI(FileName, Section, "Rotation", CStr(CogMisc.RadToDeg(RectangleAffine.Rotation)))
            Call WriteINI(FileName, Section, "Skew", CStr(CogMisc.RadToDeg(RectangleAffine.Skew)))
        Next j
    Next i

End Function

Public Sub JR_AutoRun()
On Error GoTo ErrorHandle
Dim bRet As Boolean
Dim tempstr As String
Dim ImageFolderName As String
Dim ImageFolderName2 As String
Dim sMesSendJPGPath As String
    Dim sDate As String
    Dim stime As String
    Dim sMESDate As String
    Dim sMesTime As String
    Dim sTempCode As String
    Dim starttime As Long
    Dim endtime As Long
    Dim sDataTemp As String
    Dim ltemp(0 To 99) As Long
    Dim bResult As Boolean
    Dim i As Integer

    Dim nData() As Long
    
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
    Call Create_DIR("D:\MES\")
    Call Create_DIR("D:\MES\SEND\")

    
    Do While frmMain.ActEasyIF.Open = 0 And bAutoRunOn = True
        Sleep 1000
    Loop
    
    Do While bAutoRunOn = True
    
        DoEvents
        
        Do While bAutoRunOn = True
            Dim strDeviceList As String
            Dim nSize As Long
            Dim nResult As Long
            Dim RCV As Long
                
            DoEvents
            m_Snd_Bit_1(outreadyVision) = 1
            Call Write_Send_Word(addVisionInspect1, Make_Send_Word(addVisionInspect1, True))
            
            strDeviceList = GetAddressString(lMelsecAddrInput, 4)
            nSize = 4
            
            ReDim nData(nSize)
            
            nResult = frmMain.ActEasyIF.ReadDeviceRandom(strDeviceList, nSize, nData(0))
            
            RCV = nData(0)
            For i = 1 To 4
                If nData(i) > 0 Then
                    RCV = RCV + (2 ^ i)
                End If
            Next i

            If nResult = 0 Then
                Call Read_Recieve_Bit(RCV)
            End If
            
            If RCV > 0 Then
                Exit Do
            End If
        Loop

        If bAutoRunOn = False Then
            Exit Do
        End If

        starttime = GetTickCount

        '검사중 신호 전송
        bResult = True
        Call ClearMelsecResult(addVisionInspect1)
        m_Snd_Bit_1(outreadyVision) = 1
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

        ' 조명 켬.
        Call PWM_LightAll(True, 100)

        For i = 0 To kMaxCamera - 1

            If sIDCode(i) = "" Then
                sIDCode(i) = "NOID"
            End If
            
            CogDisplayClear frmMain.CogDisplay(i)

            Set g_CogImage(i) = IDS_AcquireCognex(frmMain.uEyeCam1(i), frmMain.CogDisplay(i))
            If bCamPass = False And m_Rcv_Bit(i) = 1 Then
                Call JR_ManualRun(i, frmMain.CogDisplay(i))
            End If
            
            If m_Rcv_Bit(i) = 0 Then
                CogDisplayLabel frmMain.CogDisplay(i), 200, 200, "Pass", cogColorOrange, "Tahoma", 16
            ElseIf bResultJudge(i) = True Or bCamPass = True Then
                CogDisplayLabel frmMain.CogDisplay(i), 200, 200, "OK", cogColorGreen, "Tahoma", 16
                lOKCount = lOKCount + 1
                lToTalCount = lToTalCount + 1
                frmMain.lblCountOK.Caption = lOKCount
                frmMain.lblCountTotal.Caption = lToTalCount
                m_Snd_Bit_1(outOk1cam + (i * 2)) = 1
'                If bOKimageSave = True Then
'                    Call DJ_ImageSave(i, ImageFolderName, "OK", iImageFileMode)
'                End If
            Else
                CogDisplayLabel frmMain.CogDisplay(i), 200, 200, "NG", cogColorGreen, "Tahoma", 16
                lNGCount = lNGCount + 1
                lToTalCount = lToTalCount + 1
                frmMain.lblCountNG.Caption = lNGCount
                frmMain.lblCountTotal.Caption = lToTalCount
                m_Snd_Bit_1(outNG1cam + (i * 2)) = 1
                bResult = False
'                If bNGimageSave = True Then
'                    Call DJ_ImageSave(i, ImageFolderName, "NG", iImageFileMode)
'                End If
                Call MES_DATASEND_FUNC("NG_PRODUCT_EVENT", sIDCode(i), "")
            End If
            
            '카운트 저장
            Call SaveCount

            Call JR_WriteDataToGrid(i)
        Next i

        ' 조명 끔.
        Call PWM_LightAll(False)
        
        '검사종료 및 OK NG 신호 전송
        m_Snd_Bit_1(outBusyVision) = 0
        m_Snd_Bit_1(outEndVision) = 1
        Call Write_Send_Word(addVisionInspect1, Make_Send_Word(addVisionInspect1, True))
        
        Call MES_DATASEND_FUNC("QMS_EVENT", "", "")
        
        If bResult = True Then
            'OK 신호
            frmMain.lblResults.Caption = "O.K"
            frmMain.ShpResult.BackColor = &H8000&
        Else
            'NG 전송
            frmMain.lblResults.Caption = "N.G"
            frmMain.ShpResult.BackColor = vbRed
        End If

        Call Dlay_T(0.1)
        sMesSendJPGPath = "D:\MES\SEND\" & sIDCode(0) & "^" & sIDCode(1) & "^" & sIDCode(2) & "^" & sIDCode(3) & "_" & sMESEquipCode & "_" & 1 & "_" & sDateTimeCheck & "_001" & ".JPG"
        Call SH_ScreenSave(sMesSendJPGPath)
        Call SH_ScreenSave(ImageFolderName & sIDCode(0) & "^" & sIDCode(1) & "^" & sIDCode(2) & "^" & sIDCode(3) & "_" & sMESEquipCode & "_" & 1 & "_" & sDateTimeCheck & "_001" & ".JPG")
        Call Dlay_T(0.01)

        For i = 0 To kMaxCamera - 1
            sDataTemp = DJ_DataFileADD(i)
            Call DataFileSave(i, sDataTemp, "d:\MES\SEND\" & sIDCode(i) & "_" & sMESEquipCode & "_" & i + 1 & "_" & sDateTimeCheck & ".QCP")     'MES 에 전송할 데이터 생성
            Call DataFileSave(i, sDataTemp, ImageFolderName & sIDCode(i) & "_" & sMESEquipCode & "_" & i + 1 & "_" & sDateTimeCheck & ".QCP")     '저장되는 데이터 생성
        Next i

        Call MES_NetDriveConnect

        endtime = GetTickCount
        frmMain.lblInspecTime.Caption = CStr(endtime - starttime)

    Loop
    
    Call ClearMelsecResult(addVisionInspect1)
    m_Snd_Bit_1(outreadyVision) = 0
    Call Write_Send_Word(addVisionInspect1, Make_Send_Word(addVisionInspect1, True))
    
    Exit Sub
ErrorHandle:
    
End Sub
Public Function JR_ManualRun(CamIndex As Integer, Optional ByRef Display As CogDisplay = Nothing) As Boolean
On Error GoTo ErrorHandle
    
    Dim Distance(0 To 3) As Double
    
    Dim ToolIdx As Integer
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To kMaxTool - 1 Step 4
        For j = 0 To 3
            Set g_CogCaliperTool(CamIndex, i + j).InputImage = g_CogImage(CamIndex)
        Next j
        Distance(i / 4) = CogFindDistance(g_CogCaliperTool(CamIndex, i + 0), g_CogCaliperTool(CamIndex, i + 1), g_CogCaliperTool(CamIndex, i + 2), g_CogCaliperTool(CamIndex, i + 3), Display, dCaliPX(CamIndex), dSpecOffset(CamIndex * 10 + (i / 4)))
        g_Distance(CamIndex, i / 4) = Distance(i / 4)
    Next i
    
    For i = 0 To 3
        frmMain.lblResultData(CamIndex * 10 + i).Caption = Format(Distance(i), "#0.00")
        If Distance(i) < dSpecOriMin(i) Or Distance(i) > dSpecOriMax(i) Then
            bResultJudge(CamIndex) = False
        Else
            bResultJudge(CamIndex) = True
        End If
    Next i

    JR_ManualRun = True
  
    Exit Function
ErrorHandle:
    JR_ManualRun = False
    
End Function

Public Sub JR_InitGrid()
    Dim i As Integer
Dim j As Integer
Dim stemp As String
Dim str_BlobName As String
    frmMain.MSFlexGrid1.WordWrap = True                    '한 Cell 에 두줄 쓸수 있게 됨
    frmMain.MSFlexGrid1.AllowUserResizing = flexResizeBoth 'Cell Size 를 마우스로 조절 할수 있음
    frmMain.MSFlexGrid1.SelectionMode = flexSelectionByRow
    
    For i = 1 To 4
        'stemp = stemp & "^" & sSpecName(i - 1) & vbCrLf & dSpecOriMin(i - 1) & "~" & dSpecOriMax(i - 1) & "         |"
        stemp = stemp & "^" & sSpecName(i - 1) & Chr(13) & dSpecOriMin(i - 1) & "~" & dSpecOriMax(i - 1) & "         |"
    Next i
    
    '검사결과 Display
    frmMain.MSFlexGrid1.Rows = 1
    frmMain.MSFlexGrid1.Cols = 4 + 1
    frmMain.MSFlexGrid1.FormatString = "^Number    |" & "^검사시간                |" & "^ID_Code                    |" & "^판정  |" & stemp
    
    For j = 0 To 0
        frmMain.MSFlexGrid1.ColWidth(j) = 1500
    Next j
    
    frmMain.MSFlexGrid1.RowHeight(0) = 600
    frmMain.MSFlexGrid1.ColWidth(0) = 800
    frmMain.MSFlexGrid1.ColWidth(1) = 1400
    frmMain.MSFlexGrid1.ColWidth(2) = 4000
    frmMain.MSFlexGrid1.ColWidth(3) = 400
        
End Sub

Public Sub JR_WriteDataToGrid(Index As Integer)
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

    Dim tmpColN As Integer
    Dim tmpColN2 As Integer

        frmMain.MSFlexGrid1.TextMatrix(Rownum - 1, 0) = lInspectionNum
        frmMain.MSFlexGrid1.TextMatrix(Rownum - 1, 1) = Time
        frmMain.MSFlexGrid1.TextMatrix(Rownum - 1, 2) = sIDCode(Index)
        frmMain.MSFlexGrid1.TextMatrix(Rownum - 1, 3) = tempstr
        tmpColN = 3
                    
        frmMain.MSFlexGrid1.TextMatrix(Rownum - 1, 4) = Format(g_Distance(Index, 0), "#0.00")
        frmMain.MSFlexGrid1.TextMatrix(Rownum - 1, 5) = Format(g_Distance(Index, 1), "#0.00")
        frmMain.MSFlexGrid1.TextMatrix(Rownum - 1, 6) = Format(g_Distance(Index, 2), "#0.00")
        frmMain.MSFlexGrid1.TextMatrix(Rownum - 1, 7) = Format(g_Distance(Index, 3), "#0.00")
        frmMain.MSFlexGrid1.Col = 0
        frmMain.MSFlexGrid1.Row = Rownum - 1
        frmMain.MSFlexGrid1.CellForeColor = vbWhite
        lInspectionNum = lInspectionNum + 1

    frmMain.MSFlexGrid1.Col = 2
    frmMain.MSFlexGrid1.Row = Rownum - 1
    If sIDCode(Index) = "NOID" Then
        frmMain.MSFlexGrid1.CellForeColor = vbRed
    End If
    frmMain.MSFlexGrid1.Col = 3
    frmMain.MSFlexGrid1.Row = Rownum - 1
    
    If tempstr = "OK" Then
        frmMain.MSFlexGrid1.CellForeColor = vbBlue
    Else
        frmMain.MSFlexGrid1.CellForeColor = vbRed
    End If
    
    For i = 0 To 3
        If frmMain.lblResultData(Index * 10 + i).BackColor = vbRed Then
            frmMain.MSFlexGrid1.Col = 4 + i
            frmMain.MSFlexGrid1.Row = Rownum - 1
            frmMain.MSFlexGrid1.CellForeColor = vbRed
        End If
    Next i
    
    frmMain.MSFlexGrid1.Row = 1
    frmMain.MSFlexGrid1.Col = 0
    frmMain.MSFlexGrid1.Sort = 4

End Sub


Public Sub JR_ModelLoad(ModelName As String)
On Error Resume Next

    'Call ModelData_FileLoad(mdlname)
    Call Calibration_Load(ModelName)
    'Call Calibration_Loady(mdlname)
    'Call FixPoint_Load(mdlname)
    Call SpecName_Load(ModelName)
    Call SpecAllValue_Load(ModelName)
    Call FunctionValue_Load(ModelName)
    Call LoadCogTool(ModelName)
    
End Sub

Public Sub JR_ModelSave(ModelName As String)
On Error Resume Next
    
    Call Calibration_Save(ModelName)
    Call SpecName_Save(ModelName)
    Call SpecAllValue_Save(ModelName)
    Call FunctionValue_Save(ModelName)
    Call SaveCogTool(ModelName)
End Sub

