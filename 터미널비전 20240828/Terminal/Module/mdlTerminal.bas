Attribute VB_Name = "mdlTerminal"
Public Const kMaxCamera As Integer = 4
Public Const kMaxLight As Integer = 3
Public Const kMaxTool As Integer = 8

Public g_CogImage(0 To kMaxCamera - 1) As New CogImage8Grey

Public g_CogCalibrationTool(0 To 3) As CogCaliperTool
Public g_CogCalibrationRegion(0 To 3) As CogRectangleAffine

Public g_CogCaliperTool(0 To kMaxTool - 1) As CogCaliperTool
Public g_CogCaliperRegion(0 To kMaxTool - 1) As CogRectangleAffine
Public g_CogCaliperScorer As CogCaliperScorerPositionNeg
Public g_CogCaliperScorerPosition As CogCaliperScorerPosition

Public g_CogNsdTool(0 To 11) As CogCaliperTool
Public g_CogNsdRegion(0 To 11) As CogRectangleAffine

Public g_CogBlobTool As CogBlobTool
Public g_CogBlobRegion As CogRectangle
Public g_CogBlobIndex As Long

Public g_CogFindLineTool(0 To kMaxCamera - 1, 0 To kMaxTool - 1) As New CogFindLineTool
Public g_CogFindLineSegment(0 To kMaxCamera - 1, 0 To kMaxTool - 1) As New CogLineSegment

Public g_CogGapRegion(0 To 7) As New CogRectangleAffine

Public g_Distance(0 To 3) As Double
Public g_NsdDistance(0 To 5) As Double

Public g_ProductPt(0 To 3) As Double
Public g_CameraGrap(0 To 3) As Double

Public g_NGCount As Long
Public g_Judge(0 To 15) As Boolean

'NSD영역 좌우측 선택
Public g_NsdRegionSelection As Integer

Public g_Temp(0 To 9) As Double

Public Function InitCogTool() As Boolean
On Error Resume Next

    Dim CaliperTool As CogCaliperTool
    Dim CaliperRegion As CogRectangleAffine
    Dim CaliperScorer As CogCaliperScorerPositionNeg
    
    Dim i, j As Integer
    
    Set g_CogCaliperScorer = New CogCaliperScorerPositionNeg
    g_CogCaliperScorer.Enabled = True
    g_CogCaliperScorer.SetXYParameters -100, 100, 10000, 1, 0.1
    
    Set g_CogCaliperScorerPosition = New CogCaliperScorerPosition
    g_CogCaliperScorerPosition.Enabled = True
    g_CogCaliperScorerPosition.SetXYParameters 0, 100, 10000, 1, 0.1
    
    For i = 0 To 3
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
            .SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, 50, 0, 0
            .GraphicDOFEnable = cogRectangleAffineDOFAll
            .XDirectionAdornment = cogRectangleAffineDirectionAdornmentSolidArrow
            .YDirectionAdornment = cogRectangleAffineDirectionAdornmentArrow
            .Color = cogColorGreen
            .Interactive = True
        End With
    Next i
    
    g_CogCalibrationRegion(1).Rotation = CogMisc.DegToRad(180)
    g_CogCalibrationRegion(2).Rotation = CogMisc.DegToRad(90)
    g_CogCalibrationRegion(3).Rotation = CogMisc.DegToRad(-90)
    
    For j = 0 To kMaxTool - 1
        Set g_CogCaliperTool(j) = New CogCaliperTool
        Set g_CogCaliperRegion(j) = New CogRectangleAffine
        Set g_CogCaliperTool(j).Region = g_CogCaliperRegion(j)
        
        With g_CogCaliperTool(j).RunParams
            .ContrastThreshold = 10
            .Edge0Polarity = cogCaliperPolarityLightToDark
            .EdgeMode = cogCaliperEdgeModeSingle
            .FilterHalfSizeInPixels = 3
            .MaxResults = 1
            .SingleEdgeScorers.Clear
            .SingleEdgeScorers.Add g_CogCaliperScorer
        End With
        
        With g_CogCaliperRegion(j)
            .GraphicDOFEnable = cogRectangleAffineDOFAll
            .XDirectionAdornment = cogRectangleAffineDirectionAdornmentSolidArrow
            .YDirectionAdornment = cogRectangleAffineDirectionAdornmentArrow
            .Color = cogColorGreen
            .Interactive = True
        End With
        
    Next j
        
    Call g_CogCaliperRegion(0).SetCenterLengthsRotationSkew(XRES / 2, YRES / 2, XRES / 4, 100, CogMisc.DegToRad(0), 0)
    Call g_CogCaliperRegion(1).SetCenterLengthsRotationSkew(XRES / 2, YRES / 2, XRES / 4, 100, CogMisc.DegToRad(180), 0)
    Call g_CogCaliperRegion(2).SetCenterLengthsRotationSkew(XRES / 2, YRES / 2, XRES / 4, 100, CogMisc.DegToRad(0), 0)
    Call g_CogCaliperRegion(3).SetCenterLengthsRotationSkew(XRES / 2, YRES / 2, XRES / 4, 100, CogMisc.DegToRad(180), 0)
    Call g_CogCaliperRegion(4).SetCenterLengthsRotationSkew(XRES / 2, YRES / 2, XRES / 4, 100, CogMisc.DegToRad(90), 0)
    Call g_CogCaliperRegion(5).SetCenterLengthsRotationSkew(XRES / 2, YRES / 2, XRES / 4, 100, CogMisc.DegToRad(-90), 0)
    Call g_CogCaliperRegion(6).SetCenterLengthsRotationSkew(XRES / 2, YRES / 2, XRES / 4, 100, CogMisc.DegToRad(90), 0)
    Call g_CogCaliperRegion(7).SetCenterLengthsRotationSkew(XRES / 2, YRES / 2, XRES / 4, 100, CogMisc.DegToRad(-90), 0)
    
    For j = 0 To 11
        Set g_CogNsdTool(j) = New CogCaliperTool
        Set g_CogNsdRegion(j) = New CogRectangleAffine
        Set g_CogNsdTool(j).Region = g_CogNsdRegion(j)
        
        With g_CogNsdTool(j).RunParams
            .ContrastThreshold = 10
            .Edge0Polarity = cogCaliperPolarityLightToDark
            .EdgeMode = cogCaliperEdgeModeSingle
            Select Case j
            Case 0, 1, 4, 5 '세로 측정 NSD (1캠,3캠)
                .FilterHalfSizeInPixels = 10
            Case Else
                .FilterHalfSizeInPixels = 5
                .SingleEdgeScorers.Clear
                .SingleEdgeScorers.Add g_CogCaliperScorer
            End Select
            .MaxResults = 1
        End With
        
        With g_CogNsdRegion(j)
            .GraphicDOFEnable = cogRectangleAffineDOFAll
            .XDirectionAdornment = cogRectangleAffineDirectionAdornmentSolidArrow
            .YDirectionAdornment = cogRectangleAffineDirectionAdornmentArrow
            .Color = cogColorGreen
            .Interactive = True
        End With
        
    Next j
    
    g_CogNsdRegion(0).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, CogMisc.DegToRad(90), 0
    g_CogNsdRegion(1).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, CogMisc.DegToRad(-90), 0
    g_CogNsdRegion(2).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, CogMisc.DegToRad(90), 0
    g_CogNsdRegion(3).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, CogMisc.DegToRad(-90), 0
    g_CogNsdRegion(4).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, CogMisc.DegToRad(0), 0
    g_CogNsdRegion(5).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, CogMisc.DegToRad(180), 0
    g_CogNsdRegion(6).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, CogMisc.DegToRad(0), 0
    g_CogNsdRegion(7).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, CogMisc.DegToRad(180), 0
    g_CogNsdRegion(8).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, CogMisc.DegToRad(90), 0
    g_CogNsdRegion(9).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, CogMisc.DegToRad(-90), 0
    g_CogNsdRegion(10).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, CogMisc.DegToRad(90), 0
    g_CogNsdRegion(11).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, CogMisc.DegToRad(-90), 0
    
    For i = 0 To 7
        With g_CogGapRegion(i)
            .Visible = True
            .Interactive = True
            .LineStyle = cogGraphicLineStyleSolid
            .XDirectionAdornment = cogRectangleAffineDirectionAdornmentSolidArrow
            .YDirectionAdornment = cogRectangleAffineDirectionAdornmentNone
            .GraphicDOFEnable = cogRectangleAffineDOFPosition + cogRectangleAffineDOFSize + cogRectangleAffineDOFScale
            .Color = cogColorGreen
        End With
    Next i
    
    g_CogGapRegion(0).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, 0, 0
    g_CogGapRegion(1).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, CogMisc.DegToRad(180), 0
    g_CogGapRegion(2).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, 0, 0
    g_CogGapRegion(3).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, CogMisc.DegToRad(180), 0
    g_CogGapRegion(4).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, CogMisc.DegToRad(90), 0
    g_CogGapRegion(5).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, CogMisc.DegToRad(-90), 0
    g_CogGapRegion(6).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, CogMisc.DegToRad(90), 0
    g_CogGapRegion(7).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, CogMisc.DegToRad(-90), 0
    
    Set g_CogBlobTool = New CogBlobTool
    Set g_CogBlobRegion = New CogRectangle
    Set g_CogBlobTool.Region = g_CogBlobTool
    
    With g_CogBlobTool.RunParams
        With .SegmentationParams
            .Mode = cogBlobSegmentationModeHardFixedThreshold
            .HardFixedThreshold = 128
            .Polarity = cogBlobSegmentationPolarityLightBlobs
        End With
        .ConnectivityMinPixels = 500
    End With
    
    With g_CogBlobRegion
        .Visible = True
        .Interactive = True
        .Color = cogColorCyan
        .GraphicDOFEnable = cogRectangleDOFAll
        .SetXYWidthHeight XRES / 2, YRES / 2, XRES / 8, YRES / 8
    End With
    
    g_CogBlobIndex = 0
    
End Function



Public Function LoadMultiROI(ByVal ModelName As String, ByVal ROINo As Integer) As Boolean
On Error Resume Next

Dim FileName As String
Dim Section As String
Dim i, j As Integer
    
    If ROINo = 0 Then
        FileName = App.Path & "\model\" & ModelName & "\model.ini"
    Else
        FileName = App.Path & "\model\" & ModelName & "\model" & Format(ROINo, "00") & ".ini"
    End If
    
    
    For j = 0 To kMaxTool - 1
        Dim RectangleAffine As CogRectangleAffine
        
        Section = "CogCaliperRegion" & CStr(j)
        Set RectangleAffine = g_CogCaliperRegion(j)
        RectangleAffine.CenterX = CDbl(ReadINI(FileName, Section, "CenterX"))
        RectangleAffine.CenterY = CDbl(ReadINI(FileName, Section, "CenterY"))
        RectangleAffine.SideXLength = CDbl(ReadINI(FileName, Section, "SideXLength"))
        RectangleAffine.SideYLength = CDbl(ReadINI(FileName, Section, "SideYLength"))
        RectangleAffine.Rotation = CogMisc.DegToRad(CDbl(ReadINI(FileName, Section, "Rotation")))
        RectangleAffine.Skew = CogMisc.DegToRad(CDbl(ReadINI(FileName, Section, "Skew")))
    Next j
    
    For j = 0 To 11
        Section = "CogNsdRegion" & CStr(j)
        Set RectangleAffine = g_CogNsdRegion(j)
        RectangleAffine.CenterX = CDbl(ReadINI(FileName, Section, "CenterX"))
        RectangleAffine.CenterY = CDbl(ReadINI(FileName, Section, "CenterY"))
        RectangleAffine.SideXLength = CDbl(ReadINI(FileName, Section, "SideXLength"))
        RectangleAffine.SideYLength = CDbl(ReadINI(FileName, Section, "SideYLength"))
        RectangleAffine.Rotation = CogMisc.DegToRad(CDbl(ReadINI(FileName, Section, "Rotation")))
        RectangleAffine.Skew = CogMisc.DegToRad(CDbl(ReadINI(FileName, Section, "Skew")))
    Next j

End Function


Public Function LoadCogTool(ModelName As String) As Boolean
On Error Resume Next

    Dim FileName As String
    Dim Section As String
    Dim i, j As Integer
    
    FileName = App.Path & "\model\" & ModelName & "\model.ini"
    
    
    For j = 0 To kMaxTool - 1
        Dim CaliperTool As CogCaliperTool
        
        Section = "CogCaliperTool" & "_" & CStr(j)
        Set CaliperTool = g_CogCaliperTool(j)
        CaliperTool.RunParams.ContrastThreshold = CDbl(ReadINI(FileName, Section, "ContrastThreshold"))
        CaliperTool.RunParams.FilterHalfSizeInPixels = CLng(ReadINI(FileName, Section, "FilterHalfSizeInPixels"))
        CaliperTool.RunParams.Edge0Polarity = CInt(ReadINI(FileName, Section, "Edge0Polarity"))
        
        Dim RectangleAffine As CogRectangleAffine
        
        Section = "CogCaliperRegion" & CStr(j)
        Set RectangleAffine = g_CogCaliperRegion(j)
        RectangleAffine.CenterX = CDbl(ReadINI(FileName, Section, "CenterX"))
        RectangleAffine.CenterY = CDbl(ReadINI(FileName, Section, "CenterY"))
        RectangleAffine.SideXLength = CDbl(ReadINI(FileName, Section, "SideXLength"))
        RectangleAffine.SideYLength = CDbl(ReadINI(FileName, Section, "SideYLength"))
        RectangleAffine.Rotation = CogMisc.DegToRad(CDbl(ReadINI(FileName, Section, "Rotation")))
        RectangleAffine.Skew = CogMisc.DegToRad(CDbl(ReadINI(FileName, Section, "Skew")))
    Next j
    
    For j = 0 To 7
        'Dim RectangleAffine As CogRectangleAffine
        
        Section = "CogGapRegion" & CStr(j)
        Set RectangleAffine = g_CogGapRegion(j)
        RectangleAffine.CenterX = CDbl(ReadINI(FileName, Section, "CenterX"))
        RectangleAffine.CenterY = CDbl(ReadINI(FileName, Section, "CenterY"))
        RectangleAffine.SideXLength = CDbl(ReadINI(FileName, Section, "SideXLength"))
        RectangleAffine.SideYLength = CDbl(ReadINI(FileName, Section, "SideYLength"))
        RectangleAffine.Rotation = CogMisc.DegToRad(CDbl(ReadINI(FileName, Section, "Rotation")))
        RectangleAffine.Skew = CogMisc.DegToRad(CDbl(ReadINI(FileName, Section, "Skew")))
'
'        If RectangleAffine.SideXLength = 0 Then
'            g_CogGapRegion(0).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, 0, 0
'            g_CogGapRegion(1).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, CogMisc.DegToRad(180), 0
'            g_CogGapRegion(2).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, 0, 0
'            g_CogGapRegion(3).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, CogMisc.DegToRad(180), 0
'            g_CogGapRegion(4).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, CogMisc.DegToRad(90), 0
'            g_CogGapRegion(5).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, CogMisc.DegToRad(-90), 0
'            g_CogGapRegion(6).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, CogMisc.DegToRad(90), 0
'            g_CogGapRegion(7).SetCenterLengthsRotationSkew XRES / 2, YRES / 2, XRES / 8, YRES / 16, CogMisc.DegToRad(-90), 0
'            Exit For
'        End If
    Next j
    
    For j = 0 To 11
        Section = "CogNsdTool" & "_" & CStr(j)
        Set CaliperTool = g_CogNsdTool(j)
        CaliperTool.RunParams.ContrastThreshold = CDbl(ReadINI(FileName, Section, "ContrastThreshold"))
        CaliperTool.RunParams.FilterHalfSizeInPixels = CLng(ReadINI(FileName, Section, "FilterHalfSizeInPixels"))
        CaliperTool.RunParams.Edge0Polarity = CInt(ReadINI(FileName, Section, "Edge0Polarity"))
        
        Section = "CogNsdRegion" & CStr(j)
        Set RectangleAffine = g_CogNsdRegion(j)
        RectangleAffine.CenterX = CDbl(ReadINI(FileName, Section, "CenterX"))
        RectangleAffine.CenterY = CDbl(ReadINI(FileName, Section, "CenterY"))
        RectangleAffine.SideXLength = CDbl(ReadINI(FileName, Section, "SideXLength"))
        RectangleAffine.SideYLength = CDbl(ReadINI(FileName, Section, "SideYLength"))
        RectangleAffine.Rotation = CogMisc.DegToRad(CDbl(ReadINI(FileName, Section, "Rotation")))
        RectangleAffine.Skew = CogMisc.DegToRad(CDbl(ReadINI(FileName, Section, "Skew")))
    Next j
    
    Section = "BlobParam"
    g_CogBlobIndex = CInt(ReadINI(FileName, Section, "Index"))
    
    Call CogMisc.LoadObjectFromFile(App.Path & "\model\" & ModelName & "\BlobTool.vpp", g_CogBlobTool, cogPersistOptionMinimum)
    Call CogMisc.LoadObjectFromFile(App.Path & "\model\" & ModelName & "\BlobRegion.vpp", g_CogBlobRegion, cogPersistOptionMinimum)
    
    
End Function

Public Function SaveMultiROI(ByVal ModelName As String, ByVal ROINo As Integer) As Boolean

Dim FileName As String
Dim Section As String
Dim i, j As Integer
    
    If ROINo = 0 Then
        FileName = App.Path & "\model\" & ModelName & "\model.ini"
    Else
        FileName = App.Path & "\model\" & ModelName & "\model" & Format(ROINo, "00") & ".ini"
    End If
    

    For j = 0 To kMaxTool - 1
        Dim RectangleAffine As CogRectangleAffine
        
        Section = "CogCaliperRegion" & CStr(j)
        Set RectangleAffine = g_CogCaliperRegion(j)
        Call WriteINI(FileName, Section, "CenterX", CStr(RectangleAffine.CenterX))
        Call WriteINI(FileName, Section, "CenterY", CStr(RectangleAffine.CenterY))
        Call WriteINI(FileName, Section, "SideXLength", CStr(RectangleAffine.SideXLength))
        Call WriteINI(FileName, Section, "SideYLength", CStr(RectangleAffine.SideYLength))
        Call WriteINI(FileName, Section, "Rotation", CStr(CogMisc.RadToDeg(RectangleAffine.Rotation)))
        Call WriteINI(FileName, Section, "Skew", CStr(CogMisc.RadToDeg(RectangleAffine.Skew)))
    Next j

    For j = 0 To 11
        Section = "CogNsdRegion" & CStr(j)
        Set RectangleAffine = g_CogNsdRegion(j)
        Call WriteINI(FileName, Section, "CenterX", CStr(RectangleAffine.CenterX))
        Call WriteINI(FileName, Section, "CenterY", CStr(RectangleAffine.CenterY))
        Call WriteINI(FileName, Section, "SideXLength", CStr(RectangleAffine.SideXLength))
        Call WriteINI(FileName, Section, "SideYLength", CStr(RectangleAffine.SideYLength))
        Call WriteINI(FileName, Section, "Rotation", CStr(CogMisc.RadToDeg(RectangleAffine.Rotation)))
        Call WriteINI(FileName, Section, "Skew", CStr(CogMisc.RadToDeg(RectangleAffine.Skew)))
    Next j
    
End Function

Public Function SaveCogTool(ModelName As String) As Boolean

    Dim FileName As String
    Dim Section As String
    Dim i, j As Integer
    
    FileName = App.Path & "\model\" & ModelName & "\model.ini"
    

    For j = 0 To kMaxTool - 1
        Dim CaliperTool As CogCaliperTool
        
        Section = "CogCaliperTool" & "_" & CStr(j)
        Set CaliperTool = g_CogCaliperTool(j)
        Call WriteINI(FileName, Section, "ContrastThreshold", CaliperTool.RunParams.ContrastThreshold)
        Call WriteINI(FileName, Section, "FilterHalfSizeInPixels", CStr(CaliperTool.RunParams.FilterHalfSizeInPixels))
        Call WriteINI(FileName, Section, "Edge0Polarity", CStr(CaliperTool.RunParams.Edge0Polarity))
        
        Dim RectangleAffine As CogRectangleAffine
        
        Section = "CogCaliperRegion" & CStr(j)
        Set RectangleAffine = g_CogCaliperRegion(j)
        Call WriteINI(FileName, Section, "CenterX", CStr(RectangleAffine.CenterX))
        Call WriteINI(FileName, Section, "CenterY", CStr(RectangleAffine.CenterY))
        Call WriteINI(FileName, Section, "SideXLength", CStr(RectangleAffine.SideXLength))
        Call WriteINI(FileName, Section, "SideYLength", CStr(RectangleAffine.SideYLength))
        Call WriteINI(FileName, Section, "Rotation", CStr(CogMisc.RadToDeg(RectangleAffine.Rotation)))
        Call WriteINI(FileName, Section, "Skew", CStr(CogMisc.RadToDeg(RectangleAffine.Skew)))
    Next j
    
    For j = 0 To 7
        'Dim RectangleAffine As CogRectangleAffine
        
        Section = "CogGapRegion" & CStr(j)
        Set RectangleAffine = g_CogGapRegion(j)
        Call WriteINI(FileName, Section, "CenterX", CStr(RectangleAffine.CenterX))
        Call WriteINI(FileName, Section, "CenterY", CStr(RectangleAffine.CenterY))
        Call WriteINI(FileName, Section, "SideXLength", CStr(RectangleAffine.SideXLength))
        Call WriteINI(FileName, Section, "SideYLength", CStr(RectangleAffine.SideYLength))
        Call WriteINI(FileName, Section, "Rotation", CStr(CogMisc.RadToDeg(RectangleAffine.Rotation)))
        Call WriteINI(FileName, Section, "Skew", CStr(CogMisc.RadToDeg(RectangleAffine.Skew)))
    Next j
    
    For j = 0 To 11
        Section = "CogNsdTool" & "_" & CStr(j)
        Set CaliperTool = g_CogNsdTool(j)
        Call WriteINI(FileName, Section, "ContrastThreshold", CaliperTool.RunParams.ContrastThreshold)
        Call WriteINI(FileName, Section, "FilterHalfSizeInPixels", CStr(CaliperTool.RunParams.FilterHalfSizeInPixels))
        Call WriteINI(FileName, Section, "Edge0Polarity", CStr(CaliperTool.RunParams.Edge0Polarity))
        
        Section = "CogNsdRegion" & CStr(j)
        Set RectangleAffine = g_CogNsdRegion(j)
        Call WriteINI(FileName, Section, "CenterX", CStr(RectangleAffine.CenterX))
        Call WriteINI(FileName, Section, "CenterY", CStr(RectangleAffine.CenterY))
        Call WriteINI(FileName, Section, "SideXLength", CStr(RectangleAffine.SideXLength))
        Call WriteINI(FileName, Section, "SideYLength", CStr(RectangleAffine.SideYLength))
        Call WriteINI(FileName, Section, "Rotation", CStr(CogMisc.RadToDeg(RectangleAffine.Rotation)))
        Call WriteINI(FileName, Section, "Skew", CStr(CogMisc.RadToDeg(RectangleAffine.Skew)))
    Next j
    
    Call WriteINI(FileName, "BlobParam", "Index", CStr(g_CogBlobIndex))
    
    Call CogMisc.SaveObjectToFile(App.Path & "\model\" & ModelName & "\BlobTool.vpp", g_CogBlobTool, cogPersistOptionMinimum)
    Call CogMisc.SaveObjectToFile(App.Path & "\model\" & ModelName & "\BlobRegion.vpp", g_CogBlobRegion, cogPersistOptionMinimum)
    
End Function

Public Sub Terminal_AutoRun()
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
    Dim bRetry As Boolean
    Dim i As Integer
    
    '검사시간
    Dim tackPLCSpec As Long
    Dim tackGrab As Long
    Dim tackinspection As Long
    Dim tackJudgement As Long
    Dim tackScreenshot As Long
    Dim tackMES As Long
    Dim tackEnd As Long
    
    Dim InputData As Long
    
    Dim tempDistance(0 To 9) As String
    
    
    Call WriteMelsec(frmMain.ActEasyIF, sMelsecAddrAcqDone, 0)
    
    Do While bAutoRunOn = True
    
        DoEvents
        
        '트리거 확인
        Do While bAutoRunOn = True
            DoEvents
            
            '모델 변경 확인
            Call ModelChange
            
            'Ready 신호 전송
            Call SendSignalToMelsec(outReadyVision, 1)
            
            '트리거 신호 읽어오기
            InputData = ReadMelsec(frmMain.ActEasyIF, lMelsecAddrInput, True)
            
            If InputData > 0 Then
                Exit Do
            End If
        Loop

        If bAutoRunOn = False Then
            Exit Do
        End If

        starttime = GetTickCount
        
        '검사중 신호 전송
        bResult = True
        Call SendSignalToMelsec(outBusyVision, 1)
        
        'PROCESS 상태
        Call MES_DATASEND_FUNC("EQ_STATE_EVENT", "PROCESS", "")
        
        'PLC 정보 가져오기(ID & Spec)
        Call ReadDataFromPLC
        '스펙 표시
        Call Terminal_SpecPrint
    
        tackPLCSpec = GetTickCount() - starttime
    
        
        '폴더 생성
        sDate = Format(Date, "yy-mm-dd")
        stime = Format(Time, "hh-mm-ss")
        sMESDate = Format(Date, "YYYYMMDD")
        sMesTime = Format(Time, "HHMMSS")
        sDateTimeCheck = sMESDate & sMesTime
        ImageFolderName = "D:\Imagesave\" & sDate & "\" & sModelName & "\"
        Call Create_DIR(ImageFolderName)

        ' 조명 켬.
        If g_UseLightTimer = 0 Then
            Call PWM_LightAll(True, 100)
        Else
            If g_LightTimerCount <= 0 Then
                Call PWM_LightAll(True)
                Debug.Print "[자동조명] 켜기"
            End If
            g_LightTimerCount = g_LightTimerInterval
        End If
        

        '영상 획득
        For i = 0 To kMaxCamera - 1
            If sIDCode(i) = "" Then
                sIDCode(i) = "NOID"
            End If
                        
            CogDisplayClear frmMain.CogDisplay(i)
            Set g_CogImage(i) = IDS_AcquireCognex(frmMain.uEyeCam1(i), frmMain.CogDisplay(i))
        Next i
        
        ' 조명 끔.
        If g_UseLightTimer = 0 Then
            Call PWM_LightAll(False)
        End If
        
        tackGrab = GetTickCount() - starttime
        
        Call SendSignalToMelsec(outFinishGrab, 1)
        Call WriteMelsec(frmMain.ActEasyIF, sMelsecAddrAcqDone, 1)
        
        bRetry = False
        
AUTORUN_RETRY:


        If bRetry = False And g_UseRetry = 1 Then
            Call LoadMultiROI(sModelName, g_RetryBase)
        End If
        
        '검사
        Call PreWelding_RunWidthHeight(frmMain.CogDisplay(0), frmMain.CogDisplay(1), frmMain.CogDisplay(2), frmMain.CogDisplay(3))
        Call PreWelding_RunNSD(frmMain.CogDisplay(0), frmMain.CogDisplay(1), frmMain.CogDisplay(2), frmMain.CogDisplay(3), frmMain.CogDisplay(g_CogBlobIndex))
        
        tackinspection = GetTickCount() - starttime
        
        '판정
        Call PreWelding_Judgement
        
        '검사종료 및 OK NG 신호 전송
        m_Snd_Bit_1(outBusyVision) = 0
        If g_NGCount > 0 And bCamPass = False Then
            If g_UseRetry = 1 And bRetry = False Then
                bRetry = True
                Call LoadMultiROI(sModelName, g_RetryROI)
                GoTo AUTORUN_RETRY
            End If
            m_Snd_Bit_1(outOk1cam) = 0
            m_Snd_Bit_1(outNG1cam) = 1
        Else
            m_Snd_Bit_1(outOk1cam) = 1
            m_Snd_Bit_1(outNG1cam) = 0
        End If
        Call SendSignalToMelsec(outEndVision, 1)
        
        Dim ResultString As String
        
        '검사 결과 출력
        If g_NGCount > 0 Then
            'NG_PRODUCT_EVENT 메세지 전송
            Call MES_DATASEND_FUNC("NG_PRODUCT_EVENT", "", "")
            ResultString = "NG"
            If bNGimageSave = True Then
                Call Create_DIR(ImageFolderName & "NG")
                For i = 0 To 3
                    Call SaveCogImage(ImageFolderName & "NG" & "\" & sMESDate & "_" & sMesTime & "_" & sIDCode(0) & "_" & "CAM" & CStr(i + 1) & IIf(iImageFileMode = 1, ".bmp", ".jpg"), g_CogImage(i))
                Next i
            End If
        Else
            ResultString = "OK"
            If bOKimageSave = True Then
                Call Create_DIR(ImageFolderName & "OK")
                For i = 0 To 3
                    Call SaveCogImage(ImageFolderName & "OK" & "\" & sMESDate & "_" & sMesTime & "_" & sIDCode(0) & "_" & "CAM" & CStr(i + 1) & IIf(iImageFileMode = 1, ".bmp", ".jpg"), g_CogImage(i))
                Next i
            End If
        End If
        If bCamPass = True Then
            ResultString = "Pass"
        End If
        
        '스펙 및 PV 메세지 전송
        Call MES_DATASEND_FUNC("QMS_EVENT", "", "")
        
        '결과 출력
        Call Terminal_WriteDataToGrid(0)
        
        '카운트 저장
        Call frmMain.Counter(ResultString)
        
        '결과 표시
        frmMain.lblResults.Caption = ResultString

        
        tackJudgement = GetTickCount() - starttime
        
        
        '결과 화면 스크린샷 후 이미지 저장
        'Call Sleep(100)
        DoEvents
        Create_DIR "D:\MES\SEND"
        sMesSendJPGPath = "D:\MES\SEND\" & sIDCode(0) & "_" & sMESEquipCode & "_" & 1 & "_" & sDateTimeCheck & "_001" & ".JPG"
        'sMesSendJPGPath = sMESFileSendPath & sIDCode(0) & "_" & sMESEquipCode & "_" & 1 & "_" & sDateTimeCheck & "_001" & ".JPG"
        
        If g_SaveResultImage = 1 Then
            Call SH_ScreenSave(sMesSendJPGPath, ImageFolderName & sIDCode(0) & "_" & sMESEquipCode & "_" & 1 & "_" & sDateTimeCheck & "_001" & ".JPG")
        Else
            Call SH_ScreenSave(sMesSendJPGPath)
        End If
        

        'QCP 파일 백업
        sDataTemp = DJ_DataFileADD(0)
        Call DataFileSave(0, sDataTemp, ImageFolderName & sIDCode(0) & "_" & sMESEquipCode & "_" & 1 & "_" & sDateTimeCheck & ".QCP")     '저장되는 데이터 생성
        
        tackScreenshot = GetTickCount() - starttime
        
        
        '넷드라이브 전송
        'Call MES_NetDriveConnect

        tackMES = GetTickCount() - starttime

        endtime = GetTickCount
        frmMain.lblInspecTime.Caption = CStr(endtime - starttime) & " ㎳"
        
        '하드체크
        Call SH_HDDCheking(1)
        
        tackEnd = GetTickCount() - starttime
        
        LOG_TACK sIDCode(0), tackPLCSpec, tackGrab, tackinspection, tackJudgement, tackScreenshot, tackMES, tackEnd
        
        '설비상태 IDLE
        Call MES_DATASEND_FUNC("EQ_STATE_EVENT", "AUTO", "")
        
        '출력 신호 클리어
        Call ClearMelsecResult
        Call SendSignalToMelsec(outReadyVision, 1)
        Call WriteMelsec(frmMain.ActEasyIF, sMelsecAddrAcqDone, 0)
    Loop
    
    Call ClearMelsecResult
    Call SendSignalToMelsec(outReadyVision, 0)
    
    Exit Sub
ErrorHandle:
    
End Sub

Public Function PreWelding_Judgement() As Boolean

    Dim i As Integer
    
    g_NGCount = 0
    
    'W/H
    For i = 0 To 3
        g_Distance(i) = Format(g_Distance(i), "#0.00")
        frmMain.lblResultWH(i).Caption = g_Distance(i)
        
        g_Judge(i) = True
        If Range(g_Distance(i), dSpecOriMin(i), dSpecOriMax(i)) = False And bSpecPass(i) = False Then
            g_NGCount = g_NGCount + 1
            g_Judge(i) = False
        End If
    Next i
    
    'NSD
    For i = 0 To 5
        g_NsdDistance(i) = Format(g_NsdDistance(i), "#0.00")
        frmMain.lblResultNSD(i).Caption = Format(g_NsdDistance(i), "#0.00")
        
        g_Judge(i + 4) = True
        If Range(g_NsdDistance(i), dSpecOriMin(i + 4), dSpecOriMax(i + 4)) = False And bSpecPass(i + 4) = False Then
            g_NGCount = g_NGCount + 1
            g_Judge(i + 4) = False
        End If
    Next i
    
End Function

Public Function PreWelding_RunBlob(Optional Display As CogDisplay = Nothing) As Boolean
On Error GoTo ErrorHandle
    
    Dim i As Integer

    Set g_CogBlobTool.InputImage = g_CogImage(g_CogBlobIndex)
    Set g_CogBlobTool.Region = g_CogBlobRegion
    
    g_CogBlobTool.Run
    
    If g_CogBlobTool.Results Is Nothing Then
        PreWelding_RunBlob = False
        Exit Function
    End If
    
    If g_CogBlobTool.Results.Blobs.Count <= 0 Then
        PreWelding_RunBlob = False
        Exit Function
    End If
    
    PreWelding_RunBlob = True
    
    If Display Is Nothing Then
        Exit Function
    End If
    
    For i = 0 To g_CogBlobTool.Results.Blobs.Count - 1
        Dim CX As Double
        Dim CY As Double
        Dim Area As Double
        
        Display.StaticGraphics.Add g_CogBlobTool.Results.Blobs.Item(i).GetBoundary
        
        CX = g_CogBlobTool.Results.Blobs.Item(i).CenterOfMassX
        CY = g_CogBlobTool.Results.Blobs.Item(i).CenterOfMassY
        Area = g_CogBlobTool.Results.Blobs.Item(i).Area
        
        CogDisplayLabel Display, CX, CY, CStr(Area), cogColorYellow
    Next i
    
    Exit Function
ErrorHandle:
    MsgBox "PreWelding_RunBlob() - Error!"
    
End Function

Public Function PreWelding_RunNSD(Display1 As CogDisplay, Display2 As CogDisplay, Display3 As CogDisplay, Display4 As CogDisplay, Optional Display5 As CogDisplay = Nothing) As Boolean
On Error GoTo ErrorHandle
    
    Dim i As Integer
    
    Dim Display As CogDisplay
    Dim Index(0 To 5)
    
    g_Judge(11) = PreWelding_RunBlob(Display5)
    If g_Judge(11) = False Then
        For i = 0 To 5
            g_NsdDistance(i) = 0#
        Next i
        Exit Function
    End If
    
    Index(0) = 0
    Index(1) = 1
    Index(2) = 2
    Index(3) = 3
    Index(4) = 1
    Index(5) = 3
    
    If g_NsdRegionSelection = 1 Then
        Index(0) = 1
        Index(2) = 3
    End If
    
    For i = 0 To 11 Step 2
        If i = 2 Or i = 6 Then
            GoTo Continue
        End If
        g_NsdDistance(i / 2) = 0#
        
        Select Case Index(i / 2)
        Case 0
            Set Display = Display1
        Case 1
            Set Display = Display2
        Case 2
            Set Display = Display3
        Case 3
            Set Display = Display4
        End Select
        
        Set g_CogNsdTool(i + 0).InputImage = g_CogImage(Index(i / 2))
        Set g_CogNsdTool(i + 1).InputImage = g_CogImage(Index(i / 2))

        If i < 8 Then
            g_NsdDistance(i / 2) = CogFindCaliperY(g_CogNsdTool(i + 0), g_CogNsdTool(i + 1), Display, dCaliPX(Index(i / 2)), dSpecOffset(i / 2 + 4))
        Else
            g_NsdDistance(i / 2) = CogFindCaliperX(g_CogNsdTool(i + 0), g_CogNsdTool(i + 1), Display, dCaliPX(Index(i / 2)), dSpecOffset(i / 2 + 4))
        End If
Continue:
        
    Next i
    
    Exit Function
ErrorHandle:
    
End Function

Public Function PreWelding_RunWidthHeight(Display1 As CogDisplay, Display2 As CogDisplay, Display3 As CogDisplay, Display4 As CogDisplay) As Boolean
On Error GoTo ErrorHandle

    Dim ResultPoint1 As Double
    Dim ResultPoint2 As Double
    
    Dim Distance As Double
    
    Dim Distance1 As Double
    Dim Distance2 As Double
    
    Dim ToolIdx As Integer
    Dim i As Integer
    
    CogDisplayClear Display1
    CogDisplayClear Display2
    CogDisplayClear Display3
    CogDisplayClear Display4
    
    
    Dim Index As Integer
    
    Set g_CogCaliperTool(0).InputImage = g_CogImage(0)
    Set g_CogCaliperTool(1).InputImage = g_CogImage(1)
    
    For i = 0 To 3
        g_Distance(i) = -1#
    Next i
    
    Index = 0
    g_CogCaliperTool(Index).Run
    If g_CogCaliperTool(Index).Results.Count <= 0 Then
        Exit Function
    End If
    Display1.StaticGraphics.Add g_CogCaliperTool(Index).Results.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
    ResultPoint1 = g_CogCaliperTool(Index).Results.Item(0).PositionX
    Distance1 = ((XRES / 2) - ResultPoint1) * dCaliPX(0)
    
    Index = 1
    g_CogCaliperTool(Index).Run
    If g_CogCaliperTool(Index).Results.Count <= 0 Then
        Exit Function
    End If
    Display2.StaticGraphics.Add g_CogCaliperTool(Index).Results.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
    ResultPoint2 = g_CogCaliperTool(Index).Results.Item(0).PositionX
    Distance2 = (ResultPoint2 - (XRES / 2)) * dCaliPX(1)
    
    Distance = g_CameraGrap(0) + (Distance1 + Distance2) + dSpecOffset(0)
    
    g_Distance(0) = Distance
    
    CogDisplayLabel Display1, 200, 200, "너비1 = " & Format(Distance, "#0.00") & "㎜", cogColorGreen, "Tahoma", 16, True
        
    
    Set g_CogCaliperTool(2).InputImage = g_CogImage(2)
    Set g_CogCaliperTool(3).InputImage = g_CogImage(3)
    
    Index = 2
    g_CogCaliperTool(Index).Run
    If g_CogCaliperTool(Index).Results.Count <= 0 Then
        Exit Function
    End If
    Display3.StaticGraphics.Add g_CogCaliperTool(Index).Results.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
    ResultPoint1 = g_CogCaliperTool(Index).Results.Item(0).PositionX
    Distance1 = ((XRES / 2) - ResultPoint1) * dCaliPX(2)
    
    Index = 3
    g_CogCaliperTool(Index).Run
    If g_CogCaliperTool(Index).Results.Count <= 0 Then
        Exit Function
    End If
    Display4.StaticGraphics.Add g_CogCaliperTool(Index).Results.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
    ResultPoint2 = g_CogCaliperTool(Index).Results.Item(0).PositionX
    Distance2 = (ResultPoint2 - (XRES / 2)) * dCaliPX(3)
    
    Distance = g_CameraGrap(1) + (Distance1 + Distance2) + dSpecOffset(1)
    g_Distance(1) = Distance
    
    CogDisplayLabel Display2, 200, 200, "너비2 = " & Format(Distance, "#0.00") & "㎜", cogColorGreen, "Tahoma", 16, True
    

    Set g_CogCaliperTool(4).InputImage = g_CogImage(0)
    Set g_CogCaliperTool(5).InputImage = g_CogImage(2)
    
    Index = 4
    g_CogCaliperTool(Index).Run
    If g_CogCaliperTool(Index).Results.Count <= 0 Then
        Exit Function
    End If
    Display1.StaticGraphics.Add g_CogCaliperTool(Index).Results.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
    ResultPoint1 = g_CogCaliperTool(Index).Results.Item(0).PositionY
    Distance1 = ((YRES / 2) - ResultPoint1) * dCaliPXY(0)
    
    Index = 5
    g_CogCaliperTool(Index).Run
    If g_CogCaliperTool(Index).Results.Count <= 0 Then
        Exit Function
    End If
    Display3.StaticGraphics.Add g_CogCaliperTool(Index).Results.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
    ResultPoint2 = g_CogCaliperTool(Index).Results.Item(0).PositionY
    Distance2 = (ResultPoint2 - (YRES / 2)) * dCaliPXY(2)
    
    Distance = g_CameraGrap(2) + (Distance1 + Distance2) + dSpecOffset(2)
    g_Distance(2) = Distance
    
    CogDisplayLabel Display3, 200, 200, "높이1 = " & Format(Distance, "#0.00") & "㎜", cogColorGreen, "Tahoma", 16, True
    
    
    Set g_CogCaliperTool(6).InputImage = g_CogImage(1)
    Set g_CogCaliperTool(7).InputImage = g_CogImage(3)
    
    Index = 6
    g_CogCaliperTool(Index).Run
    If g_CogCaliperTool(Index).Results.Count <= 0 Then
        Exit Function
    End If
    Display2.StaticGraphics.Add g_CogCaliperTool(Index).Results.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
    ResultPoint1 = g_CogCaliperTool(Index).Results.Item(0).PositionY
    Distance1 = ((YRES / 2) - ResultPoint1) * dCaliPXY(1)
    
    Index = 7
    g_CogCaliperTool(Index).Run
    If g_CogCaliperTool(Index).Results.Count <= 0 Then
        Exit Function
    End If
    Display4.StaticGraphics.Add g_CogCaliperTool(Index).Results.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
    ResultPoint2 = g_CogCaliperTool(Index).Results.Item(0).PositionY
    Distance2 = (ResultPoint2 - (YRES / 2)) * dCaliPXY(3)
    
    Distance = g_CameraGrap(3) + (Distance1 + Distance2) + dSpecOffset(3)
    g_Distance(3) = Distance
    
    CogDisplayLabel Display4, 200, 200, "높이2 = " & Format(Distance, "#0.00") & "㎜", cogColorGreen, "Tahoma", 16, True

    Exit Function
ErrorHandle:

End Function
Public Function JR_ManualRun(CamIndex As Integer, Optional ByRef Display As CogDisplay = Nothing) As Boolean
On Error GoTo ErrorHandle
    
'    Dim Distance(0 To 3) As Double
'
'    Dim ToolIdx As Integer
'    Dim i As Integer
'    Dim j As Integer
'
'    For i = 0 To kMaxTool - 1 Step 4
'        For j = 0 To 3
'            Set g_CogCaliperTool(CamIndex, i + j).InputImage = g_CogImage(CamIndex)
'        Next j
'        Distance(i / 4) = CogFindDistance(g_CogCaliperTool(CamIndex, i + 0), g_CogCaliperTool(CamIndex, i + 1), g_CogCaliperTool(CamIndex, i + 2), g_CogCaliperTool(CamIndex, i + 3), Display, dCaliPX(CamIndex), dSpecOffset(CamIndex * 10 + (i / 4)))
'        g_Distance(CamIndex, i / 4) = Distance(i / 4)
'    Next i
'
'    For i = 0 To 3
'        frmMain.lblResultData(CamIndex * 10 + i).Caption = Format(Distance(i), "#0.00")
'        If Distance(i) < dSpecOriMin(i) Or Distance(i) > dSpecOriMax(i) Then
'            bResultJudge(CamIndex) = False
'        Else
'            bResultJudge(CamIndex) = True
'        End If
'    Next i

    JR_ManualRun = True
  
    Exit Function
ErrorHandle:
    JR_ManualRun = False
    
End Function

Public Sub Terminal_InitGrid()
    Dim i As Integer
Dim j As Integer
Dim stemp As String
Dim str_BlobName As String
    frmMain.MSFlexGrid1.WordWrap = True                    '한 Cell 에 두줄 쓸수 있게 됨
    frmMain.MSFlexGrid1.AllowUserResizing = flexResizeBoth 'Cell Size 를 마우스로 조절 할수 있음
    frmMain.MSFlexGrid1.SelectionMode = flexSelectionByRow
    
    For i = 1 To 10
        'stemp = stemp & "^" & sSpecName(i - 1) & vbCrLf & dSpecOriMin(i - 1) & "~" & dSpecOriMax(i - 1) & "         |"
        stemp = stemp & "^" & sSpecName(i - 1) & Chr(13) & dSpecOriMin(i - 1) & "~" & dSpecOriMax(i - 1)
        
        If i < 10 Then
            stemp = stemp & "|"
        End If
    Next i
    
    '검사결과 Display
    frmMain.MSFlexGrid1.Rows = 1
    frmMain.MSFlexGrid1.Cols = 4 + 10
    frmMain.MSFlexGrid1.FormatString = "^Number    |" & "^검사시간                |" & "^ID_Code                    |" & "^판정  |" & stemp
        
    frmMain.MSFlexGrid1.RowHeight(0) = 600
    frmMain.MSFlexGrid1.ColWidth(0) = 800
    frmMain.MSFlexGrid1.ColWidth(1) = 1400
    frmMain.MSFlexGrid1.ColWidth(2) = 4000
    frmMain.MSFlexGrid1.ColWidth(3) = 500
        
    For j = 0 To 9
        frmMain.MSFlexGrid1.ColWidth(4 + j) = 1500
    Next j
    
End Sub

Public Sub Terminal_WriteDataToGrid(Index As Integer)
Dim i As Integer
Dim Rownum As Long
Dim tempstr As String

    If g_NGCount = 0 And sIDCode(Index) <> "NOID" Then
        tempstr = "OK"
    Else
        tempstr = "NG"
    End If
    
    frmMain.MSFlexGrid1.AddItem "", 1
    
    Rownum = frmMain.MSFlexGrid1.Rows
    If frmMain.MSFlexGrid1.Rows > 3000 Then
        frmMain.MSFlexGrid1.Rows = 2
    End If
    
    Rownum = 2
    
    frmMain.MSFlexGrid1.TextMatrix(Rownum - 1, 0) = lInspectionNum
    lInspectionNum = lInspectionNum + 1
    frmMain.MSFlexGrid1.TextMatrix(Rownum - 1, 1) = Format(Time, "hh:mm:ss") & Format(Timer - Int(Timer), ".00")
    frmMain.MSFlexGrid1.TextMatrix(Rownum - 1, 2) = sIDCode(Index)
    frmMain.MSFlexGrid1.TextMatrix(Rownum - 1, 3) = sZigID
    frmMain.MSFlexGrid1.TextMatrix(Rownum - 1, 4) = tempstr
    
    '폭 & 높이
    For i = 0 To 3
        frmMain.MSFlexGrid1.TextMatrix(Rownum - 1, 5 + i) = Format(g_Distance(i), "#0.00")
    Next i
    
    'NSD
    For i = 0 To 5
        frmMain.MSFlexGrid1.TextMatrix(Rownum - 1, 9 + i) = Format(g_NsdDistance(i), "#0.00")
    Next i

    frmMain.MSFlexGrid1.Row = Rownum - 1
    
    frmMain.MSFlexGrid1.Col = 2
    If sIDCode(Index) = "NOID" Then
        frmMain.MSFlexGrid1.CellForeColor = vbRed
    End If
    
    frmMain.MSFlexGrid1.Col = 3
    If sZigID = "NOID" Then
        frmMain.MSFlexGrid1.CellForeColor = vbRed
    End If
    
    frmMain.MSFlexGrid1.Col = 4
    If tempstr = "OK" Then
        frmMain.MSFlexGrid1.CellForeColor = vbBlue
    Else
        frmMain.MSFlexGrid1.CellForeColor = vbRed
    End If
    
    For i = 0 To 9
        If g_Judge(i) = False Then
            frmMain.MSFlexGrid1.Col = 5 + i
            frmMain.MSFlexGrid1.CellForeColor = vbRed
        End If
    Next i

End Sub

Public Sub LoadCameraPosition(ModelName As String)
On Error Resume Next

    Dim FileName As String
    Dim Section As String
    
    FileName = App.Path & "\model\" & ModelName & "\model.ini"
    
    Section = "CAMERA_POSITION"
    g_CameraGrap(0) = CDbl(ReadINI(FileName, Section, "Width1"))
    g_CameraGrap(1) = CDbl(ReadINI(FileName, Section, "Width2"))
    g_CameraGrap(2) = CDbl(ReadINI(FileName, Section, "Height1"))
    g_CameraGrap(3) = CDbl(ReadINI(FileName, Section, "Height2"))
    
    Section = "CAMERA_POSITION_PRODUCT"
    g_ProductPt(0) = CDbl(ReadINI(FileName, Section, "Width1"))
    g_ProductPt(1) = CDbl(ReadINI(FileName, Section, "Width2"))
    g_ProductPt(2) = CDbl(ReadINI(FileName, Section, "Height1"))
    g_ProductPt(3) = CDbl(ReadINI(FileName, Section, "Height2"))
    
End Sub

Public Sub SaveCameraPosition(ModelName As String)
On Error Resume Next

    Dim FileName As String
    Dim Section As String
    
    FileName = App.Path & "\model\" & ModelName & "\model.ini"
    
    Section = "CAMERA_POSITION"
    Call WriteINI(FileName, Section, "Width1", CStr(g_CameraGrap(0)))
    Call WriteINI(FileName, Section, "Width2", CStr(g_CameraGrap(1)))
    Call WriteINI(FileName, Section, "Height1", CStr(g_CameraGrap(2)))
    Call WriteINI(FileName, Section, "Height2", CStr(g_CameraGrap(3)))
    
    Section = "CAMERA_POSITION_PRODUCT"
    Call WriteINI(FileName, Section, "Width1", CStr(g_ProductPt(0)))
    Call WriteINI(FileName, Section, "Width2", CStr(g_ProductPt(1)))
    Call WriteINI(FileName, Section, "Height1", CStr(g_ProductPt(2)))
    Call WriteINI(FileName, Section, "Height2", CStr(g_ProductPt(3)))
End Sub


Public Sub LoadModel(ModelName As String)
On Error Resume Next

    Call Calibration_Load(ModelName)
    Call LoadCameraPosition(ModelName)
    Call SpecName_Load(ModelName)
    Call SpecAllValue_Load(ModelName)
    Call FunctionValue_Load(ModelName)
    Call LoadResultSaving(ModelName)
    Call LoadCogTool(ModelName)
    
    '옵셋
    For i = 0 To 9
        frmMain.txtOffset(i).Text = Format(dSpecOffset(i), "#0.00")
    Next i
End Sub

Public Sub SaveModel(ModelName As String)
On Error Resume Next
    
    Call Calibration_Save(ModelName)
    Call SaveCameraPosition(ModelName)
    Call SpecName_Save(ModelName)
    Call SpecAllValue_Save(ModelName)
    Call FunctionValue_Save(ModelName)
    Call SaveResultSaving(ModelName)
    Call SaveCogTool(ModelName)
    
End Sub

Public Sub Terminal_SpecPrint()

    Dim i As Integer
    
    If sIDCode(0) = "" Then
        sIDCode(0) = "NOID"
    End If
    
    frmMain.lblIDCodeNum(0).Caption = sIDCode(0)
    
    If sZigID = "" Then
        sZigID = "NOID"
    End If
    
    frmMain.lblIDCodeNum(1).Caption = sZigID
    
    frmMain.grdSpec.Rows = 1
    For i = 1 To 10
        Call frmMain.grdSpec.AddItem("0", i)
        frmMain.grdSpec.TextMatrix(i, 0) = sSpecName(i - 1)
        frmMain.grdSpec.TextMatrix(i, 1) = Format(dSpecMin(i - 1), "#0.00")
        frmMain.grdSpec.TextMatrix(i, 2) = Format(dSpecOri(i - 1), "#0.00")
        frmMain.grdSpec.TextMatrix(i, 3) = Format(dSpecMax(i - 1), "#0.00")
    Next i
    
    Call frmMain.grdSpec.RemoveItem(8)
    Call frmMain.grdSpec.RemoveItem(6)
    
End Sub
