Attribute VB_Name = "commonCogTool"
Public Enum CogDisplayClearContants
    cogDisplayClearInteractive = 0
    cogDisplayClearStatic = 1
    cogDisplayClearAll = 2
End Enum

Public Enum CogImageFileTypeContants
    cogImageFileTypeBmp = 0
    cogImageFileTypeJpg = 1
End Enum

Public Function IDS_AcquireCognex(ByRef IDSCam As uEyeCam, Optional ByRef Display As CogDisplay = Nothing, Optional ByRef CalibrationTool As CogCalibCheckerboardTool = Nothing) As CogImage8Grey
On Error GoTo ErrorHandle

    Dim tme As IDisposable
    
    Dim IdsImage As New CogImage8Grey
    Dim Buffer As ICogImage8RootBuffer
    
    Dim nID As Integer
    
    Dim ImageMem As Long
    
    If IDSCam Is Nothing Then
        Set IDS_AcquireCognex = Nothing
        Exit Function
    End If
    
    If Not Display Is Nothing Then
        Set Display.Image = Nothing
    End If
    
    IDSCam.SetErrorReport 0
    nID = IDSCam.GetCameraID
    If nID < 0 Then
        Set IDS_AcquireCognex = Nothing
        
        Exit Function
    End If
    
    Call IDSCam.FreezeImage(1)
    ImageMem = IDSCam.GetImageMem()
    'fvImageBuf(nID - 1) = ImageMem

    Set Buffer = New CogImage8Root
    Buffer.Initialize XRES, YRES, ImageMem, XRES, tme
        
    IdsImage.SetRoot Buffer
    
    If CalibrationTool Is Nothing Then
        Set IDS_AcquireCognex = IdsImage.Copy
    Else
        Set CalibrationTool.InputImage = IdsImage
        CalibrationTool.Run
        Set IDS_AcquireCognex = CalibrationTool.OutputImage
    End If
    
    If Not Display Is Nothing Then
        Call Display.InteractiveGraphics.Clear
        Call Display.StaticGraphics.Clear
        Set Display.Image = IDS_AcquireCognex
    End If

    Exit Function
ErrorHandle:
    
    Set IDS_AcquireCognex = Nothing

End Function

Public Function IDS_AcquireCognexColor(ByRef IDSCam As uEyeCam, Optional _
    ByRef Display As CogDisplay = Nothing) As CogImage8Grey
On Error GoTo ErrorHandle

    Dim tme As IDisposable
    
    Dim IdsImage As New CogImage24PlanarColor
    Dim Buffer As ICogImage8RootBuffer
    
    Dim nID As Integer
    
    Dim ImageMem As Long
    
    If IDSCam Is Nothing Then
        Set IDS_AcquireCognexColor = Nothing
        Exit Function
    End If
    
    nID = IDSCam.GetCameraID
    Call IDSCam.FreezeImage(1)
    ImageMem = IDSCam.GetImageMem()
    fvImageBuf(nID - 1) = ImageMem

    Set Buffer = New CogImage8Root
    Buffer.Initialize XRES, YRES, ImageMem, XRES, tme
        
    IdsImage.SetRoots Buffer, Buffer, Buffer
          
    Set IDS_AcquireCognexColor = IdsImage.Copy
    
    If Not Display Is Nothing Then
        Call Display.InteractiveGraphics.Clear
        Call Display.StaticGraphics.Clear
        Set Display.Image = IdsImage
    End If

    Exit Function
ErrorHandle:
    
    Set IDS_AcquireCognexColor = Nothing

End Function

Public Function CogDisplayClear(ByRef Display As CogDisplay, Optional _
    ByVal OptionA As CogDisplayClearContants = cogDisplayClearAll) As _
    Boolean
   
    If Display Is Nothing Then
        CogDisplayClear = False
        Exit Function
    End If
    
    If OptionA = cogDisplayClearInteractive Or cogDisplayClearAll Then
        Display.InteractiveGraphics.Clear
    End If
    
    If OptionA = cogDisplayClearStatic Or cogDisplayClearAll Then
        Display.StaticGraphics.Clear
    End If
    
    CogDisplayClear = True
    
End Function

Public Sub CogDisplayRectangle(ByRef Display As CogDisplay, ByRef RectangleAffine As CogRectangleAffine, Optional Interactive As Boolean = False, Optional Color As CogColorConstants = cogColorGreen)

    If Display Is Nothing Then
        Exit Sub
    End If
    
    If RectangleAffine Is Nothing Then
        Exit Sub
    End If
    
    Dim Shape As New CogRectangleAffine
    
    Set Shape = RectangleAffine.Copy
    Shape.Color = Color
    
    If Interactive = True Then
        Display.StaticGraphics.Add Shape
    Else
        Display.InteractiveGraphics.Add Shape
    End If
    
End Sub

Public Sub CogDisplaySegment(ByRef Display As CogDisplay, ByRef Segment As CogLineSegment, Optional Interactive As Boolean = False, Optional Color As CogColorConstants = cogColorGreen)

    If Display Is Nothing Then
        Exit Sub
    End If
    
    If Segment Is Nothing Then
        Exit Sub
    End If
    
    Dim Shape As New CogLineSegment
    
    Set Shape = Segment.Copy
    Shape.Color = Color
    
    If Interactive = True Then
        Display.StaticGraphics.Add Shape
    Else
        Display.InteractiveGraphics.Add Shape
    End If
    
End Sub

Public Sub CogDisplayPoint(ByRef Display As CogDisplay, CX As Double, CY As Double, Optional Color As CogColorConstants = cogColorRed)

    If Display Is Nothing Then
        Exit Sub
    End If
    
    Dim Point As New CogPointMarker
    
    Point.SetCenterRotationSize CX, CY, 0, 12
    Point.Color = Color
    Point.GraphicType = cogPointMarkerGraphicTypeCross
    
    Display.StaticGraphics.Add Point
    
End Sub

Public Sub CogDisplayLabel(ByRef Display As CogDisplay, CX As Double, CY As Double, Message As String, Optional Color As CogColorConstants = cogColorGreen, Optional FontName As String = "Tahoma", Optional FontSize As Long = 12, Optional AlignLeft As Boolean = False)

    If Display Is Nothing Then
        Exit Sub
    End If
    
    Dim Label As New CogGraphicLabel
    
    Label.Font.Bold = True
    Label.Font.Name = FontName
    Label.Font.size = FontSize
    Label.Alignment = cogGraphicLabelAlignmentBaselineCenter
    If AlignLeft = True Then
        Label.Alignment = cogGraphicLabelAlignmentBaselineLeft
    End If
    
    Label.Color = Color
    Label.SetXYText CX, CY, Message
        
    Display.StaticGraphics.Add Label

End Sub

Public Function CogDisplayTrainedPattern(ByRef Tool As CogPMAlignTool, ByRef Display As CogDisplay) As Boolean

    CogDisplayTrainedPattern = False
    
    If Tool Is Nothing Or Display Is Nothing Then
        Exit Function
    End If
    
    If Tool.Pattern.Trained = False Then
        Set Display.Image = Nothing
        Exit Function
    End If
    
    Set Display.Image = Tool.Pattern.GetTrainedPatternImage
    
    Dim GraphicsFine As CogGraphicCollection
    Dim GraphicsCoarse  As CogGraphicCollection
    Set GraphicsFine = Tool.Pattern.CreateGraphicsFine(cogColorGreen)
    Set GraphicsCoarse = Tool.Pattern.CreateGraphicsCoarse(cogColorOrange)
    
    Display.StaticGraphics.AddList GraphicsFine
    Display.StaticGraphics.AddList GraphicsCoarse
    
    CogDisplayPoint Display, Tool.Pattern.Origin.TranslationX, Tool.Pattern.Origin.TranslationY, cogColorRed
    
    CogDisplayTrainedPattern = True
    
End Function

Public Function CogMakeMaskImage(ByRef Image As CogImage8Grey, ByRef _
    GraphicCollection As CogGraphicCollection) As CogImage8Grey
On Error GoTo ErrorHandle

    Dim tme As IDisposable
    
    Dim ReturnImage As New CogImage8Grey
    Dim IdsImage As New CogImage8Grey
    Dim Buffer As ICogImage8RootBuffer

    Dim ImageSize As Long
    Dim ImageMem() As Byte
    
    Dim PtrImage As Long
    
    ImageSize = Image.Width * Image.Height
     
    ReDim ImageMem(ImageSize - 1)
    
    For i = 0 To ImageSize - 1
        ImageMem(i) = 255
    Next i
    
    PtrImage = VarPtr(ImageMem(0))
    
    Set Buffer = New CogImage8Root
    Buffer.Initialize XRES, YRES, PtrImage, XRES, tme
    IdsImage.SetRoot Buffer
    
    Dim CopyTool As New CogCopyRegionTool
    
    Set CopyTool.InputImage = Image
    Set CopyTool.DestinationImage = IdsImage
    CopyTool.RunParams.RegionMode = _
        cogRegionModePixelAlignedBoundingBoxAdjustMask
    CopyTool.RunParams.FillBoundingBox = False
    CopyTool.RunParams.FillRegion = True
    CopyTool.RunParams.FillRegionValue = 0
    CopyTool.RunParams.ImageAlignmentEnabled = True
    
    For i = 0 To GraphicCollection.Count - 1
        Set CopyTool.Region = GraphicCollection.Item(i)
        CopyTool.Run
    Next i
    
    Set ReturnImage = CopyTool.OutputImage
    Set CogMakeMaskImage = ReturnImage.Copy
    
    Exit Function
ErrorHandle:
    Set CogMakeMaskImage = Nothing

End Function

Public Function CogTrainPattern(ByRef Tool As CogPMAlignTool, _
    TrainImage As CogImage8Grey, TrainImageMask As CogImage8Grey, Region _
    As CogCircle) As Boolean
On Error GoTo ErrorHandler
    
    Set Tool.Pattern.TrainImage = TrainImage
    Set Tool.Pattern.TrainImageMask = TrainImageMask
    Set Tool.Pattern.TrainRegion = Region
    Tool.Pattern.TrainAlgorithm = cogPMAlignTrainAlgorithmPatMax
    Tool.Pattern.IgnorePolarity = True
    Tool.Pattern.AutoEdgeThresholdEnabled = False
    Tool.Pattern.EdgeThreshold = 10
    Tool.Pattern.GrainLimitAutoSelect = False
    Tool.Pattern.GrainLimitCoarse = 9
    Tool.Pattern.GrainLimitFine = 5
    Tool.Pattern.Origin.TranslationX = Region.CenterX
    Tool.Pattern.Origin.TranslationY = Region.CenterY
    
    Call Tool.Pattern.Train
    
    CogTrainPattern = True
    
    Exit Function
ErrorHandler:
    CogTrainPattern = False
    
End Function

Public Function CogTrainPatternRectangle(ByRef Tool As CogPMAlignTool, TrainImage As CogImage8Grey, Region As CogRectangle, Optional ByRef Origin As CogCoordinateAxes = Nothing) As Boolean
On Error GoTo ErrorHandler
    
    Set Tool.Pattern.TrainImage = TrainImage
    Set Tool.Pattern.TrainRegion = Region
    
    If Origin Is Nothing Then
        Tool.Pattern.Origin.TranslationX = Region.CenterX
        Tool.Pattern.Origin.TranslationY = Region.CenterY
        Tool.Pattern.Origin.Rotation = 0#
    Else
        Tool.Pattern.Origin.TranslationX = Origin.OriginX
        Tool.Pattern.Origin.TranslationY = Origin.OriginY
        Tool.Pattern.Origin.Rotation = Origin.Rotation
    End If
    
    Call Tool.Pattern.Train
    
    CogTrainPatternRectangle = Tool.Pattern.Trained
    
    Exit Function
ErrorHandler:
    CogTrainPatternRectangle = False
    
End Function

Public Function LoadCogImage(FilePath As String) As CogImage8Grey
On Error GoTo ErrorHandle

    Dim Image As CogImage8Grey
    Dim ImageFile As New CogImageFile
        
           
    Call ImageFile.Open(FilePath, cogImageFileModeRead)
    
    If ImageFile.Count <= 0 Then
        Set LoadCogImage = Nothing
        Exit Function
    End If

    Set Image = ImageFile.Item(0)
    Call ImageFile.Close
    
    Set LoadCogImage = Image.Copy
    
    Exit Function
ErrorHandle:

    Set LoadCogImage = Nothing
        
End Function

Public Function SaveCogImage(FilePath As String, Image As CogImage) As Boolean
On Error GoTo ErrorHandle

    Dim ImageFile As New CogImageFile
    
    Call ImageFile.Open(FilePath, cogImageFileModeWrite)
    Call ImageFile.Append(Image)
    Call ImageFile.Close
    
    SaveCogImage = True
    
    Exit Function
ErrorHandle:
    SaveCogImage = False

End Function

Public Function CogFindBlob(nCamIndex As Integer, ByRef Tool As CogBlobTool, ByRef Region As CogRegion, ByRef Image As CogImage8Grey, Optional ByRef Display As CogDisplay = Nothing) As Long

    Dim tempHole As Integer
    Dim tempArea As Double

    If Tool Is Nothing Then
        CogFindBlob = -1
        Exit Function
    End If
    
    Dim tmpSpaceName As String
    
    Set Tool.InputImage = Image
    Set Tool.Region = Region
    Tool.Run
    
    nBlobNGCount = 0
    
    tempHole = 0
    
    If Tool.Results Is Nothing Then
        CogFindBlob = 0
        Exit Function
    End If
    
    If Tool.Results.Blobs.Count <= 0 Then
        CogFindBlob = 0
        
        bResultJudge_Blob(nCamIndex) = False
        Exit Function
    Else
        CogFindBlob = Tool.Results.Blobs(False).Count
    End If
    
    If Display Is Nothing Then
        Exit Function
    End If
    
    Dim i As Integer
    
    
    
    For i = 0 To Tool.Results.Blobs.Count - 1
        Display.StaticGraphics.Add Tool.Results.Blobs.Item(i).GetBoundary
    Next i
    
    For i = 0 To CogFindBlob - 1
        If Tool.Results.Blobs(False).Item(i).FilteredOut = False Then
            tempArea = tempArea + Tool.Results.Blobs(False).Item(i).Area
        ElseIf Tool.Results.Blobs(False).Item(i).Label = cogBlobLabelHole Then
            tempHole = tempHole + 1
        End If
    Next i
    
    Debug.Print "홀 개수", tempHole
    
    If tempArea < dResultJudge_BlobArea(nCamIndex) Then
        bResultJudge_Blob(nCamIndex) = False
        
        nBlobNGCount = nBlobNGCount + 1
    Else
        bResultJudge_Blob(nCamIndex) = True
    End If
    
    dResultBlobArea(nCamIndex) = tempArea
    
    'frmSetting.lblBlobResult.Caption = dResultBlobArea(nCamIndex)
    
'    MsgBox (tempArea)
'    Region.Color = cogColorDarkGreen
'    Display.StaticGraphics.Add Region
'    Region.Color = cogColorGreen
    
End Function

Public Function CogFindEdge(ByRef Tool As CogCaliperTool, ByRef PX As Double, ByRef PY As Double, Optional ByRef Display As CogDisplay = Nothing) As Boolean
On Error GoTo ErrorHandle

    CogFindEdge = False
    
    If Tool Is Nothing Then
        Exit Function
    End If
    
    Tool.Run
    
    If Tool.Results.Count <= 0 Then
        Exit Function
    End If
    
    CogFindEdge = True
    
    PX = Tool.Results.Item(0).PositionX
    PY = Tool.Results.Item(0).PositionY
    
    If Display Is Nothing Then
        Exit Function
    End If
    
    Display.StaticGraphics.Add Tool.Results.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
        
    Exit Function
ErrorHandle:
    

End Function

Public Function CogFindLineAngle(ByRef Tool1 As CogFindLineTool, ByRef Tool2 As CogFindLineTool, Optional ByRef Display As CogDisplay = Nothing) As Double
On Error GoTo ErrorHandle

    If Tool1 Is Nothing Or Tool2 Is Nothing Then
        CogFindLineAngle = -1#
        Exit Function
    End If
        
    Dim Angle As Double
    
    Dim Line1 As CogLine
    Dim Line2 As CogLine
    
    Tool1.Run
    If Tool1.Results.Count < 2 Then
        CogFindLineAngle = -1#
        Exit Function
    End If
    Line1 = Tool1.Results.GetLine
    
    Tool2.Run
    If Tool2.Results.Count < 2 Then
        CogFindLineAngle = -1#
        Exit Function
    End If
    Line2 = Tool1.Results.GetLine
    
    Angle = CogMath.AngleLineLine(Line1, Line2, Tool1.InputImage)
    CogFindLineAngle = Angle
    
    If Display Is Nothing Then
        Exit Function
    End If
    
    '라인 출력
    Line1.Color = cogColorGreen
    Line2.Color = cogColorGreen
    Display.StaticGraphics.Add Line1
    Display.StaticGraphics.Add Line2
        
    Dim NumPoints As Long
    Dim X As Double
    Dim Y As Double
    
    '라인과 라인이 만나는 지점 구하기
    CogMath.IntersectLineLine Line1, Line2, Tool1.InputImage, NumPoints, X, Y
    '각도값 출력
    CogDisplayLabel Display, X, Y, Format(Angle, "#0.00"), cogColorGreen, "Tahoma", 16
    
    Exit Function
ErrorHandle:
    CogFindLineAngle = -1#
    
End Function

Public Function CogFindCaliper(ByRef Tool1 As CogCaliperTool, ByRef Tool2 As CogCaliperTool, Optional ByRef Display As CogDisplay = Nothing, Optional ByVal Calib As Double = 1#, Optional Offset As Double = 0#) As Double
On Error GoTo ErrorHandle

    Dim Distance As Double
    
    CogFindCaliper = -1#
    
    If Tool1 Is Nothing Or Tool2 Is Nothing Then
        Exit Function
    End If
    
    Dim r As CogCaliperResults
    Tool1.Run
    
    If Tool1.Results.Count <= 0 Then
        Exit Function
    End If
    
    Tool2.Run
    If Tool2.Results.Count <= 0 Then
        Exit Function
    End If
    
    Distance = CogMath.DistancePointPoint(Tool1.Results.Edges(0).PositionX, Tool1.Results.Edges(0).PositionY, Tool2.Results.Edges(0).PositionX, Tool2.Results.Edges(0).PositionY) * Calib
    
    CogFindCaliper = Distance
    
    If Display Is Nothing Then
        Exit Function
    End If
    
    Display.StaticGraphics.Add Tool1.Results.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
    Display.StaticGraphics.Add Tool2.Results.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
    
    Dim Segment As New CogLineSegment
    Segment.StartPointAdornment = cogLineSegmentAdornmentSolidArrow
    Segment.EndPointAdornment = cogLineSegmentAdornmentSolidArrow
    Segment.SetStartEnd Tool1.Results.Edges(0).PositionX, Tool1.Results.Edges(0).PositionY, Tool2.Results.Edges(0).PositionX, Tool2.Results.Edges(0).PositionY
    Segment.Color = cogColorOrange
    
    Display.StaticGraphics.Add Segment
    
    CogDisplayLabel Display, Segment.MidpointX, Segment.MidpointY, Format(Distance, "#0.00") & "㎜", cogColorGreen, "Tahoma", 16
'    Dim Label As New CogGraphicLabel
'
'    Label.Alignment = cogGraphicLabelAlignmentBaselineCenter
'    Label.Font.Name = "Tahoma"
'    Label.Font.Bold = True
'    Label.Font.size = 16
'    Label.Color = cogColorGreen
'    Label.SetXYText Segment.MidpointX, Segment.MidpointY, Format(Distance, "#0.00")
'
'    Display.StaticGraphics.Add Label
    
    Exit Function
ErrorHandle:
    CogFindCaliper = -1#
    
End Function

Public Function CogFindCaliperPointX(ByRef Tool As CogCaliperTool, ByVal PX As Double, ByVal PY As Double, Optional ByRef Display As CogDisplay = Nothing, Optional ByVal Calib As Double = 1#, Optional Offset As Double = 0#) As Double
On Error GoTo ErrorHandle

    Dim Distance As Double
    
    CogFindCaliperPointX = -1#
    
    If Tool Is Nothing Then
        Exit Function
    End If
    
    Tool.Run
    
    If Tool.Results.Count <= 0 Then
        Exit Function
    End If
    
    Dim Line1 As New CogLine
    Line1.SetXYRotation PX, 0, CogMisc.DegToRad(90)
    
    Dim X As Double
    Dim Y As Double
    
    Distance = CogMath.DistancePointLine(Tool.Results.Item(0).PositionX, Tool.Results.Item(0).PositionY, Line1, Tool.InputImage, X, Y) * Calib + Offset
    
    CogFindCaliperPointX = Distance
    
    If Display Is Nothing Then
        Exit Function
    End If
    
    Display.StaticGraphics.Add Tool.Results.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
    
    Dim Segment As New CogLineSegment
    Segment.StartPointAdornment = cogLineSegmentAdornmentSolidArrow
    Segment.EndPointAdornment = cogLineSegmentAdornmentSolidArrow
    Segment.SetStartEnd Tool.Results.Item(0).PositionX, Tool.Results.Item(0).PositionY, X, Y
    Segment.Color = cogColorOrange
    
    Display.StaticGraphics.Add Segment
    
    CogDisplayLabel Display, Segment.MidpointX, Segment.MidpointY, Format(Distance, "#0.00") & "㎜"
    
    Exit Function
ErrorHandle:
    CogFindCaliperPointX = -1#
    
End Function

Public Function CogFindCaliperX(ByRef Tool1 As CogCaliperTool, ByRef Tool2 As CogCaliperTool, Optional ByRef Display As CogDisplay = Nothing, Optional ByVal Calib As Double = 1#, Optional Offset As Double = 0#) As Double
On Error GoTo ErrorHandle

    Dim Distance As Double
    
    CogFindCaliperX = -1#
    
    If Tool1 Is Nothing Or Tool2 Is Nothing Then
        Exit Function
    End If
    
    Dim r As CogCaliperResults
    
    Tool1.Run
    Tool2.Run
    
    If Not Display Is Nothing Then
        Tool1.Region.Color = cogColorDarkGreen
        Tool2.Region.Color = cogColorDarkGreen
        
        Display.StaticGraphics.Add Tool1.Region
        Display.StaticGraphics.Add Tool2.Region
        
        Tool1.Region.Color = cogColorGreen
        Tool2.Region.Color = cogColorGreen
    End If
    
    If Tool1.Results.Count <= 0 Then
        Exit Function
    End If
    
    
    If Tool2.Results.Count <= 0 Then
        Exit Function
    End If
    
    Dim Line1 As New CogLine
    Line1.SetXYRotation Tool2.Results.Item(0).PositionX, 0, CogMisc.DegToRad(90)
    
    Dim X As Double
    Dim Y As Double
    
    'Distance = CogMath.DistancePointPoint(Tool1.Results.Edges(0).PositionX, Tool1.Results.Edges(0).PositionY, Tool2.Results.Edges(0).PositionX, Tool2.Results.Edges(0).PositionY) * Calib
    Distance = CogMath.DistancePointLine(Tool1.Results.Item(0).PositionX, Tool1.Results.Item(0).PositionY, Line1, Tool1.InputImage, X, Y) * Calib + Offset
    
    CogFindCaliperX = Distance
    
    If Display Is Nothing Then
        Exit Function
    End If
    
    Display.StaticGraphics.Add Tool1.Results.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
    Display.StaticGraphics.Add Tool2.Results.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
    
    Dim Segment As New CogLineSegment
    Segment.StartPointAdornment = cogLineSegmentAdornmentSolidArrow
    Segment.EndPointAdornment = cogLineSegmentAdornmentSolidArrow
    Segment.SetStartEnd Tool1.Results.Item(0).PositionX, Tool1.Results.Item(0).PositionY, X, Y
    Segment.Color = cogColorOrange
    
    Display.StaticGraphics.Add Segment
    
    CogDisplayLabel Display, Segment.MidpointX, Segment.MidpointY, Format(Distance, "#0.00") & "㎜", cogColorGreen, "Tahoma", 16
'    Dim Label As New CogGraphicLabel
'
'    Label.Alignment = cogGraphicLabelAlignmentBaselineCenter
'    Label.Font.Name = "Tahoma"
'    Label.Font.Bold = True
'    Label.Font.size = 16
'    Label.Color = cogColorGreen
'    Label.SetXYText Segment.MidpointX, Segment.MidpointY, Format(Distance, "#0.00")
'
'    Display.StaticGraphics.Add Label
    
    Exit Function
ErrorHandle:
    CogFindCaliperX = -1#
    
End Function

Public Function CogFindCaliperPointY(ByRef Tool As CogCaliperTool, ByVal PX As Double, ByVal PY As Double, Optional ByRef Display As CogDisplay = Nothing, Optional ByVal Calib As Double = 1#, Optional Offset As Double = 0#) As Double
On Error GoTo ErrorHandle

    Dim Distance As Double
    
    CogFindCaliperPointY = -1#
    
    If Tool Is Nothing Then
        Exit Function
    End If
    
    Dim r As CogCaliperResults
    Tool.Run
    
    If Tool.Results.Count <= 0 Then
        Exit Function
    End If
    
    Dim Line1 As New CogLine
    Line1.SetXYRotation 0, PY, 0
    
    Dim X As Double
    Dim Y As Double
    
    Distance = CogMath.DistancePointLine(Tool.Results.Item(0).PositionX, Tool.Results.Item(0).PositionY, Line1, Tool.InputImage, X, Y) * Calib + Offset
    
    CogFindCaliperPointY = Distance
    
    If Display Is Nothing Then
        Exit Function
    End If
    
    Display.StaticGraphics.Add Tool.Results.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
    
    Dim Segment As New CogLineSegment
    Segment.StartPointAdornment = cogLineSegmentAdornmentSolidArrow
    Segment.EndPointAdornment = cogLineSegmentAdornmentSolidArrow
    Segment.SetStartEnd Tool.Results.Item(0).PositionX, Tool.Results.Item(0).PositionY, X, Y
    Segment.Color = cogColorOrange
    
    Display.StaticGraphics.Add Segment
    
    CogDisplayLabel Display, Segment.MidpointX, Segment.MidpointY, Format(Distance, "#0.00") & "㎜"
    
    Exit Function
ErrorHandle:
    CogFindCaliperPointY = -1#
    
End Function




Public Function CogFindCaliperY(ByRef Tool1 As CogCaliperTool, ByRef Tool2 As CogCaliperTool, Optional ByRef Display As CogDisplay = Nothing, Optional ByVal Calib As Double = 1#, Optional Offset As Double = 0#) As Double
On Error GoTo ErrorHandle

    Dim Distance As Double
    
    CogFindCaliperY = -1#
    
    If Tool1 Is Nothing Or Tool2 Is Nothing Then
        Exit Function
    End If
    
    Dim r As CogCaliperResults
    
    Tool1.Run
    Tool2.Run
    
    If Not Display Is Nothing Then
        Tool1.Region.Color = cogColorDarkGreen
        Display.StaticGraphics.Add Tool1.Region
        
        Tool2.Region.Color = cogColorDarkGreen
        Display.StaticGraphics.Add Tool2.Region
        
        Tool1.Region.Color = cogColorGreen
        Tool2.Region.Color = cogColorGreen
    End If
    
    If Tool1.Results.Count <= 0 Then
        Exit Function
    End If
    
    If Tool2.Results.Count <= 0 Then
        Exit Function
    End If
    
    Dim Line1 As New CogLine
    Line1.SetXYRotation 0, Tool2.Results.Item(0).PositionY, 0
    
    Dim X As Double
    Dim Y As Double
    
    'Distance = CogMath.DistancePointPoint(Tool1.Results.Edges(0).PositionX, Tool1.Results.Edges(0).PositionY, Tool2.Results.Edges(0).PositionX, Tool2.Results.Edges(0).PositionY) * Calib
    Distance = CogMath.DistancePointLine(Tool1.Results.Item(0).PositionX, Tool1.Results.Item(0).PositionY, Line1, Tool1.InputImage, X, Y) * Calib + Offset
    
    CogFindCaliperY = Distance
    
    If Display Is Nothing Then
        Exit Function
    End If
    
    Display.StaticGraphics.Add Tool1.Results.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
    Display.StaticGraphics.Add Tool2.Results.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
    
    Dim Segment As New CogLineSegment
    Segment.StartPointAdornment = cogLineSegmentAdornmentSolidArrow
    Segment.EndPointAdornment = cogLineSegmentAdornmentSolidArrow
    Segment.SetStartEnd Tool1.Results.Item(0).PositionX, Tool1.Results.Item(0).PositionY, X, Y
    Segment.Color = cogColorOrange
    
    Display.StaticGraphics.Add Segment
    
    CogDisplayLabel Display, Segment.MidpointX, Segment.MidpointY, Format(Distance, "#0.00") & "㎜", cogColorGreen, "Tahoma", 16
'    Dim Label As New CogGraphicLabel
'
'    Label.Alignment = cogGraphicLabelAlignmentBaselineCenter
'    Label.Font.Name = "Tahoma"
'    Label.Font.Bold = True
'    Label.Font.size = 16
'    Label.Color = cogColorGreen
'    Label.SetXYText Segment.MidpointX, Segment.MidpointY, Format(Distance, "#0.00")
'
'    Display.StaticGraphics.Add Label
    
    Exit Function
ErrorHandle:
    CogFindCaliperY = -1#
    
End Function


Public Function CogFindDistance(ByRef Tool1 As CogCaliperTool, ByRef Tool2 As CogCaliperTool, ByRef Tool3 As CogCaliperTool, ByRef Tool4 As CogCaliperTool, Optional ByRef Display As CogDisplay = Nothing, Optional ByVal Calib As Double = 1#, Optional ByVal Offset As Double = 0#, Optional ByRef Label As CogGraphicLabel = Nothing) As Double
On Error GoTo ErrorHandle

    Dim Distance As Double
    Dim Distance1 As Double
    Dim Distance2 As Double
    
    CogFindDistance = -1#
    
    If Tool1 Is Nothing Or Tool2 Is Nothing Or Tool3 Is Nothing Or Tool4 Is Nothing Then
        Exit Function
    End If
    
    Tool1.Run
    If Tool1.Results.Count <= 0 Then
        Exit Function
    End If
    
    Tool2.Run
    If Tool2.Results.Count <= 0 Then
        Exit Function
    End If
    
    Tool3.Run
    If Tool1.Results.Count <= 0 Then
        Exit Function
    End If
    
    Tool4.Run
    If Tool1.Results.Count <= 0 Then
        Exit Function
    End If
    
    Dim Segment1 As New CogLineSegment
    Dim Segment2 As New CogLineSegment
    
    Dim Line1 As New CogLine
    Dim Line2 As New CogLine
    
    Dim SX As Double
    Dim SY As Double
    Dim EX As Double
    Dim EY As Double
    
    Segment1.Color = cogColorRed
    Segment1.SetStartEnd Tool1.Results.Item(0).PositionX, Tool1.Results.Item(0).PositionY, Tool2.Results.Item(0).PositionX, Tool2.Results.Item(0).PositionY
    
    Segment2.Color = cogColorRed
    Segment2.SetStartEnd Tool3.Results.Item(0).PositionX, Tool3.Results.Item(0).PositionY, Tool4.Results.Item(0).PositionX, Tool4.Results.Item(0).PositionY
    
    Line1.Color = cogColorDarkGreen
    Line1.SetFromStartXYEndXY Tool1.Results.Item(0).PositionX, Tool1.Results.Item(0).PositionY, Tool2.Results.Item(0).PositionX, Tool2.Results.Item(0).PositionY
    
    Line2.Color = cogColorDarkGreen
    Line2.SetFromStartXYEndXY Tool3.Results.Item(0).PositionX, Tool3.Results.Item(0).PositionY, Tool4.Results.Item(0).PositionX, Tool4.Results.Item(0).PositionY
    
'    Distance1 = CogMath.DistancePointLine(Tool1.Results.Item(0).PositionX, Tool1.Results.Item(0).PositionY, Line2, Tool1.InputImage, EX, EY)
'    Distance2 = CogMath.DistancePointLine(Tool2.Results.Item(0).PositionX, Tool2.Results.Item(0).PositionY, Line2, Tool1.InputImage)
    Distance1 = CogMath.DistancePointPoint(Tool1.Results.Item(0).PositionX, Tool1.Results.Item(0).PositionY, Tool3.Results.Item(0).PositionX, Tool3.Results.Item(0).PositionY)
    Distance2 = CogMath.DistancePointPoint(Tool2.Results.Item(0).PositionX, Tool2.Results.Item(0).PositionY, Tool4.Results.Item(0).PositionX, Tool4.Results.Item(0).PositionY)
    Distance = (Distance1 + Distance2) / 2 * Calib + Offset
    
    CogFindDistance = Distance
    
    If Display Is Nothing Then
        Exit Function
    End If
    
    Dim Point1 As New CogPointMarker
    Dim Point2 As New CogPointMarker
    Dim Point3 As New CogPointMarker
    Dim Point4 As New CogPointMarker
    
    Point1.Color = cogColorRed
    Point1.SetCenterRotationSize Tool1.Results.Item(0).PositionX, Tool1.Results.Item(0).PositionY, 0, 12
    Point2.Color = cogColorRed
    Point2.SetCenterRotationSize Tool2.Results.Item(0).PositionX, Tool2.Results.Item(0).PositionY, 0, 12
    Point3.Color = cogColorRed
    Point3.SetCenterRotationSize Tool3.Results.Item(0).PositionX, Tool3.Results.Item(0).PositionY, 0, 12
    Point4.Color = cogColorRed
    Point4.SetCenterRotationSize Tool4.Results.Item(0).PositionX, Tool4.Results.Item(0).PositionY, 0, 12
    
'    Display.StaticGraphics.Add Line1
'    Display.StaticGraphics.Add Line2
     
    Dim Segment As New CogLineSegment
    Segment.StartPointAdornment = cogLineSegmentAdornmentSolidArrow
    Segment.EndPointAdornment = cogLineSegmentAdornmentSolidArrow
    Segment.SetStartEnd Segment1.MidpointX, Segment1.MidpointY, Segment2.MidpointX, Segment2.MidpointY
    Segment.Color = cogColorOrange
    
    Display.StaticGraphics.Add Segment
    Display.StaticGraphics.Add Segment1
    Display.StaticGraphics.Add Segment2
    
    Display.StaticGraphics.Add Point1
    Display.StaticGraphics.Add Point2
    Display.StaticGraphics.Add Point3
    Display.StaticGraphics.Add Point4
    
    If Label Is Nothing Then
        CogDisplayLabel Display, Segment.MidpointX, Segment.MidpointY, Format(Distance, "#0.00") & IIf(Calib <> 1, " ㎜", "")
    Else
        Label.X = Segment.MidpointX
        Label.Y = Segment.MidpointY
        Label.Text = Format(Distance, "#0.00") & IIf(Calib <> 1, " ㎜", "")
        Label.Color = cogColorGreen
        Label.Font.Bold = True
        Label.Font.Name = "Tahoma"
        Label.Font.size = 12
    End If
    
    
    
    Exit Function
ErrorHandle:
    CogFindDistance = -1#
    
End Function

Public Function CogCaliperDiff(ByRef Tool As CogCaliperTool, ByRef Region As CogRectangleAffine, ByRef Image As CogImage8Grey, Optional ByRef Display As CogDisplay = Nothing, Optional ByVal Calib As Double = 1#, Optional ByVal Offset As Double = 0#) As Double


    Dim MaxX As Double
    Dim MinX As Double
    Dim MaxY As Double
    Dim MinY As Double
    
    Tool.Run
    
    If Tool.Results.Count <= 0 Then
        CogCaliperDiff = -1#
        Exit Function
    End If
    
    Dim i As Integer
    
    Dim PointMarker As New CogPointMarker
    PointMarker.Color = cogColorRed
    
    Dim Caliper As New CogCaliper
    Caliper.EdgeMode = cogCaliperEdgeModeSingle
    Caliper.Edge0Polarity = cogCaliperPolarityLightToDark
    Caliper.ContrastThreshold = Tool.RunParams.ContrastThreshold
    Tool.RunParams.MaxResults = 2
    
    Dim Scorer As New CogCaliperScorerPositionNeg
    Caliper.SingleEdgeScorers.Clear
    Caliper.SingleEdgeScorers.Add Scorer
    
    Dim Rectangle As New CogRectangleAffine
    Rectangle.SetCenterLengthsRotationSkew XRES / 2, XRES / 2, 200, 100, CogMisc.DegToRad(180), 0
    
    MaxX = 0#
    MinX = XRES
    
    For i = 0 To Tool.Results.Count - 1
        PointMarker.SetCenterRotationSize Tool.Results.Item(i).PositionX, Tool.Results.Item(i).PositionY, 0, 12
        Display.StaticGraphics.Add PointMarker
        Display.StaticGraphics.Add Tool.Results.Item(i).CreateResultGraphics(cogCaliperResultGraphicEdges)
        
        Rectangle.CenterX = Tool.Results.Item(i).PositionX + Region.SideYLength / 2
        Rectangle.CenterY = Tool.Results.Item(i).PositionY
        
        Dim Results As CogCaliperResults
        Set Results = Caliper.Execute(Image, Rectangle)
        
        If Results.Count > 0 Then
            Display.StaticGraphics.Add Results.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
        End If
        
        If MinX > Results.Item(0).PositionX Then
            MinX = Results.Item(0).PositionX
            MinY = Results.Item(0).PositionY
        End If
        
        If MaxX < Results.Item(0).PositionX Then
            MaxX = Results.Item(0).PositionX
            MaxY = Results.Item(0).PositionY
        End If
    Next i
    
    Dim Line1 As New CogLine
    Dim Line2 As New CogLine
    
    Line1.Color = cogColorRed
    Line1.SetXYRotation MinX, 0, CogMisc.DegToRad(90)
    
    Line2.Color = cogColorRed
    Line2.SetXYRotation MaxX, 0, CogMisc.DegToRad(90)
    
    Display.StaticGraphics.Add Line1
    Display.StaticGraphics.Add Line2
    
    Dim Distance As Double
    Dim CX As Double
    Dim CY As Double
    
    Distance = CogMath.DistancePointLine(MinX, MinY, Line2, Image, CX, CY) * Calib + Offset
    CogCaliperDiff = Distance
    
    Dim Segment As New CogLineSegment
    
    Segment.Color = cogColorOrange
    Segment.StartPointAdornment = cogLineSegmentAdornmentArrow
    Segment.EndPointAdornment = cogLineSegmentAdornmentArrow
    Segment.SetStartEnd MinX, MinY, CX, CY
    
    Display.StaticGraphics.Add Segment
    
    CogDisplayLabel Display, Segment.MidpointX, Segment.MidpointY, Format(Distance, "#0.00"), cogColorGreen, "Tahoma", 16
    
    Exit Function
ErrorHandle:
    CogCaliperDiff = -1#

End Function

Public Function CogBlobDiffEx(ByRef BlobTool As CogBlobTool, Optional ByRef Display As CogDisplay = Nothing, Optional ByVal Calib As Double = 1#, Optional ByVal Offset As Double = 0#) As Double
On Error Resume Next
    
    If BlobTool Is Nothing Then
        CogBlobDiffEx = 0#
        Exit Function
    End If
    
    BlobTool.Run
    
    If BlobTool.Results.Blobs.Count < 2 Then
        CogBlobDiffEx = 0#
        Exit Function
    End If
    
    Dim Line1 As New CogLine
    Dim Line2 As New CogLine
    
    Line1.Color = cogColorRed
    Line1.SetXYRotation BlobTool.Results.Blobs.Item(0).Measure(cogBlobMeasureBoundingBoxPixelAlignedNoExcludeMaxX), 0, CogMisc.DegToRad(90)
    
    Line2.Color = cogColorRed
    Line2.SetXYRotation BlobTool.Results.Blobs.Item(1).Measure(cogBlobMeasureBoundingBoxPixelAlignedNoExcludeMaxX), 0, CogMisc.DegToRad(90)
    
    If Not Display Is Nothing Then
        Display.StaticGraphics.Add Line1
        Display.StaticGraphics.Add Line2
    End If
    
    Dim Distance As Double
    Dim CX As Double
    Dim CY As Double
    
    Distance = CogMath.DistancePointLine(Line1.X, YRES / 2, Line2, BlobTool.InputImage, CX, CY) * Calib + Offset
    CogBlobDiffEx = Distance
    
    Dim Segment As New CogLineSegment
    
    Segment.Color = cogColorOrange
    Segment.StartPointAdornment = cogLineSegmentAdornmentArrow
    Segment.EndPointAdornment = cogLineSegmentAdornmentArrow
    Segment.SetStartEnd Line1.X, CY, Line2.X, CY
    
    If Display Is Nothing Then
        Exit Function
    End If
    
    Display.StaticGraphics.Add Segment
    CogDisplayLabel Display, Segment.MidpointX, Segment.MidpointY, Format(Distance, "#0.00"), cogColorGreen, "Tahoma", 16
    
    For i = 0 To BlobTool.Results.Blobs.Count - 1
        Dim X As Double
        Dim Y As Double
        Dim Area As Double

        Display.StaticGraphics.Add BlobTool.Results.Blobs.Item(i).GetBoundary
        X = BlobTool.Results.Blobs.Item(i).CenterOfMassX
        Y = BlobTool.Results.Blobs.Item(i).CenterOfMassY
        Area = BlobTool.Results.Blobs.Item(i).Area

        CogDisplayLabel Display, X, Y, Format(Area, "#0"), cogColorGreen, "Tahoma", 8
        
    Next i
        
  
End Function


Public Function CogBlobDiff(ByRef Image As CogImage8Grey, ByRef Region As CogRectangleAffine, Optional ByRef Display As CogDisplay = Nothing, Optional ByVal Calib As Double = 1#, Optional ByVal Offset As Double = 0#) As Double

    Dim BlobTool As New CogBlobTool
    Dim i As Long
    
    With BlobTool.RunParams.SegmentationParams
        .Mode = cogBlobSegmentationModeHardFixedThreshold
        .Polarity = cogBlobSegmentationPolarityDarkBlobs
        .HardFixedThreshold = 180
    End With
    
    BlobTool.RunParams.ConnectivityMinPixels = 10000
    
    Set BlobTool.InputImage = Image
    Set BlobTool.Region = Region
    
    BlobTool.Run
    
    If BlobTool.Results.Blobs.Count < 2 Then
        CogBlobDiff = 0#
        Exit Function
    End If
    
    Dim Line1 As New CogLine
    Dim Line2 As New CogLine
    
    Line1.Color = cogColorRed
    Line1.SetXYRotation BlobTool.Results.Blobs.Item(0).Measure(cogBlobMeasureBoundingBoxPixelAlignedNoExcludeMaxX), 0, CogMisc.DegToRad(90)
    
    Line2.Color = cogColorRed
    Line2.SetXYRotation BlobTool.Results.Blobs.Item(1).Measure(cogBlobMeasureBoundingBoxPixelAlignedNoExcludeMaxX), 0, CogMisc.DegToRad(90)
    
    Dim Distance As Double
    Dim CX As Double
    Dim CY As Double
    
    Distance = CogMath.DistancePointLine(Line1.X, YRES / 2, Line2, Image, CX, CY) * Calib + Offset
    CogBlobDiff = Distance
    
    Dim Segment As New CogLineSegment
    
    Segment.Color = cogColorOrange
    Segment.StartPointAdornment = cogLineSegmentAdornmentArrow
    Segment.EndPointAdornment = cogLineSegmentAdornmentArrow
    Segment.SetStartEnd Line1.X, CY, Line2.X, CY
    
    If Display Is Nothing Then
        Exit Function
    End If
    
    Display.StaticGraphics.Add Line1
    Display.StaticGraphics.Add Line2
    Display.StaticGraphics.Add Segment
    CogDisplayLabel Display, Segment.MidpointX, Segment.MidpointY, Format(Distance, "#0.00"), cogColorGreen, "Tahoma", 16
    
    For i = 0 To BlobTool.Results.Blobs.Count - 1
        Dim X As Double
        Dim Y As Double
        Dim Area As Double

        Display.StaticGraphics.Add BlobTool.Results.Blobs.Item(i).GetBoundary
        X = BlobTool.Results.Blobs.Item(i).CenterOfMassX
        Y = BlobTool.Results.Blobs.Item(i).CenterOfMassY
        Area = BlobTool.Results.Blobs.Item(i).Area

        CogDisplayLabel Display, X, Y, Format(Area, "#0"), cogColorGreen, "Tahoma", 8
        
    Next i
        
  
End Function

Public Function CogCaliperAngle(ByRef Tool1 As CogCaliperTool, ByRef Tool2 As CogCaliperTool, Optional ByRef Display As CogDisplay = Nothing, Optional ByVal Hori As Boolean = False, Optional ByRef Line As CogLine = Nothing, Optional ByVal Offset As Double = 0#) As Double
On Error GoTo ErrorHandle
    
    If Tool1 Is Nothing Or Tool2 Is Nothing Then
        CogCaliperAngle = 0#
        Exit Function
    End If
    
    Tool1.Run
    If Tool1.Results.Count <= 0 Then
        CogCaliperAngle = 0#
        Exit Function
    End If
    
    Tool2.Run
    If Tool2.Results.Count <= 0 Then
        CogCaliperAngle = 0#
        Exit Function
    End If
    
    Dim X1 As Double
    Dim Y1 As Double
    Dim X2 As Double
    Dim Y2 As Double
    
    X1 = Tool1.Results.Item(0).PositionX
    Y1 = Tool1.Results.Item(0).PositionY
    X2 = Tool2.Results.Item(0).PositionX
    Y2 = Tool2.Results.Item(0).PositionY
    
    Dim Angle As Double
    Angle = CogMisc.RadToDeg(CogMath.AnglePointPoint(X1, Y1, X2, Y2)) + Offset
    
    If Angle < 0 And Hori = False Then
        Angle = CogMisc.RadToDeg(CogMath.AnglePointPoint(X2, Y2, X1, Y1)) + Offset
    End If
    
    CogCaliperAngle = Angle
    
    Dim Line1 As New CogLine
    
    Line1.Color = cogColorOrange
    Line1.SetFromStartXYEndXY X1, Y1, X2, Y2
    
    If Not Line Is Nothing Then
        Set Line = Line1.Copy
    End If
    
    If Display Is Nothing Then
        Exit Function
    End If
    Display.StaticGraphics.Add Line1
    
    Dim Point1 As New CogPointMarker
    Dim Point2 As New CogPointMarker
    
    Point1.Color = cogColorRed
    Point1.SetCenterRotationSize X1, Y1, 0, 12
    
    Point2.Color = cogColorRed
    Point2.SetCenterRotationSize X2, Y2, 0, 12
    
    Display.StaticGraphics.Add Point1
    Display.StaticGraphics.Add Point2
    
    Dim Segment As New CogLineSegment
    Segment.SetStartEnd X1, Y1, X2, Y2
    
    CogDisplayLabel Display, Segment.MidpointX, Segment.MidpointY, Format(Angle, "#0.00"), cogColorGreen, "Tahoma", 16
    
    Exit Function
ErrorHandle:
    CogCaliperAngle = 0#
    
End Function

Public Function CogFindLineY(ByRef Tool As CogFindLineTool, PointX As Double, PointY As Double, Optional ByRef Image As CogImage8Grey = Nothing, Optional ByRef Region As CogRectangleAffine = Nothing, Optional ByRef Display As CogDisplay = Nothing, Optional Calib As Double = 1#, Optional Offset As Double = 0#, Optional ByRef Label As CogGraphicLabel = Nothing) As Double
    
    Dim Segment As CogLineSegment
    Dim i As Integer
    
    If Tool Is Nothing Then
        CogFindLineY = -1#
        Exit Function
    End If
    
    If Not Image Is Nothing Then
        Set Tool.InputImage = Image
    End If
    
    If Not Region Is Nothing Then
        Dim X As Double
        Dim Y As Double
        
        Set Segment = New CogLineSegment
        
        Segment.StartX = Region.CenterX - (Region.SideYLength / 2)
        Segment.StartY = Region.CenterY
        Segment.EndX = Region.CenterX + (Region.SideYLength / 2)
        Segment.EndY = Region.CenterY
        
        Set Tool.RunParams.ExpectedLineSegment = Segment
        Tool.RunParams.CaliperSearchLength = Region.SideXLength
        
        CogDisplayRectangle Display, Region, False, cogColorDarkGreen
        CogDisplaySegment Display, Segment, False, cogColorOrange
    End If
    
    Tool.Run
    
    If Tool.Results.Count <= 0 Then
        CogFindLineY = -1#
        Exit Function
    End If
    
    If Tool.Results.GetLineSegment Is Nothing Then
        CogFindLineY = -1#
        Exit Function
    End If
    
    Display.StaticGraphics.Add Tool.Results.GetLineSegment
    
    Dim Line1 As New CogLine
    Line1.SetXYRotation PointX, PointY, 0
    
    Dim Distance As Double
    Dim SX As Double
    Dim SY As Double
    Dim EX As Double
    Dim EY As Double
    
    Dim MinDistance As Double
    Dim MinIndex As Integer
    
    MinDistance = XRES
    
'    For i = 0 To Tool.Results.Count - 1
'        Dim CX As Double
'        Dim CY As Double
'        Dim Color As CogColorConstants
'        If Tool.Results.Item(i).Found = True Then
'            CX = Tool.Results.Item(i).x
'            CY = Tool.Results.Item(i).y
'            Color = IIf(Tool.Results.Item(i).Used = True, cogColorGreen, cogColorRed)
'            CogDisplayPoint Display, CX, CY, Color
'        End If
'
'        If Tool.Results.Item(i).Used = True Then
'            Distance = CogMath.DistancePointLine(CX, CY, Line1, Image)
'            If Distance < MinDistance Then
'                MinDistance = Distance
'                MinIndex = i
'            End If
'
'            CogDisplayLabel Display, CX, CY, Format(Distance, "#0.0"), cogColorOrange, "Tahoma", 8
'        End If
'    Next i
    
    'SX = Tool.Results.Item(MinIndex).x
    'SY = Tool.Results.Item(MinIndex).y
        
    Distance = CogMath.DistanceSegmentLine(Tool.Results.GetLineSegment, Line1, Image, SX, SY, EX, EY) * Calib + Offset
    CogFindLineY = Distance
    
    If g_bUseCaliperTool = True Then
        Dim Region2 As New CogRectangleAffine
        Region2.SetCenterLengthsRotationSkew SX, SY, 30, 10, CogMisc.DegToRad(-90), 0
        Region2.Color = cogColorDarkGreen
        
        Dim Caliper As New CogCaliper
        Caliper.ContrastThreshold = 5
        Caliper.FilterHalfSizeInPixels = 3
        Caliper.Edge0Polarity = cogCaliperPolarityDarkToLight
        Caliper.MaxResults = 1
        Caliper.SingleEdgeScorers.Clear
        Caliper.SingleEdgeScorers.Add g_CogCaliperScorerPosition
        
        Dim CaliperResults As CogCaliperResults
        Set CaliperResults = Caliper.Execute(Image, Region2)
        
        If CaliperResults Is Nothing Then
            CogFindLineY = -1#
            Exit Function
        End If
        
        Display.StaticGraphics.Add Region2
        
        SX = CaliperResults.Item(0).PositionX
        SY = CaliperResults.Item(0).PositionY
        
        CogDisplayPoint Display, SX, SY
        
        Distance = CogMath.DistancePointLine(SX, SY, Line1, Image, EX, EY) * Calib + Offset
        CogFindLineY = Distance
    Else
        
    End If
    
    Dim Segment2 As New CogLineSegment
    Segment2.SetStartEnd SX, SY, EX, EY
    Segment2.StartPointAdornment = cogLineSegmentAdornmentSolidArrow
    Segment2.EndPointAdornment = cogLineSegmentAdornmentSolidArrow
    Segment2.Color = cogColorOrange
    
    Display.StaticGraphics.Add Segment2
    
    If Label Is Nothing Then
        CogDisplayLabel Display, Segment2.MidpointX, Segment2.MidpointY, Format(Distance, "#0.00") & IIf(Calib <> 1, " ㎜", ""), cogColorGreen, "Tahoma", 16
    Else
        Label.X = Segment2.MidpointX
        Label.Y = Segment2.MidpointY
        Label.Text = Format(Distance, "#0.00") & IIf(Calib <> 1, " mm", "")
        Label.Color = cogColorGreen
        Label.Font.Bold = True
        Label.Font.Name = "Tahoma"
        Label.Font.size = 16
    End If
    
    If Tool Is Nothing Then
        CogFindLineY = -1#
        Exit Function
    End If
    
End Function
