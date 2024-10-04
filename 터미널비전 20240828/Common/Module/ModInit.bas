Attribute VB_Name = "ModInit"
Public Sub Grid_Init()
Dim i As Integer
Dim j As Integer
Dim stemp As String
Dim str_BlobName As String
    frmMain.MSFlexGrid1.WordWrap = True                    '한 Cell 에 두줄 쓸수 있게 됨
    frmMain.MSFlexGrid1.AllowUserResizing = flexResizeBoth 'Cell Size 를 마우스로 조절 할수 있음
    frmMain.MSFlexGrid1.SelectionMode = flexSelectionByRow
    
    For i = 1 To iToolCount / 2
        'stemp = stemp & "^" & sSpecName(i - 1) & vbCrLf & dSpecOriMin(i - 1) & "~" & dSpecOriMax(i - 1) & "         |"
        stemp = stemp & "^" & sSpecName(i - 1) & " " & dSpecOriMin(i - 1) & "~" & dSpecOriMax(i - 1) & "         |"
    Next i
    
    str_BlobName = ""
    For i = 1 To iBlobToolCount
        str_BlobName = str_BlobName & "^" & sBlobName(i - 1) & "|"
    Next i
    
    '검사결과 Display
    frmMain.MSFlexGrid1.Rows = 1
    frmMain.MSFlexGrid1.Cols = 4 + (iToolCount / 2)
    If iProName = 0 Then
        frmMain.MSFlexGrid1.FormatString = "^Number    |" & "^검사시간                |" & "^ID_Code 1                   |" & "^ID_Code 2                   |" _
                                        & "^ID_Code 3                   |" & "^ID_Code 4                   |" & "^판정  |" & stemp
                                       '& "^" & sSpecName(8) & "           |" & "^" & sSpecName(9) & "           " vbcrlf
    Else
        frmMain.MSFlexGrid1.FormatString = "^Number    |" & "^검사시간                |" & "^ID_Code                    |" & "^판정  |" & stemp & str_BlobName
    End If
    If iProName = 0 Then
        For j = 0 To (iToolCount / 2) + 6
            frmMain.MSFlexGrid1.ColWidth(j) = 1500
        Next j
    Else
        For j = 0 To (iToolCount / 2) + 4
            frmMain.MSFlexGrid1.ColWidth(j) = 1500
        Next j
    End If
        frmMain.MSFlexGrid1.RowHeight(0) = 400
        frmMain.MSFlexGrid1.ColWidth(0) = 800
        frmMain.MSFlexGrid1.ColWidth(1) = 1400
        frmMain.MSFlexGrid1.ColWidth(2) = 4000
        If iProName = 0 Then
            frmMain.MSFlexGrid1.ColWidth(3) = 1800
            frmMain.MSFlexGrid1.ColWidth(4) = 1800
            frmMain.MSFlexGrid1.ColWidth(5) = 1800
            frmMain.MSFlexGrid1.ColWidth(6) = 400
        Else
            frmMain.MSFlexGrid1.ColWidth(3) = 400
        End If
End Sub

Public Sub Tool_Init()
On Error GoTo err:
Dim i As Integer
Dim j As Integer

'    Set m_favFlip = New FavFlipTool
'
'    For i = 0 To 3
''        Set favImageFileT(i) = New FvImageFileTool
'        favImageFileT(i).FileType = 1 'File Type 이 1 이면 Bmp , 2 면 JPG , 0 은 Raw
''        For j = 0 To 29
''            Set favEdgeT(i, j) = New FavCaliperTool
''            'favEdgeT(i, j).Direction = 0
''            'favEdgeT(i, j).FilterWidth = 2
''
''            favEdgeT(i, j).Mode = FavCaliperModeFirstEdge
''            'favEdgeT(i, j).Mode = FavCaliperModeMaxContrast
''            favEdgeT(i, j).Polarity = FavCaliperPolarityDarkToLight
''            favEdgeT(i, j).Threshold = 10
''            favEdgeT(i, j).SetRegion 200, 100, 200, 100, 0
''            dEdgeCenterX(i, j) = 200
''            dEdgeCenterY(i, j) = 100
''            lEdgeSideX(i, j) = 200
''            lEdgeSideY(i, j) = 100
''            dEdgeRotation(i, j) = 0
''            'favEdgeT(i, j).InputImage = fvImageBuf(i)
'''            favEdgeT(i, j).ImageWidth = XRES
'''            favEdgeT(i, j).ImageHeight = YRES
''        Next j
'
''        For j = 0 To 3
''            Set favFixEdgeT(i, j) = New FavCaliperTool
''            'favFixEdgeT(i, j).Direction = 0
''            'favFixEdgeT(i, j).FilterWidth = 2
''            favFixEdgeT(i, j).Mode = FavCaliperModeFirstEdge
''            favFixEdgeT(i, j).Polarity = FavCaliperPolarityDarkToLight
''            favFixEdgeT(i, j).Threshold = 10
''            favFixEdgeT(i, j).SetRegion 200, 100, 200, 100, 0
''            dFixEdgeCenterX(i, j) = 200
''            dFixEdgeCenterY(i, j) = 100
''            lFixEdgeSideX(i, j) = 200
''            lFixEdgeSideY(i, j) = 100
''            dFixEdgeRotation(i, j) = 0
''            'favFixEdgeT(i, j).InputImage = fvImageBuf(i)
'''            favFixEdgeT(i, j).ImageWidth = XRES
'''            favFixEdgeT(i, j).ImageHeight = YRES
''        Next j
''
''        For j = 0 To 3
''            Set favCalEdgeT(i, j) = New FavCaliperTool
''            'favCalEdgeT(i, j).Direction = 0
''            'favCalEdgeT(i, j).FilterWidth = 2
''            favCalEdgeT(i, j).Mode = FavCaliperModeFirstEdge
''            favCalEdgeT(i, j).Polarity = FavCaliperPolarityDarkToLight
''            favCalEdgeT(i, j).Threshold = 10
''            favCalEdgeT(i, j).SetRegion 200, 100, 200, 100, 0
''            dCalEdgeCenterX(i, j) = 200
''            dCalEdgeCenterY(i, j) = 100
''            lCalEdgeSideX(i, j) = 200
''            lCalEdgeSideY(i, j) = 100
''            dCalEdgeRotation(i, j) = 0
''            'favCalEdgeT(i, j).InputImage = fvImageBuf(i)
'''            favCalEdgeT(i, j).ImageWidth = XRES
'''            favCalEdgeT(i, j).ImageHeight = YRES
''        Next j
''    Next i
''
''    For i = 0 To 3
''        For j = 0 To 29
''            Set favBlobT(i, j) = New FavBlobTool
''            favBlobT(i, j).Polarity = FavBlobPolarityDark
''            favBlobT(i, j).Binalization = FavHardFixedThreshold          ' 단순 이진화
''            favBlobT(i, j).Threshold = 50
''            favBlobT(i, j).MinWidth = 3                                  '
''            favBlobT(i, j).MaxWidth = XRES
''            favBlobT(i, j).MinHeight = 3
''            favBlobT(i, j).MaxHeight = YRES
''            favBlobT(i, j).MinArea = 5
''            favBlobT(i, j).ClearMorphology
''            favBlobT(i, j).SetRegion 200, 200, 50, 50
''            favBlobT(i, j).UseBoundary = True            'true 로 하면 툴 검사시 세번째(툴넘버상관없이) 검사때 런타임 오류
''            favBlobT(i, j).ClearMorphology
''
''            lBlobCenterX(i, j) = 200
''            lBlobCenterY(i, j) = 200
''            lBlobSideX(i, j) = 50
''            lBlobSideY(i, j) = 50
''
''        Next j
'    Next i
Exit Sub

err:
    MsgBox "툴초기화에 실패하였습니다.", vbCritical, "초기화 오류"
    
End Sub
