Attribute VB_Name = "ModFileSave"

Public Sub Modeldata_Filesave(mdlname As String)
On Error GoTo err:
Dim i As Integer
Dim j As Integer

Dim CurRecord As Long
Dim RecordCount As Long
Dim FileNumber As Integer
Dim RecordLen As Long

    CurRecord = 0
    RecordCount = 0
    FileNumber = FreeFile
    RecordLen = Len(Modelinfo)

    Open App.Path & "\Model\" & mdlname & "\" & mdlname & ".dat" For Random As FileNumber Len = RecordLen
    CurRecord = 1
    RecordCount = LOF(FileNumber) / RecordLen
    If RecordCount = 0 Then
        RecordCount = 1
    End If
        Modelinfo.ModelName = mdlname
        For i = 0 To 3
            Modelinfo.Exposure(i) = 25 'frmMain.uEyeCam1(i).SetExposureTime(25)
        Next i
        For i = 0 To 3
            For j = 0 To 29 'Edge Region
                Modelinfo.favEdgeTCenterX(i, j) = dEdgeCenterX(i, j)
                Modelinfo.favEdgeTCenterY(i, j) = dEdgeCenterY(i, j)
                Modelinfo.favEdgeTSideX(i, j) = lEdgeSideX(i, j)
                Modelinfo.favEdgeTSideY(i, j) = lEdgeSideY(i, j)
                Modelinfo.favEdgeTRotation(i, j) = dEdgeRotation(i, j)
            Next j
            For j = 0 To 3
                Modelinfo.favFixEdgeTCenterX(i, j) = dFixEdgeCenterX(i, j)
                Modelinfo.favFixEdgeTCenterY(i, j) = dFixEdgeCenterY(i, j)
                Modelinfo.favFixEdgeTSideX(i, j) = lFixEdgeSideX(i, j)
                Modelinfo.favFixEdgeTSideY(i, j) = lFixEdgeSideY(i, j)
                Modelinfo.favFixEdgeTRotation(i, j) = dFixEdgeRotation(i, j)
            Next j
            For j = 0 To 3
                Modelinfo.favCalEdgeTCenterX(i, j) = dCalEdgeCenterX(i, j)
                Modelinfo.favCalEdgeTCenterY(i, j) = dCalEdgeCenterY(i, j)
                Modelinfo.favCalEdgeTSideX(i, j) = lCalEdgeSideX(i, j)
                Modelinfo.favCalEdgeTSideY(i, j) = lCalEdgeSideY(i, j)
                Modelinfo.favCalEdgeTRotation(i, j) = dCalEdgeRotation(i, j)
            Next j
            For j = 0 To 29
                Modelinfo.favBlobTCenterX(i, j) = lBlobCenterX(i, j)
                Modelinfo.favBlobTCenterY(i, j) = lBlobCenterY(i, j)
                Modelinfo.favBlobTWidth(i, j) = lBlobSideX(i, j)
                Modelinfo.favBlobTHeight(i, j) = lBlobSideY(i, j)
            Next j
        Next i

    Put #FileNumber, CurRecord, Modelinfo
    Close #FileNumber

Exit Sub

err:
    
    
End Sub

Public Function ModelData_FileLoad(fName As String) As Boolean
On Error GoTo LPerr
Dim CurRecord As Long
Dim RecordCount As Long
Dim FileNumber As Integer
Dim RecordLen As Long
Dim bRet As Boolean
Dim i As Integer
Dim j As Integer
Dim FilePath As String


    CurRecord = 0
    RecordCount = 0
    FileNumber = FreeFile
    RecordLen = Len(Modelinfo)
    
    FilePath = App.Path & "\Model\" & fName & "\" & fName & ".dat"
    
    Open App.Path & "\Model\" & fName & "\" & fName & ".dat" For Random As FileNumber Len = RecordLen
        CurRecord = 1
        RecordCount = LOF(FileNumber) / RecordLen
        If RecordCount = 0 Then
            RecordCount = 1
        End If
        Get #FileNumber, CurRecord, Modelinfo
    Close #FileNumber
    
    For i = 0 To 3
        'Exposure
        'bRet = frmMain.uEyeCam1(i).SetExposureTime(Modelinfo.Exposure(i))
        'Caliper
        For j = 0 To 29
            dEdgeCenterX(i, j) = Modelinfo.favEdgeTCenterX(i, j)
            dEdgeCenterY(i, j) = Modelinfo.favEdgeTCenterY(i, j)
            lEdgeSideX(i, j) = Modelinfo.favEdgeTSideX(i, j)
            lEdgeSideY(i, j) = Modelinfo.favEdgeTSideY(i, j)
            dEdgeRotation(i, j) = Modelinfo.favEdgeTRotation(i, j)
        Next j
        For j = 0 To 3
            dFixEdgeCenterX(i, j) = Modelinfo.favFixEdgeTCenterX(i, j)
            dFixEdgeCenterY(i, j) = Modelinfo.favFixEdgeTCenterY(i, j)
            lFixEdgeSideX(i, j) = Modelinfo.favFixEdgeTSideX(i, j)
            lFixEdgeSideY(i, j) = Modelinfo.favFixEdgeTSideY(i, j)
            dFixEdgeRotation(i, j) = Modelinfo.favFixEdgeTRotation(i, j)
        Next j
        For j = 0 To 3
            dCalEdgeCenterX(i, j) = Modelinfo.favCalEdgeTCenterX(i, j)
            dCalEdgeCenterY(i, j) = Modelinfo.favCalEdgeTCenterY(i, j)
            lCalEdgeSideX(i, j) = Modelinfo.favCalEdgeTSideX(i, j)
            lCalEdgeSideY(i, j) = Modelinfo.favCalEdgeTSideY(i, j)
            dCalEdgeRotation(i, j) = Modelinfo.favCalEdgeTRotation(i, j)
        Next j
        'Blob
        For j = 0 To 29
            lBlobCenterX(i, j) = Modelinfo.favBlobTCenterX(i, j)
            lBlobCenterY(i, j) = Modelinfo.favBlobTCenterY(i, j)
            lBlobSideX(i, j) = Modelinfo.favBlobTWidth(i, j)
            lBlobSideY(i, j) = Modelinfo.favBlobTHeight(i, j)
        Next j
    Next i
    
    Call LOGWrite("모델정보 로딩에 성공하였습니다.")
    ModelData_FileLoad = True
Exit Function

LPerr:
    ModelData_FileLoad = False
    Close #FileNumber
    Call LOGWrite("모델정보 로딩에 실패 하였습니다.")
    
End Function
Public Sub ProgramSelect_Load()
On Error GoTo err
Dim i As Integer
Dim fp As Integer
Dim temp(0 To 9) As String

    fp = FreeFile
    Open App.Path & "\Info" & "\PSelect.fav" For Input As fp
        Line Input #fp, temp(0)
        iProName = temp(0)
    Close fp
    Select Case iProName
        Case 0
            iCamNumber = 0
        Case 1
            iCamNumber = 3
        Case 2
            iCamNumber = 3
        Case Else
            GoTo err:
    End Select
    
Exit Sub
err:
    Close fp
    iProName = 0
    iCamNumber = 0
End Sub
Public Sub ProgramSelect_Save()
On Error GoTo err
Dim fp As Integer
Dim i As Integer
Dim j As Integer
fp = FreeFile
    
    Open App.Path & "\Info" & "\PSelect.fav" For Output As fp
        Print #fp, iProName
    Close fp
Exit Sub
    
err:
    Close fp
    
End Sub

Public Sub BlobName_Save(fName As String)
On Error GoTo err
Dim fp As Integer
Dim i As Integer
Dim j As Integer
fp = FreeFile
    
    Open App.Path & "\Model\" & fName & "\BlobName" & ".fav" For Output As fp
        For i = 0 To iBlobToolCount - 1
            Print #fp, sBlobName(i)
        Next i
    Close fp
Exit Sub
    
err:
    Close fp
    
End Sub
Public Sub BlobName_Load(fName As String)
On Error GoTo err
Dim fp As Integer
Dim i As Integer
Dim j As Integer
Dim temp As String
fp = FreeFile
    
    Open App.Path & "\Model\" & fName & "\BlobName" & ".fav" For Input As fp
        For i = 0 To 29
            Line Input #fp, temp
            sBlobName(i) = temp
        Next i
    Close fp
Exit Sub
    
err:
    Close fp
    
End Sub
Public Sub SocketSET_Load()
On Error GoTo err
Dim i As Integer
Dim fp As Integer
Dim temp(0 To 9) As String

    fp = FreeFile
    Open App.Path & "\Info" & "\SocketSET.fav" For Input As fp
        For i = 0 To 3
            Line Input #fp, temp(i)
        Next i
        sPLCIP = temp(0)
        sPLCPort = temp(1)
        sMESIP = temp(2)
        sMESPort = temp(3)
    Close fp

    'frmMain.txtMESIP.Text = sMESIP
    'frmMain.txtMESPort.Text = sMESPort
    
Exit Sub
err:
    Close fp
    MsgBox "통신 설정을 확인 바랍니다. ", vbCritical, "통신설정 불러오기 오류"
    
End Sub
Public Sub SocketSET_Save()
On Error GoTo err
Dim fp As Integer
Dim i As Integer
Dim j As Integer

    fp = FreeFile
'    sPLCIP = frmMain.txtPLCIP.Text
'    sPLCPort = frmMain.txtPLCPort.Text
    'sMESIP = frmMain.txtMESIP.Text
    'sMESPort = frmMain.txtMESPort.Text
    Open App.Path & "\Info" & "\SocketSET.fav" For Output As fp
        Print #fp, sPLCIP
        Print #fp, sPLCPort
        Print #fp, sMESIP
        Print #fp, sMESPort
    Close fp
    
    Dim FileName As String
    Dim Section As String
    
    FileName = App.Path & "\" & App.title & ".ini"
    
    Call WriteINI(FileName, "MES", "IP", sPLCIP)
    Call WriteINI(FileName, "MES", "PORT", sPLCPort)
    
Exit Sub
    
err:
    Close fp
    
End Sub
Public Function LastModelWrite()
On Error Resume Next
    
    Dim FileName As String
    FileName = App.Path & "\" & App.title & ".ini"
    
    Call WriteINI(FileName, "SYSTEM", "LAST_MODEL", sModelName)
    Call WriteINI(FileName, "SYSTEM", "MODEL_NUMBER", CStr(g_ModelNumber))
    Call WriteINI(FileName, "SYSTEM", "MODEL_CHANGED_DATE", g_ModelChangedDate)

End Function

Public Function LastModelRead() As Boolean
On Error Resume Next
    
    Dim FileName As String
    FileName = App.Path & "\" & App.title & ".ini"
    
    sModelName = ReadINI(FileName, "SYSTEM", "LAST_MODEL")
    g_ModelNumber = CInt(ReadINI(FileName, "SYSTEM", "MODEL_NUMBER"))
    g_ModelChangedDate = ReadINI(FileName, "SYSTEM", "MODEL_CHANGED_DATE")
    
End Function

Public Function ModelList_SAVE() As Boolean
On Error Resume Next

    ' 쓰기
    Dim FileName As String
    Dim Section As String
    
    Dim i As Integer
    
    FileName = App.Path & "\Model\ModelList.ini"
    Section = "Rooms"
    
    For i = 1 To 100
        Call WriteINI(FileName, Section, CStr(i), Trim(sModelRoom(i)))
    Next i

End Function

Public Function ModelList_LOAD() As Boolean
On Error Resume Next

    ' 쓰기
    Dim FileName As String
    Dim Section As String
    
    Dim i As Integer
    
    FileName = App.Path & "\Model\ModelList.ini"
    Section = "Rooms"
    
    For i = 1 To 100
        sModelRoom(i) = ReadINI(FileName, Section, CStr(i))
    Next i
    
End Function
Public Sub Calibration_Load(mdname As String)
On Error Resume Next

    Dim FileName As String
    Dim Section As String
    Dim Key As String
    
    Dim i As Integer
    
    FileName = App.Path & "\model\" & mdname & "\model.ini"
    
    For i = 0 To kMaxCamera - 1
        Section = "Calibration_Camera" & CStr(i + 1)
        dCaliMM(i) = CDbl(ReadINI(FileName, Section, "mm"))
        dCaliPX(i) = CDbl(ReadINI(FileName, Section, "mm/Pixel"))
        dCaliPXY(i) = CDbl(ReadINI(FileName, Section, "mm/Pixel_Y"))
    Next i
    
End Sub
Public Sub Calibration_Save(mdname As String)
On Error Resume Next

    Dim FileName As String
    Dim Section As String
    Dim Key As String
    
    Dim i As Integer
    
    FileName = App.Path & "\model\" & mdname & "\model.ini"
    
    For i = 0 To kMaxCamera - 1
        Section = "Calibration_Camera" & CStr(i + 1)
        WriteINI FileName, Section, "mm", CStr(dCaliMM(i))
        WriteINI FileName, Section, "mm/Pixel", CStr(dCaliPX(i))
        WriteINI FileName, Section, "mm/Pixel_Y", CStr(dCaliPXY(i))
    Next i
    
End Sub
Public Sub Calibration_Loady(mdname As String)
On Error GoTo err
Dim i As Integer
Dim fp As Integer
Dim temp(0 To 9) As String

    fp = FreeFile
    Open App.Path & "\model\" & mdname & "\Calibrationy.fav" For Input As fp
        For i = 0 To 7
            Line Input #fp, temp(i)
        Next i
        For i = 0 To 3
            dCaliMM(i) = temp(0 + i * 2)
            dCaliPXY(i) = temp(1 + i * 2)
        Next i
    Close fp
    
Exit Sub
err:
    Close fp
    'MsgBox "켈리브레이션 값이 설정되어 있지 않습니다. ", vbCritical, "CALIBRATION LOAD FAIL"
    
End Sub
Public Sub Calibration_Savey(mdname As String)
On Error GoTo err
Dim fp As Integer
Dim i As Integer
Dim j As Integer

    fp = FreeFile
    
'    dCaliMM(iCamNumberS) = Format(CDbl(frmSetting.txtCalmm.Text), "#00.0000")
'    dCaliPX(iCamNumberS) = Format(CDbl(frmSetting.txtCalmmPy.Text), "#00.0000")
    
    Open App.Path & "\model\" & mdname & "\Calibration.favy" For Output As fp
        For i = 0 To 3
            Print #fp, dCaliMM(i)
            Print #fp, dCaliPXY(i)
        Next i
'        Print #fp, dCaliMM(iCamNumberS)
'        Print #fp, dCaliPX(iCamNumberS)

    Close fp
Exit Sub
    
err:
    Close fp
    
End Sub
Public Sub FixPoint_Load(mdname As String)
On Error GoTo err
Dim i As Integer
Dim fp As Integer
Dim temp(0 To 50) As String

    fp = FreeFile
    Open App.Path & "\model\" & mdname & "\Fixpoint.fav" For Input As fp
        For i = 0 To 15
            Line Input #fp, temp(i)
        Next i
        For i = 0 To 3
            Fix_PointX(i) = temp(0 + (i * 4))
            Fix_PointY(i) = temp(1 + (i * 4))
            Fix_PointAngle(i) = temp(2 + (i * 4))
            Fix_UseMode(i) = temp(3 + (i * 4))
        Next i
    Close fp
    
Exit Sub
err:
    Close fp
    MsgBox "기준점 값이 설정되어 있지 않습니다. ", vbCritical, "기준점 값 불러오기 실패"
    
End Sub
Public Sub FixPoint_Save(mdname As String)
On Error GoTo err
Dim fp As Integer
Dim i As Integer
Dim j As Integer

    fp = FreeFile
    
    Open App.Path & "\model\" & mdname & "\Fixpoint.fav" For Output As fp
        
        For i = 0 To 3
            Print #fp, Fix_PointX(i)
            Print #fp, Fix_PointY(i)
            Print #fp, Fix_PointAngle(i)
            Print #fp, Fix_UseMode(i)
        Next i
    Close fp
Exit Sub
    
err:
    Close fp
    
End Sub
Public Sub SpecName_Load(mdname As String)
On Error Resume Next

    Dim FileName As String
    Dim Section As String
    Dim i As Integer

    FileName = App.Path & "\Model\" & mdname & "\model.ini"
    Section = "SpecName"
    
    For i = 0 To 9
        sSpecName(i) = ReadINI(FileName, Section, "Name" & CStr(i + 1))
    Next i
    
End Sub
Public Sub SpecName_Save(mdname As String)
On Error Resume Next
    
    Dim FileName As String
    Dim Section As String
    Dim i As Integer

    FileName = App.Path & "\Model\" & mdname & "\model.ini"
    Section = "SpecName"
        
    For i = 0 To 9
        sSpecName(i) = frmMain.txtSpecName(i).Text
        Call WriteINI(FileName, Section, "Name" & CStr(i + 1), CStr(sSpecName(i)))
    Next i
    
End Sub
Public Sub SpecAllValue_Load(mdname As String)
On Error Resume Next
    
    Dim FileName As String
    Dim Section As String
    
    Dim i, j As Integer

    FileName = App.Path & "\model\" & mdname & "\model.ini"
    
    For i = 0 To 10
        Section = "SpecValue" & CStr(i + 1)
        dSpecOri(i) = CDbl(ReadINI(FileName, Section, "Origin"))
        dSpecMin(i) = CDbl(ReadINI(FileName, Section, "Min"))
        dSpecMax(i) = CDbl(ReadINI(FileName, Section, "Max"))
        For j = 0 To kMaxCamera - 1
            dSpecOffset(i + (j * 10)) = CDbl(ReadINI(FileName, Section, "Offset" & CStr(j + 1)))
        Next j
        dSpecOriMin(i) = CDbl(ReadINI(FileName, Section, "OriginMin"))
        dSpecOriMax(i) = CDbl(ReadINI(FileName, Section, "OriginMax"))
        bSpecPass(i) = CBool(ReadINI(FileName, Section, "Pass"))
        
        frmMain.chkJudgement(i).Value = IIf(bSpecPass(i) = True, 0, 1)
    Next i
    
    
End Sub

Public Sub SpecAllValue_Save(mdname As String)
On Error Resume Next
    
    Dim FileName As String
    Dim Section As String
    
    Dim i, j As Integer

    FileName = App.Path & "\model\" & mdname & "\model.ini"
    
    For i = 0 To 10
        Section = "SpecValue" & CStr(i + 1)
        Call WriteINI(FileName, Section, "Origin", CStr(dSpecOri(i)))
        Call WriteINI(FileName, Section, "Min", CStr(dSpecMin(i)))
        Call WriteINI(FileName, Section, "Max", CStr(dSpecMax(i)))
        Call WriteINI(FileName, Section, "OriginMin", CStr(dSpecOriMin(i)))
        Call WriteINI(FileName, Section, "OriginMax", CStr(dSpecOriMax(i)))
        Call WriteINI(FileName, Section, "Pass", CStr(bSpecPass(i)))
        
        For j = 0 To kMaxCamera - 1
            Call WriteINI(FileName, Section, "Offset" & CStr(j + 1), CStr(dSpecOffset(i + (j * 10))))
        Next j
    Next i
    
End Sub
Public Sub FunctionValue_Load(mdname As String)
On Error Resume Next

    Dim FileName As String
    Dim Section As String
    
    FileName = App.Path & "\model\" & mdname & "\model.ini"
    
    Section = "Option"
    
    bCamPass = IIf(ReadINI(FileName, Section, "CamPass") = True, 1, 0)
    bOKimageSave = 0
    bNGimageSave = 0
    bWriteDataSave = IIf(ReadINI(FileName, Section, "WriteDataSave") = True, 1, 0)
    iImageFileMode = CInt(ReadINI(FileName, Section, "ImageFileMode"))

End Sub

Public Sub FunctionValue_Save(mdname As String)
On Error Resume Next

    Dim FileName As String
    Dim Section As String
    
    FileName = App.Path & "\model\" & mdname & "\model.ini"
    
    Section = "Option"

    Call WriteINI(FileName, Section, "CamPass", CStr(bCamPass))
    Call WriteINI(FileName, Section, "OKImageSave", CStr(bOKimageSave))
    Call WriteINI(FileName, Section, "NGImageSave", CStr(bNGimageSave))
    Call WriteINI(FileName, Section, "WriteDataSave", CStr(bWriteDataSave))
    Call WriteINI(FileName, Section, "ImageFileMode", CStr(iImageFileMode))
    
End Sub

Public Sub DJ_CountSave()

On Error GoTo err

Dim ff As Integer
Dim i As Integer
    
    ff = FreeFile
    
    lOKCount = CDbl(frmMain.lblCountOK.Caption)
    lNGCount = CDbl(frmMain.lblCountNG.Caption)
    lToTalCount = CDbl(frmMain.lblCountTotal.Caption)
    
    If lToTalCount = 999999 Then
        lOKCount = 0
        lNGCount = 0
        lToTalCount = 0
        lInspectionNum = 0
        'Call FrontTape_InitGrid
    End If
    
    Open App.Path & "\Count" & ".fav" For Output As ff
    
        Print #ff, lOKCount
        Print #ff, lNGCount
        Print #ff, lToTalCount

    Close ff

    Exit Sub
err:
    Close ff

End Sub
Public Sub DJ_CountLoad()

'On Error GoTo err
On Error Resume Next

Dim i As Integer
Dim ff As Integer
Dim temp(0 To 30) As String

    
    ff = FreeFile
    
    If Dir(App.Path & "\Count" & ".fav", vbNormal) = "" Then
        MsgBox "카운트가 저장 되어 있지 않습니다..!!", vbCritical, "COUNT LOAD FAIL"
        Exit Sub
        
    End If
    
    Open App.Path & "\Count" & ".fav" For Input As ff
        For i = 0 To 2
            Line Input #ff, temp(i)
        Next i
        
        lOKCount = temp(0)
        lNGCount = temp(1)
        lToTalCount = temp(2)

    Close ff
        
    frmMain.lblCountOK.Caption = lOKCount
    frmMain.lblCountNG.Caption = lNGCount
    frmMain.lblCountTotal.Caption = lToTalCount
        
    Exit Sub
err:
    Close ff

    MsgBox "저장된 카운트가 없습니다.", vbCritical, "COUNT LOAD FAIL"

End Sub
Public Sub DJ_MainLoginSave()
On Error GoTo err
Dim ff As Integer
Dim i As Integer
    ff = FreeFile
    'sNowPassWord = frmMainLogin.txtPassword.Text
    Open "C:\Windows\system32" & "\VisionPW" & ".fav" For Output As ff
        Print #ff, sNowPassWord
    Close ff
    Exit Sub
err:
    Close ff
End Sub
Public Sub DJ_MainLoginLoad()
On Error Resume Next
Dim i As Integer
Dim ff As Integer
Dim temp As String
    ff = FreeFile
    Open "C:\Windows\system32" & "\VisionPW" & ".fav" For Input As ff
        Line Input #ff, temp
        sNowPassWord = temp
    Close ff
    Exit Sub
err:
    Close ff
End Sub
Public Sub DJ_ToolCountSave(mdname As String)
On Error GoTo err
Dim ff As Integer
Dim i As Integer
    ff = FreeFile
    Open App.Path & "\Model\" & mdname & "\ToolCount" & ".fav" For Output As ff
        Print #ff, iToolCount
        Print #ff, iBlobToolCount
    Close ff
    Exit Sub
err:
    Close ff
End Sub
Public Sub DJ_ToolCountLoad(mdname As String)
On Error Resume Next
    
    Dim FileName As String
    
    FileName = App.Path & "\Model\" & mdname & "\model" & ".ini"
Dim i As Integer
Dim ff As Integer
Dim temp As String
    ff = FreeFile
    iBlobToolCount = 0
    Open App.Path & "\Model\" & mdname & "\ToolCount" & ".fav" For Input As ff
        Line Input #ff, temp
        iToolCount = temp
        Line Input #ff, temp
        iBlobToolCount = temp
    Close ff
    Exit Sub
err:
    Close ff
End Sub

