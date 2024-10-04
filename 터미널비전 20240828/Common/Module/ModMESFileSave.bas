Attribute VB_Name = "commonMESFileSave"
Public Sub DJ_MESFunctionSave()
On Error GoTo err
Dim ff As Integer
Dim i As Integer
    ff = FreeFile
    'sNowPassWord = frmMainLogin.txtPassword.Text
    Open App.Path & "\Recipe" & "\MESFunction" & ".fav" For Output As ff
        Print #ff, sMESEquipCode
        Print #ff, sMESEquipName
        Print #ff, sMESLineNum
        Print #ff, sMESProgressCode
        Print #ff, sMESProcess
        Print #ff, sMESFileSavePath
        Print #ff, sMESFileSendPath
        Print #ff, sMESLogSavePath
        Print #ff, sMesPCIP
        Print #ff, sMesPCID
        Print #ff, sMesPCPW
        For i = 1 To 11
            Print #ff, sParamName_SV(iNowRecipeID, i)
        Next i
        Print #ff, sParamName_SVsj
        For i = 1 To 11
            Print #ff, sParamName_PV(iNowRecipeID, i)
        Next i
        For i = 1 To 11
            Print #ff, sParamName_NG(iNowRecipeID, i)
        Next i
    Close ff
    Exit Sub
err:
    Close ff
End Sub
Public Sub DJ_MESFunctionLoad()
On Error Resume Next
Dim i As Integer
Dim ff As Integer
Dim temp(0 To 50) As String
    ff = FreeFile
    Open App.Path & "\Recipe" & "\MESFunction" & ".fav" For Input As ff  '
        For i = 0 To 50
            Line Input #ff, temp(i)
        Next i
        sMESEquipCode = temp(0)
        sMESEquipName = temp(1)
        sMESLineNum = temp(2)
        sMESProgressCode = temp(3)
        sMESProcess = temp(4)
        sMESFileSavePath = temp(5)
        sMESFileSendPath = temp(6)
        sMESLogSavePath = temp(7)
        sMesPCIP = temp(8)
        sMesPCID = temp(9)
        sMesPCPW = temp(10)
        For i = 1 To 11
            sParamName_SV(iNowRecipeID, i) = temp(i + 10)
        Next i
        sParamName_SVsj = temp(12)
        For i = 1 To 11
            sParamName_PV(iNowRecipeID, i) = temp(i + 22)
        Next i
        For i = 1 To 11
            sParamName_NG(iNowRecipeID, i) = temp(i + 33)
        Next i
    Close ff
    Exit Sub
err:
    Close ff
End Sub
Public Sub DJ_MESLastLoginSave()
On Error GoTo err
Dim ff As Integer
Dim i As Integer
    ff = FreeFile
    Open App.Path & "\Recipe" & "\MESLogin" & ".fav" For Output As ff
        Print #ff, sMesUserID
    Close ff
    Exit Sub
err:
    Close ff
End Sub
Public Sub DJ_MESLastLoginLoad()
On Error Resume Next
Dim i As Integer
Dim ff As Integer
Dim temp(0 To 29) As String
    ff = FreeFile
    Open App.Path & "\Recipe" & "\MESLogin" & ".fav" For Input As ff  '
        Line Input #ff, temp(0)
        sMesUserID = temp(0)
    Close ff
    Exit Sub
err:
    frmMESLogin.txtMesID.Text = ""
    Close ff
End Sub
Public Sub DJ_MESRecipeSave(Index As Integer)
On Error GoTo err
Dim ff As Integer
Dim i As Integer
    ff = FreeFile
    Open App.Path & "\Recipe" & "\" & "RECIPE" & Index & ".fav" For Output As ff
        Print #ff, sRecipeID(Index)
        Print #ff, sRecipeComment(Index)
        Print #ff, sParamCount(Index)
        For i = 1 To CInt(sParamCount(Index))
            'Print #ff, sParamName_SV(Index, i)
            Print #ff, sParamValue(Index, i)
            Print #ff, sParamMinValue(Index, i)
            Print #ff, sParamMaxValue(Index, i)
        Next i
    Close ff
    Exit Sub
err:
    Close ff
End Sub
Public Sub DJ_MESRecipeLoad(Index As Integer)
On Error GoTo err:
Dim i As Integer
Dim ff As Integer
Dim temp(0 To 99) As String
    ff = FreeFile
    Open App.Path & "\Recipe" & "\" & "RECIPE" & Index & ".fav" For Input As ff '
'        For i = 0 To 49            '49 는 의미없음
'            Line Input #ff, temp(i)
'        Next i
        Do
            Line Input #ff, temp(i)
            i = i + 1
        Loop Until EOF(ff)
        sRecipeID(Index) = temp(0)
        sRecipeComment(Index) = temp(1)
        sParamCount(Index) = temp(2)
        If sParamCount(Index) = "" Then
            iParamCount(Index) = 0
        Else
            iParamCount(Index) = CInt(sParamCount(Index))
        End If
        For i = 0 To iParamCount(Index) - 1
            'sParamName_SV(Index, i + 1) = temp((i * 4) + 3)
            sParamValue(Index, i + 1) = temp((i * 4) + 4)
            sParamMinValue(Index, i + 1) = temp((i * 4) + 5)
            sParamMaxValue(Index, i + 1) = temp((i * 4) + 6)
            
            If sParamValue(Index, i + 1) = "" Then
                dParamValue(Index, i + 1) = 0
            Else
                dParamValue(Index, i + 1) = CDbl(sParamValue(Index, i + 1))
            End If
            If sParamMinValue(Index, i + 1) = "" Then
                dParamMinValue(Index, i + 1) = 0
            Else
                dParamMinValue(Index, i + 1) = CDbl(sParamMinValue(Index, i + 1))
            End If
            If sParamMaxValue(Index, i + 1) = "" Then
                dParamMaxValue(Index, i + 1) = 0
            Else
                dParamMaxValue(Index, i + 1) = CDbl(sParamMaxValue(Index, i + 1))
            End If
        Next i
        
        
    Close ff
    Exit Sub
err:
    MsgBox "저장된 Recipe 가 없습니다. Recipe 를 받으세요", vbCritical, "RECIPE 정보"
    Close ff
End Sub
Public Sub DJ_MESMowRecipeSave()
On Error GoTo err
Dim ff As Integer
Dim i As Integer
    ff = FreeFile
    Open App.Path & "\Recipe" & "\NowRecipe" & ".fav" For Output As ff
        Print #ff, sNowRecipeID
        Print #ff, iNowRecipeID
    Close ff
    Exit Sub
err:
    Close ff
End Sub
Public Sub DJ_MESMowRecipeLoad()
On Error Resume Next
Dim i As Integer
Dim ff As Integer
Dim temp(0 To 29) As String
    ff = FreeFile
    Open App.Path & "\Recipe" & "\NowRecipe" & ".fav" For Input As ff  '
        For i = 0 To 1
            Line Input #ff, temp(i)
        Next i
        sNowRecipeID = temp(0)
        iNowRecipeID = temp(1)
    Close ff
    Exit Sub
err:

    Close ff
End Sub
Public Sub DJ_MESRecipeIDCountSave()
On Error GoTo err
Dim ff As Integer
Dim i As Integer
    ff = FreeFile
    Open App.Path & "\Recipe" & "\RecipeIDCount" & ".fav" For Output As ff
        Print #ff, iRecipeIDcount
    Close ff
    Exit Sub
err:
    Close ff
End Sub
Public Sub DJ_MESRecipeIDCountLoad()
On Error Resume Next
Dim i As Integer
Dim ff As Integer
Dim temp(0 To 29) As String
    ff = FreeFile
    Open App.Path & "\Recipe" & "\RecipeIDCount" & ".fav" For Input As ff '
        Line Input #ff, temp(0)
        iRecipeIDcount = temp(0)
    Close ff
    Exit Sub
err:

    Close ff
End Sub
Public Sub DJ_MESmsgLogSave(msg As String)
On Error GoTo err
Dim ff As Integer
Dim lstr_date
Dim lstr_time
Dim length As Integer
Dim MSD_ID As String

    lstr_date = Format(Date, "YYYYMMDD")
    lstr_time = Format(Time, "HHMMSS")
    
    Call Create_DIR("D:\MES\LOG\" & lstr_date)
    msg_id = DJSJ_XMLData_Find(1, "<MSG_ID>", "</MSG_ID>", msg, length)
    ff = FreeFile
    Open "D:\MES" & "\Log\" & lstr_date & "\" & lstr_time & Format((Timer - Int(Timer)) * 100, "00") & "_" & msg_id & ".fav" For Output As ff
        Print #ff, msg
    Close ff
    Exit Sub
err:
    Close ff
End Sub
Public Sub MelsecAddressLoad()
On Error GoTo err

Dim strFilename As String
Dim strSection As String
Dim strKey As String
Dim i, j, k As Long

    strFilename = App.Path & "\Melsec.INI"
    
    strSection = "Parameter"
    nMelsecChannel = CInt(ReadINI(strFilename, strSection, "channel"))
    nMelsecMode = CInt(ReadINI(strFilename, strSection, "mode"))
    
    'strSection = "ADDRESS" & CStr(iProName)
    strSection = "ADDRESS"

    lMelsecAddrInput = ReadINI(strFilename, strSection, "input")
    lMelsecAddrOutput = ReadINI(strFilename, strSection, "output")
    lMelsecAddrCellID = ReadINI(strFilename, strSection, "cellid")
    
    sMelsecAddrModelNumber = ReadINI(strFilename, strSection, "MODELNUMBER")
    
    For i = 0 To kMaxCamera - 1
        sMelsecAddrCellID(i) = ReadINI(strFilename, strSection, "cellid" & CStr(i + 1))
    Next i
    
    lMelsecAddrNgCode = ReadINI(strFilename, strSection, "ngcode")
    
    For i = 0 To 9
        For k = 0 To 3
            strKey = "inspection" & CStr(k) & "-" & CStr(i)
            lMelsecAddrInspection(k, i) = ReadINI(strFilename, strSection, strKey)
        Next k
        
        For j = 0 To 2
            strKey = "base_spec_" & CStr(i) & "_" & CStr(j)
            lMelsecAddrBaseSpec(i, j) = ReadINI(strFilename, strSection, strKey)
        Next j
    Next i
        
    For i = 0 To 3
            strKey = "CELLUSE" & CStr(i)
            lMelsecAddrCelluse(i) = ReadINI(strFilename, strSection, strKey)
    Next i
    
    
    sMelsecAddrAcqDone = ReadINI(strFilename, strSection, "ACQDONE")
    sMelsecAddrAutoRemove = ReadINI(strFilename, strSection, "AUTOREMOVE")
    sMelsecAddrAlarmCamera = ReadINI(strFilename, strSection, "CameraAlarm")
    sMelsecAddrAlarmCIM = ReadINI(strFilename, strSection, "CIMAlarm")
    sMelsecAddrAlarmHDD = ReadINI(strFilename, strSection, "HDDAlarm")
    sMelsecAddrAlarmNetDrive = ReadINI(strFilename, strSection, "NetDriveAlarm")
    
    
    sMelsecAddrAlarm = ReadINI(strFilename, strSection, "Alarm")
    sMelsecAddrZigID = ReadINI(strFilename, strSection, "ZigID")
    
    Exit Sub
err:
    Debug.Print "MelsecAddressLoad Function Failed!"

End Sub


