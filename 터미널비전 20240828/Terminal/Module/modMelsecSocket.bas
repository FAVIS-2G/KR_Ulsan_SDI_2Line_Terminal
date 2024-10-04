Attribute VB_Name = "modMelsecSocket"
Option Explicit

'CONST
Public Const pcl_Qcolor_Red     As Long = 12
Public Const pcl_Qcolor_Grey    As Long = 7
Public Const pcl_Qcolor_White   As Long = 0

Public Const conMelsecWordSize  As Integer = 16

Public Enum MELSEC_ADDRESS
    addControl = 0
    addVisionInspect1 = 1
    addVisionInspect2 = 2
End Enum

Public Enum INLINE_INPUT
    inStart = 0
End Enum

Public Enum INLINE_OUTPUT
    outReadyVision = 0
    outBusyVision = 1
    outEndVision = 2
    outOk1cam = 3
    outNG1cam = 4
    outOk2cam = 5
    outNG2cam = 6
    outOk3cam = 7
    outNG3cam = 8
    outOk4cam = 9
    outNG4cam = 10
    outFinishGrab = 11
    outspare2 = 12
    outspare3 = 13
    outspare4 = 14
    outspare5 = 15
End Enum

Public Enum INLINE_OUTPUT_THICK
    outGoodPressThick = 0
    outNGPressThickUP = 1
    outNGPressThickDOWN = 2
End Enum

Public Enum INLINE_OUTPUT_TAPE
    outGoodBottomTape = 0
    outNGBottomTape = 1
End Enum

'------------------------------------------
'MELSEC 주소변수
Public MELSEC_CONTROL           As String
Public MELSEC_VISION_INSPECT_1  As String
Public MELSEC_VISION_INSPECT_2  As String

'------------------------------------------

Public m_Rcv_Bit(0 To conMelsecWordSize - 1) As Byte
Public m_Snd_Bit_1(0 To conMelsecWordSize - 1) As Byte
Public m_Snd_Bit_2(0 To conMelsecWordSize - 1) As Byte


Public m_bMelsecConnected As Boolean


Public m_bInspectionVision1 As Boolean
Public m_bInspectionVision2 As Boolean

' PLC 이더넷 통신 초기화
Public Function MelsecSocketInit() As Boolean
On Error GoTo ErrHandler

    Dim nResult As Long
    
    Call InitializeMelsecValue
    
    frmMain.ActEasyIF.ActLogicalStationNumber = 1
                                           
    nResult = frmMain.ActEasyIF.Open
    
    If nResult = 0 Then
        MelsecSocketInit = True
        frmMain.shpPLCSock.BackColor = vbGreen
    Else
        MelsecSocketInit = False
        frmMain.shpPLCSock.BackColor = vbRed
    End If
    
    Exit Function
ErrHandler:
    Call WriteErrorLog("MelsecSocketInit" & " : " & err.Description)
End Function

Public Sub MelsecExit()
On Error GoTo ErrHandler

    frmMain.tmrMelsec.Enabled = False

    Dim nResult As Long
    
    nResult = frmMain.ActEasyIF.Close
    
    Exit Sub
    
ErrHandler:
    Call WriteErrorLog("MelsecExit" & " : " & err.Description)
End Sub
Public Sub InitializeMelsecValue()
    Dim i As Integer
    
    For i = 0 To conMelsecWordSize - 1
        m_Rcv_Bit(i) = 0
        m_Snd_Bit_1(i) = 0
        m_Snd_Bit_2(i) = 0
    Next i
    
    m_bInspectionVision1 = False
    m_bInspectionVision2 = False
End Sub

Public Sub ClearMelsecResult(Optional nType As Integer = 1)
    Dim i As Integer
    
    Select Case nType
    Case addVisionInspect1
        For i = 1 To conMelsecWordSize - 1
            m_Snd_Bit_1(i) = 0
        Next i
    Case addVisionInspect2
        For i = 1 To conMelsecWordSize - 1
            m_Snd_Bit_2(i) = 0
        Next i
    End Select
End Sub

Public Sub Test_Send_Word()
    m_Snd_Bit_1(0) = 1
    m_Snd_Bit_1(11) = 1
    
    Call Write_Send_Word(addVisionInspect1, Make_Send_Word(addVisionInspect1, True))
End Sub

'신호를 검색한다.
Public Function CheckInspectionSignal()
On Error GoTo ErrorHandelr
    Dim inspectionTime As Long
    Dim i As Integer
    
    'Ready신호를 Enable.
    Call ClearMelsecResult(addVisionInspect1)
    m_Snd_Bit_1(outReadyVision) = 1
    Call Write_Send_Word(addVisionInspect1, Make_Send_Word(addVisionInspect1, True))
        
    Call ClearMelsecResult(addVisionInspect2)
    m_Snd_Bit_2(outReadyVision) = 1
    Call Write_Send_Word(addVisionInspect2, Make_Send_Word(addVisionInspect2, True))
        
'    While (m_bAutorunSW = True)
'        Call AutoRunFlow
'        DoEvents
'    Wend
'
AutoEnd:
    
    'Ready신호를 Disable.
    m_Snd_Bit_1(outReadyVision) = 0
    Call Write_Send_Word(addVisionInspect1, Make_Send_Word(addVisionInspect1, True))
        
    m_Snd_Bit_2(outReadyVision) = 0
    Call Write_Send_Word(addVisionInspect2, Make_Send_Word(addVisionInspect2, True))
    
    Exit Function
ErrorHandelr:
    WriteLog ("CheckInspectionSignal (" & err.Number & ") : " & err.Description)
End Function



Public Function Make_Send_Word(nType As Integer, Optional bDisplay As Boolean = False) As Long
On Error GoTo ErrHandler
    Dim i As Integer
    Dim nRet As Long
    
    nRet = 0
    
    Select Case nType
    Case addVisionInspect1
        For i = 0 To conMelsecWordSize - 1
            nRet = nRet + 2 ^ i * m_Snd_Bit_1(i)
            If bDisplay = True Then
                If m_Snd_Bit_1(i) = 1 Then
                    frmMain.shpVision2(i).FillColor = QBColor(pcl_Qcolor_Red)
                Else
                    frmMain.shpVision2(i).FillColor = QBColor(pcl_Qcolor_Grey)
                End If
            End If
        Next i
    Case addVisionInspect2
        For i = 0 To conMelsecWordSize - 1
            nRet = nRet + 2 ^ i * m_Snd_Bit_2(i)
            If bDisplay = True Then
                If m_Snd_Bit_2(i) = 1 Then
                   ' frmMain.shpVision2(i).FillColor = QBColor(pcl_Qcolor_Red)
                Else
                    'frmMain.shpVision2(i).FillColor = QBColor(pcl_Qcolor_Grey)
                End If
            End If
        Next i
    
    End Select
    
    Make_Send_Word = nRet
    
    Exit Function
ErrHandler:
    Call WriteErrorLog("Make_Send_Word" & " : " & err.Description)
End Function

Public Sub Write_Send_Word(Optional nType As Integer = 1, Optional nValue As Long = -1)
On Error GoTo ErrHandler

    Dim strDeviceList As String
    Dim nSize As Long
    Dim nData() As Long
    Dim bCheck As Integer
    Dim i As Integer
                
    bCheck = False
    
    If nValue < 0 Then
        nValue = Make_Send_Word(addVisionInspect1, True)
    End If
  
    Select Case nType
    Case addVisionInspect1
        'MELSEC_VISION_INSPECT_1 = "R9998"
        strDeviceList = lMelsecAddrOutput
        bCheck = True
    Case addVisionInspect2
        strDeviceList = MELSEC_VISION_INSPECT_2
        bCheck = True
    End Select
    
    nSize = 1
    ReDim nData(nSize)
    
    nData(0) = nValue
    
    If bCheck = True Then
        Call frmMain.ActEasyIF.WriteDeviceRandom(strDeviceList, nSize, nData(0))
    End If

    Exit Sub
ErrHandler:
    Call WriteErrorLog("Write_Send_Word" & " : " & err.Description)
End Sub
'PLC 비트 정보 읽기 (화면에 표시)
Public Sub Read_Recieve_Bit(nValue As Long)
On Error GoTo ErrHandler
    Dim i As Integer
    Dim tempstr As String
    Dim tempSplit

    tempstr = ChangeToBit(CInt(nValue))
    tempSplit = Split(tempstr, ",")
    
    For i = 0 To 3
        m_Rcv_Bit(i) = CByte(tempSplit(i))
        
        'If m_bAutorunSW = True Then
            If m_Rcv_Bit(i) = 1 Then
                frmMain.shpInput(i).FillColor = QBColor(pcl_Qcolor_Red)
            Else
                frmMain.shpInput(i).FillColor = QBColor(pcl_Qcolor_Grey)
            End If
        'End If
    Next i

    Exit Sub
ErrHandler:
    Call WriteErrorLog("Read_Recieve_Bit" & " : " & err.Description)
End Sub
'PLC 비트 정보 변경
Public Function ChangeToBit(iValue As Integer) As String
On Error GoTo ErrHandler
    Dim i, iShift As Integer
    Dim bit_Status(1 To 9) As Byte
    
    ChangeToBit = ""
    iShift = 1
    'Check every Digtial data bit
    For i = 0 To 8
            'Check, it is change to 1 or 0
        If (iValue And iShift) = iShift Then
            bit_Status(i + 1) = 1
            If i = 0 Then
                ChangeToBit = ChangeToBit & "1"
            Else
                ChangeToBit = ChangeToBit & ",1"
            End If
        Else
            bit_Status(i + 1) = 0
             If i = 0 Then
                ChangeToBit = ChangeToBit & "0"
            Else
                ChangeToBit = ChangeToBit & ",0"
            End If
        End If

        'Check next bit
        iShift = iShift * 2
    Next
    Exit Function
    
ErrHandler:
    ChangeToBit = ""
    Call WriteErrorLog("ChangeToBit" & " : " & err.Description)
End Function
Public Function ReadDataFromPLC()
On Error GoTo err

    Dim strDeviceList As String
    Dim strType As String
    Dim Addr As Long
    Dim i As Integer
    Dim j As Integer
    Dim nSize As Long
    Dim nData() As Long
    Dim nResult As Long
    Dim nTemp As Integer
    
    Dim Index As Integer
    Dim cellID As String
    
    Dim Count As Integer
    
    strDeviceList = ""
    
    nTemp = 1
        
    
'    strType = Left$(lMelsecAddrCellID, 2)
'    Addr = CLng(Right$(lMelsecAddrCellID, Len(lMelsecAddrCellID) - 2))
'
'    For i = 1 To 80
'        If i <> 1 Then
'            strDeviceList = strDeviceList & vbLf
'        End If
'
'        strDeviceList = strDeviceList & strType & Format(Addr, "0")
'        Addr = Addr + 1
'    Next i
    
    strDeviceList = strDeviceList & GetAddressString(sMelsecAddrCellID(0), 10)
    strDeviceList = strDeviceList & vbLf & GetAddressString(sMelsecAddrCellID(1), 10)
    strDeviceList = strDeviceList & vbLf & GetAddressString(sMelsecAddrCellID(2), 10)
    strDeviceList = strDeviceList & vbLf & GetAddressString(sMelsecAddrCellID(3), 10)
    
    For i = 0 To 9
        For j = 0 To 2
            'Debug.Print i, j
            strDeviceList = strDeviceList & vbLf & GetAddressString(lMelsecAddrBaseSpec(i, j), 2)
        Next j
    Next i
    
    strDeviceList = strDeviceList & vbLf & GetAddressString(sMelsecAddrZigID, 4)
    
'    For j = 0 To 3
'            'Debug.Print i, j
'            strDeviceList = strDeviceList & vbLf & GetAddressString(lMelsecAddrCelluse(j), 1)
'    Next j
    
    
    Debug.Print strDeviceList
    
    For i = 0 To Len(strDeviceList)
        If Right$(Left$(strDeviceList, i + 1), 1) = vbLf Then
            Count = Count + 1
        End If
    Next i
    
    nSize = Count + 1
    
    ReDim nData(nSize)
    
    Dim stime As Long
    Dim etime As Long
    
    stime = GetTickCount
    nResult = frmMain.ActEasyIF.ReadDeviceRandom(strDeviceList, nSize, nData(0))
    etime = GetTickCount
    Debug.Print etime - stime
    
    For i = 0 To 3
        cellID = ""
        For j = 0 To 9
            Index = (i * 10) + j
            cellID = cellID & Chr(nData(Index) Mod 256) & Chr(nData(Index) / 256)
        Next j
        sIDCode(i) = Trim(Replace(cellID, Chr(0), " "))
        frmMain.lblIDCodeNum(i).Caption = cellID
    Next i
    
    
    For i = 0 To 59 Step 6
        Index = i + 40
        j = i / 6
        If j < 4 Then
            dSpecOri(j) = nData(Index + 0) / 100# '20240828 KCG /10# -> /100# 으로 변경
            dSpecMax(j) = nData(Index + 2) / 100# '20240828 KCG /10# -> /100# 으로 변경
            dSpecMin(j) = nData(Index + 4) / 100# '20240828 KCG /10# -> /100# 으로 변경
        Else
            dSpecOri(j) = nData(Index + 0) / 100#
            dSpecMax(j) = nData(Index + 2) / 100#
            dSpecMin(j) = nData(Index + 4) / 100#
        End If
'        dSpecOri(j) = 0
'        dSpecMax(j) = 1000000
'        dSpecMin(j) = 1000000
        dSpecOriMax(j) = dSpecOri(j) + dSpecMax(j)
        dSpecOriMin(j) = dSpecOri(j) - dSpecMin(j)
    Next i
    
    cellID = ""
    For i = 0 To 3
        Index = 40 + 60 + i
        cellID = cellID & Chr(nData(Index) Mod 256) & Chr(nData(Index) / 256)
    Next i
    sZigID = Trim(Replace(cellID, Chr(0), " "))
    
'    For i = 0 To 9
'        If frmMain.chkSpecPass(i).Value = 1 Then
'            frmMain.txtSpecOri(i).Text = Format(dSpecOri(i), "0.00")
'            frmMain.txtSpecMin(i).Text = Format(dSpecMin(i), "0.00")
'            frmMain.txtSpecMax(i).Text = Format(dSpecMax(i), "0.00")
'        Else
'            frmMain.txtSpecOri(i).Text = Format(0, "0.00")
'            frmMain.txtSpecMin(i).Text = Format(0, "0.00")
'            frmMain.txtSpecMax(i).Text = Format(0, "0.00")
'        End If
'    Next i
    
'    For i = 0 To 3
'        Index = i + 40 + 60
'        dCelluse(i) = nData(Index)
'    Next i
    'Cell ID
    
'    frmMain.lblIDCodeNum.Caption = "TEST"
'
'    '기준값
'    For i = 0 To 9
'        index = 10 + i * 3 * 2
'
'        If frmMain.chkSpecPass(i).Value = 1 Then
'            frmMain.txtSpecOri(i).Text = CStr(buf(index + 0))
'            frmMain.txtSpecMax(i).Text = CStr(buf(index + 2))
'            frmMain.txtSpecMin(i).Text = CStr(buf(index + 4))
'        Else
'            frmMain.txtSpecOri(i).Text = "0.00"
'            frmMain.txtSpecMax(i).Text = "0.00"
'            frmMain.txtSpecMin(i).Text = "0.00"
'
'        End If
'    Next i
    Exit Function
err:
    
End Function

Public Function GetAddressString(Address As String, size As Integer) As String
    
    Dim strResult As String
    Dim strType As String
    Dim Addr As Long
    Dim nTemp As Integer
    
    Dim i As Integer
    
    If Address = "" Then
        Address = "D0"
    End If
    
    nTemp = 1
    
    If IsNumeric(Right(Left(Address, 2), 1)) = False Then
        nTemp = 2
    End If
    
    strType = Left$(Address, nTemp)
    
    
    Addr = CLng(Right$(Address, Len(Address) - nTemp))
    
    For i = 1 To size
        If i <> 1 Then
            strResult = strResult & vbLf
        End If
        
        strResult = strResult & strType & Format(Addr, "0")
        Addr = Addr + 1
    Next i
    
    GetAddressString = strResult
    
End Function

Public Function WriteDataToPLC()
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
            Debug.Print k, i, lMelsecAddrInspection(k, i)
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

