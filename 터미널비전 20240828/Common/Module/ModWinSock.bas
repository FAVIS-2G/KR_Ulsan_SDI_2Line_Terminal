Attribute VB_Name = "ModWinSock"

Public Sub Win_Connect(Index As Integer)

On Error GoTo err

    Dim tm As Single
    
    Select Case Index
    Case 0
        ListBox_Append Time & "  " & sPLCIP & " 컴퓨터 " & sPLCPort & " 포트에 연결중...", Index
'        frmMain.lstPLCSocket.Refresh
'        frmMain.WinsockPLC.RemoteHost = sPLCIP
'        frmMain.WinsockPLC.RemotePort = sPLCPort
'        frmMain.WinsockPLC.Connect
'
'        tm = Timer + 2
'        Do
'            DoEvents
'        Loop Until (bWinsock(Index) = True Or tm < Timer)
'        If bWinsock(Index) = False Then GoTo err:
'        ListBox_Append Time & "  " & sPLCIP & "컴퓨터 " & sPLCPort & "포트에 연결 성공.", Index
'        frmMain.lstPLCSocket.Refresh
    Case 1
        ListBox_Append Time & "  " & sMESIP & " 컴퓨터 " & sMESPort & " 포트에 연결중...", Index
'        frmMain.lstMESSocket.Refresh
        
        frmMain.WinsockMES.RemoteHost = sMESIP
        frmMain.WinsockMES.RemotePort = sMESPort
        frmMain.WinsockMES.Connect
        tm = Timer + 2
        Do
            DoEvents
        Loop Until (bWinsock(Index) = True Or tm < Timer)
        If bWinsock(Index) = False Then GoTo err:
        ListBox_Append Time & "  " & sMESIP & "컴퓨터 " & sMESPort & "포트에 연결 성공.", Index
        frmMain.lstMESSocket.Refresh
    End Select
    
Exit Sub

err:
    Select Case Index
    Case 0
'        tm = Timer + 1
'        Do
'            DoEvents
'        Loop Until tm < Timer
'        frmMain.WinsockPLC.Close
'        ListBox_Append Time & " 연결 실패", Index
'        bWinsock(Index) = False
'        frmMain.lstPLCSocket.Refresh
    Case 1
        tm = Timer + 1
        Do
            DoEvents
        Loop Until tm < Timer
        frmMain.WinsockMES.Close
        ListBox_Append Time & " 연결 실패", Index
        bWinsock(Index) = False
        frmMain.lstMESSocket.Refresh
    End Select
End Sub

Public Sub Win_Disable(Index As Integer)

On Error GoTo err

Dim tm As Single
    Select Case Index
    Case 0
'        frmMain.WinsockPLC.Close
'        tm = Timer + 1
'        Do
'            DoEvents
'        Loop Until tm < Timer
'        ListBox_Append "연결을 종료", Index
'        bWinsock(Index) = False
'        frmMain.lstPLCSocket.Refresh
    Case 1
        frmMain.WinsockMES.Close
        tm = Timer + 1
        Do
            DoEvents
        Loop Until tm < Timer
        ListBox_Append "연결을 종료", Index
        bWinsock(Index) = False
        frmMain.lstMESSocket.Refresh
    End Select
Exit Sub

err:
    bWinsock(Index) = False
End Sub
'
'Public Sub QJ71E71Connect()
'On Error GoTo err:
''Dim ret As Long
''Dim Error_Code As String
''    sPLCIP = "192.168.0.1"
''    ListBox_Append Time & "  " & sPLCIP & " 컴퓨터 " & sPLCPort & " 포트에 연결중.", 0
''    ret = frmMain.ActQJ71E71TCP.Open
''    If ret = 0 Then
''        ListBox_Append Time & "  " & sPLCIP & "컴퓨터 " & sPLCPort & "포트에 연결 성공.", 0
''        frmMain.lstPLCSocket.Refresh
''    Else
''        Error_Code = Hex$(lRet)
''        If (Error_Code = "180840B") Then
''            ListBox_Append Time & "  " & "타임아웃 에러!! 통신상태를 확인하세요!", 0
''            GoTo err:
''        Else
''            ListBox_Append Time & "  " & "에러 발생(에러 코드：" + Error_Code + ")", 0
''            GoTo err:
''        End If
''    End If
'Exit Sub
'err:
'End Sub
'
'Public Sub QJ71E71DisConnect()
'On Error GoTo err:
''Dim ret As Long
''Dim Error_Code As String
''    ListBox_Append Time & "  " & sPLCIP & " 컴퓨터 " & sPLCPort & " 포트를 종료합니다.", 0
''    ret = frmMain.ActQJ71E71TCP.Close
''    If ret = 0 Then
''        ListBox_Append Time & "  " & sPLCIP & "컴퓨터 " & sPLCPort & "포트에 종료 성공.", 0
''        frmMain.lstPLCSocket.Refresh
''    Else
''        Error_Code = Hex$(lRet)
''        If (Error_Code = "F0000004") Then
''            ListBox_Append Time & "  " & "타임아웃 에러!! 통신상태를 확인하세요!", 0
''            GoTo err:
''        Else
''            ListBox_Append Time & "  " & "에러 발생(에러 코드：" + Error_Code + ")", 0
''            GoTo err:
''        End If
''    End If
'Exit Sub
'err:
'End Sub
'
'Public Function QJ71E71ReadData(sAddr As String, lsize As Long, lAddrSize As Integer) As String
'Dim ret As Long
'Dim sdata As Long
'    ret = frmMain.ActQJ71E71TCP.ReadDeviceBlock(frmMain.txtAddr_Rg.Text & CStr(CLng(sAddr) + lAddrSize), lsize, sdata)
'    If ret = 0 Then
'        ListBox_Append Time & "   " & frmMain.txtAddr_Rg.Text & CLng(sAddr) + lAddrSize & " 에서 " & sdata & " 를 받았습니다.", 0
'        frmMain.lstPLCSocket.Refresh
'        QJ71E71ReadData = sdata
'        sIDCode_Q71(lAddrSize) = sdata
'    Else
'         ListBox_Append Time & "   " & frmMain.txtAddr_Rg.Text & CLng(sAddr) + lAddrSize & " 에서 " & " 데이터 받기를 실패하였습니다.", 0
'         QJ71E71ReadData = sdata
'         sIDCode_Q71(lAddrSize) = sdata
'         Call QJ71E71DisConnect
'         Call QJ71E71Connect
'    End If
'End Function
'
'Public Function QJ71E71WriteData(sAddr As String, sdata As Long) As Boolean
'Dim ret As Long
'    ret = frmMain.ActQJ71E71TCP.WriteDeviceBlock(frmMain.txtAddr_Rg.Text & sAddr, 1, sdata)
'    If ret = 0 Then
'        ListBox_Append Time & "   " & frmMain.txtAddr_Rg.Text & sAddr & " 에 " & sdata & " 를 보냈습니다.", 0
'        frmMain.lstPLCSocket.Refresh
'        QJ71E71WriteData = True
'    Else
'         ListBox_Append Time & "   " & frmMain.txtAddr_Rg.Text & sAddr & " 에서 " & " 데이터 받기를 실패하였습니다.", 0
'         QJ71E71WriteData = False
'    End If
'End Function
