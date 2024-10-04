Attribute VB_Name = "ModPing"
Option Explicit

Private Const IP_SUCCESS As Long = 0
Private Const IP_STATUS_BASE As Long = 11000
Private Const IP_BUF_TOO_SMALL As Long = (11000 + 1)
Private Const IP_DEST_NET_UNREACHABLE As Long = (11000 + 2)
Private Const IP_DEST_HOST_UNREACHABLE As Long = (11000 + 3)
Private Const IP_DEST_PROT_UNREACHABLE As Long = (11000 + 4)
Private Const IP_DEST_PORT_UNREACHABLE As Long = (11000 + 5)
Private Const IP_NO_RESOURCES As Long = (11000 + 6)
Private Const IP_BAD_OPTION As Long = (11000 + 7)
Private Const IP_HW_ERROR As Long = (11000 + 8)
Private Const IP_PACKET_TOO_BIG As Long = (11000 + 9)
Private Const IP_REQ_TIMED_OUT As Long = (11000 + 10)
Private Const IP_BAD_REQ As Long = (11000 + 11)
Private Const IP_BAD_ROUTE As Long = (11000 + 12)
Private Const IP_TTL_EXPIRED_TRANSIT As Long = (11000 + 13)
Private Const IP_TTL_EXPIRED_REASSEM As Long = (11000 + 14)
Private Const IP_PARAM_PROBLEM As Long = (11000 + 15)
Private Const IP_SOURCE_QUENCH As Long = (11000 + 16)
Private Const IP_OPTION_TOO_BIG As Long = (11000 + 17)
Private Const IP_BAD_DESTINATION As Long = (11000 + 18)
Private Const IP_ADDR_DELETED As Long = (11000 + 19)
Private Const IP_SPEC_MTU_CHANGE As Long = (11000 + 20)
Private Const IP_MTU_CHANGE As Long = (11000 + 21)
Private Const IP_UNLOAD As Long = (11000 + 22)
Private Const IP_ADDR_ADDED As Long = (11000 + 23)
Private Const IP_GENERAL_FAILURE As Long = (11000 + 50)
Private Const MAX_IP_STATUS As Long = (11000 + 50)
Private Const IP_PENDING As Long = (11000 + 255)
Private Const PING_TIMEOUT As Long = 500
Private Const WS_VERSION_REQD As Long = &H101
Private Const MIN_SOCKETS_REQD As Long = 1
Private Const SOCKET_ERROR As Long = -1
Private Const INADDR_NONE As Long = &HFFFFFFFF
Private Const MAX_WSADescription As Long = 256
Private Const MAX_WSASYSStatus As Long = 128

Private Type ICMP_OPTIONS
    Ttl             As Byte
    Tos             As Byte
    Flags           As Byte
    OptionsSize     As Byte
    OptionsData     As Long
End Type

Public Type ICMP_ECHO_REPLY
    Address         As Long
    status          As Long
    RoundTripTime   As Long
    DataSize        As Long
    DataPointer     As Long
    Options         As ICMP_OPTIONS
    Data            As String * 250
End Type

Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To MAX_WSADescription) As Byte
    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
    wMaxSockets As Long
    wMaxUDPDG As Long
    dwVendorInfo As Long
End Type

Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long

Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long

Private Declare Function IcmpSendEcho Lib "icmp.dll" _
       (ByVal IcmpHandle As Long, _
        ByVal DestinationAddress As Long, _
        ByVal RequestData As String, _
        ByVal RequestSize As Long, _
        ByVal RequestOptions As Long, _
        ReplyBuffer As ICMP_ECHO_REPLY, _
        ByVal ReplySize As Long, _
        ByVal Timeout As Long) As Long

Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long

Private Declare Function WSAStartup Lib "WSOCK32.DLL" _
       (ByVal wVersionRequired As Long, _
        lpWSADATA As WSADATA) As Long

Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Private Declare Function gethostname Lib "WSOCK32.DLL" _
       (ByVal szHost As String, _
        ByVal dwHostLen As Long) As Long

Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal szHost As String) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
       (xDest As Any, _
        xSource As Any, _
        ByVal nbytes As Long)

Private Declare Function inet_addr Lib "WSOCK32.DLL" (ByVal s As String) As Long


Public Function GetStatusCode(status As Long) As String
    Dim msg As String

    Select Case status
        Case IP_SUCCESS:               msg = "ip success"
        Case INADDR_NONE:              msg = "inet_addr: bad IP format"
        Case IP_BUF_TOO_SMALL:         msg = "ip buf too_small"
        Case IP_DEST_NET_UNREACHABLE:  msg = "ip dest net unreachable"
        Case IP_DEST_HOST_UNREACHABLE: msg = "ip dest host unreachable"
        Case IP_DEST_PROT_UNREACHABLE: msg = "ip dest prot unreachable"
        Case IP_DEST_PORT_UNREACHABLE: msg = "ip dest port unreachable"
        Case IP_NO_RESOURCES:          msg = "ip no resources"
        Case IP_BAD_OPTION:            msg = "ip bad option"
        Case IP_HW_ERROR:              msg = "ip hw_error"
        Case IP_PACKET_TOO_BIG:        msg = "ip packet too_big"
        Case IP_REQ_TIMED_OUT:         msg = "ip req timed out"
        Case IP_BAD_REQ:               msg = "ip bad req"
        Case IP_BAD_ROUTE:             msg = "ip bad route"
        Case IP_TTL_EXPIRED_TRANSIT:   msg = "ip ttl expired transit"
        Case IP_TTL_EXPIRED_REASSEM:   msg = "ip ttl expired reassem"
        Case IP_PARAM_PROBLEM:         msg = "ip param_problem"
        Case IP_SOURCE_QUENCH:         msg = "ip source quench"
        Case IP_OPTION_TOO_BIG:        msg = "ip option too_big"
        Case IP_BAD_DESTINATION:       msg = "ip bad destination"
        Case IP_ADDR_DELETED:          msg = "ip addr deleted"
        Case IP_SPEC_MTU_CHANGE:       msg = "ip spec mtu change"
        Case IP_MTU_CHANGE:            msg = "ip mtu_change"
        Case IP_UNLOAD:                msg = "ip unload"
        Case IP_ADDR_ADDED:            msg = "ip addr added"
        Case IP_GENERAL_FAILURE:       msg = "ip general failure"
        Case IP_PENDING:               msg = "ip pending"
        Case PING_TIMEOUT:             msg = "ping timeout"
        Case Else:                     msg = "unknown  msg returned"
    End Select

    GetStatusCode = CStr(status) & "   [ " & msg & " ]"
End Function

Public Function Ping(sAddress As String, sDataToSend As String, ECHO As ICMP_ECHO_REPLY) As Boolean
    ' Ping 성공시
    ' RoundTripTime = Ping 가 성공한 시간 리턴
    ' Data = 리턴된 데이타
    ' Address = IP 의 실제 값
    ' DataSize = Data 의 크기
    ' Status = 0
    ' Ping 실패시 에러코드 리턴

    Dim hPort As Long
    Dim dwAddress As Long

    On Error GoTo ErrHandle
    
    ' IP 값을 Long 값으로 변환한다.
    dwAddress = inet_addr(sAddress)

    If dwAddress <> INADDR_NONE Then        ' 유효한 IP Address 이면
        ' 포드틀 오픈한다.
        hPort = IcmpCreateFile()

        ' 포트오픈이 성공하면
        If hPort Then
            ' Ping 시도
            Call IcmpSendEcho(hPort, dwAddress, sDataToSend, Len(sDataToSend), _
                             0, ECHO, Len(ECHO), PING_TIMEOUT)

            ' Ping 의 상태값을 리턴한다.
            If ECHO.status = 0 Then
                If Abs(ECHO.Address) > 0 Then
                    Ping = True
                Else
                    Ping = False
                End If
            Else
                Ping = False
            End If
            Call IcmpCloseHandle(hPort)     ' 포트를 닫는다
        End If
    Else
        ' 잘못된 IP 의 경우
        Ping = INADDR_NONE
    End If
    Exit Function
ErrHandle:
'    Call ErrProcess("User Function - Ping()")
End Function

Public Sub SocketsCleanup()
    If WSACleanup() <> 0 Then
        MsgBox "윈도우 소켓의 CleanUp 실패", vbExclamation
    End If
End Sub

Public Function SocketsInitialize() As Boolean
    Dim WSAD As WSADATA
    
    On Error GoTo ErrHandle

    SocketsInitialize = WSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS
    
    Exit Function
ErrHandle:
'    Call ErrProcess("User Function - SocketsInitialize()")
End Function

Public Function NetWork_MDB() As Boolean
'******************************************
'   네트워크 연결상태 MDB 생성
'******************************************
        
    Dim sDbPath                 As String
    Dim sFileName               As String
    Dim ClsSys                  As New ClsSystem
    Dim sSystemFolder           As String
    Dim i                       As Integer
    
    sFileName = "HJ_NET" & Mid(gtSetting.sNPSName, 2, 2)
    sDbPath = App.Path & "\DATA\DAT\" & sFileName & ".MDB"
    
    If ClsSys.SearchFile(sDbPath) = False Then
        '=====  비 존재
        '*************************************************************
        sSystemFolder = "C:\Windows\system"
            
        For i = 1 To 5
            If ClsSys.CopyFiles(sSystemFolder & "\" & "JCDF" & "\" & "HJ_NETON.MDB", _
                                sDbPath, True) = True Then
                NetWork_MDB = True
                Exit For
            End If
            NetWork_MDB = False
        Next i
    Else
        NetWork_MDB = True
    End If

    Set ClsSys = Nothing

End Function

Public Function NetWork_MDB_Load() As Boolean
'*********************************************************************
'   판매 그린드 복사
'   판매 데이터 복사
'*********************************************************************

    Dim i               As Integer
    Dim sDay            As String
    Dim ECHO            As ICMP_ECHO_REPLY
    Dim bCheck          As Boolean
    Dim bNet            As Boolean
    Dim sSource         As String
    Dim sDestination    As String
    Dim sSou            As String
    Dim sDes            As String
    Dim ClsSys          As ClsSystem
    Dim sJuDay                          As String
    Dim sMDBFile                        As String
    
    On Error GoTo ErrHandle
    
    Set ClsSys = New ClsSystem
    
    bNet = False
    NetWork_MDB_Load = False
    
    sJuDay = Jual_Data(Format(Now, "YYYY-MM-DD"))
    sMDBFile = Mid(gtSetting.sNPSName, 2, 2) & sJuDay
    
    sSource = App.Path & "\DATA\GRD\SALE\" & sJuDay
    sDestination = gtsServer_Path & "\DATA\SALE\GRD\" & gtSetting.sNPSName & "\"
    '************************************************************************
    '   네트워크의 단절된 날짜의 Grd와 판매DB를 모두 전송
    '************************************************************************
    
    For i = 1 To 5
        If SocketsInitialize() Then
             bCheck = Ping(gtSetting.sNPSIP, "", ECHO)
             If bCheck Then
                bNet = True
                Exit For
             End If
        End If
        Sleep (500)
    Next i
    
    If bNet Then
        If ClsSys.SearchFolder(sSource & "\") = True Then
            If ClsSys.CopyFolder(sSource, sDestination) Then
                
                '=====  판매 DB 복사
                '========================================================
                sSou = App.Path & "\DATA\SALE\" & "S" & sMDBFile & ".MDB"
                sDes = gtsServer_Path & "\DATA\SALE\GRD\" & gtSetting.sNPSName & "\" & sJuDay & "\" & "S" & sMDBFile & ".MDB"
                
                If ClsSys.SearchFile(sSou) Then
                    If ClsSys.CopyFiles(sSou, sDes, False) Then
                        NetWork_MDB_Load = True
                    End If
                Else
                    NetWork_MDB_Load = True
                End If
                            
                '=====  정산 DB 복사
                '=========================================================
                sSou = App.Path & "\DATA\ADJUST\" & "A" & sMDBFile & ".MDB"
                sDes = gtsServer_Path & "\DATA\SALE\GRD\" & gtSetting.sNPSName & "\" & sJuDay & "\" & "A" & sMDBFile & ".MDB"
                
                If ClsSys.SearchFile(sSou) Then
                    If ClsSys.CopyFiles(sSou, sDes, False) Then
                        NetWork_MDB_Load = True
                    End If
                Else
                    NetWork_MDB_Load = True
                End If
            End If
        End If
    End If

    Exit Function
ErrHandle:
    Call ErrProcess("User Function - NetWork_MDB_Load()")
End Function

Public Function Ping_On(bOn As Boolean) As Boolean
'=========================================================
'   DB를 Open
'   네트워크 연결시는 네트워크 DB를 로그온해서 사용
' Parameter         1.bOn       :   연결여부
'=========================================================
    Dim ECHO                As ICMP_ECHO_REPLY
    Dim bCheck              As Boolean
    Dim bPing               As Boolean
    Dim i                   As Integer
    Dim sMDBFile            As String
    
    Ping_On = False
    
    On Error GoTo ErrHandle
    
    If bOn Then
        '=====  네트워크 연결중
        FrmNet.timShow.Enabled = True
        FrmNet.pgbStatus.Value = 0
        FrmNet.panNetMsg = "네트워크 연결중"
        
        If SocketsInitialize() Then
            bCheck = Ping(gtSetting.sNPSIP, "", ECHO)
            '=====  Ping으로 체크
            If bCheck Then
                Ping_On = True
                '=====  상품을 UpData
                frmMain.timUpData.Enabled = True
                
                '====================================================
                '   네트워크의 연결여부를 Check
                '====================================================
                '=====  마스터 파일 복사
                '=====  복사하기전에 기존에 Open한 MDB를 닫기=>오픈
'                FrmNet.panNetMsg = "마스터 복사"
'                FrmNet.pgbStatus.Value = 1000
'                Call HJMDB_Copy

                '=====  MDB 초기화
                FrmNet.panNetMsg = "신상품 전송"
                FrmNet.pgbStatus.Value = 1000
'                Call MDB_Setting(True)
                Call NewCommo_Send
                
                FrmNet.panNetMsg = "가격변경"
                FrmNet.pgbStatus.Value = 2000
                Call ChangCommo_Send
                
                '=====================================================================
                '   네트워크가 단절된 내용을 DB에서 검색을 해서
                '   단절된 날짜별로 복사
                '=====================================================================
                '=====  판매 데이터 전송중
                FrmNet.pgbStatus.Value = 3000
                FrmNet.panNetMsg = "판매 데이터 전송중"
                '=====  판매 데이터 전송
                '=====  판매 그리드 전송
                
                If NetWork_MDB_Load = False Then
                    MsgBox "네트워크에 문제가 있습니다." & Chr(10) & Chr(13) & _
                            "네트워크를 확인하고 다시 시도해 주십시요.", , "파일 전송"

                    FrmNet.pgbStatus.Value = 5500
                    FrmNet.panNetMsg = "네트워크 단절중"

                    FrmNetOff.panNetMsg = "네트워크 단절중"
                    FrmNetOff.pgbStatus.Value = 1000
                    FrmNetOff.pgbStatus.Value = 1500

    '                Call MDB_Setting(False)
                    gtbPing = False
                    gtbNetOn = False
                    Ping_On = False
                    Exit Function
                End If
                
                FrmNet.pgbStatus.Value = 3000
'                FrmNet.panNetMsg = "데이터 업데이트중"
                FrmNet.panNetMsg = "데이터 재정의"
'                Call Login_MDB
                
                gtbPing = True
                Ping_On = True
                FrmNet.pgbStatus.Value = 5000
            Else
                FrmNet.pgbStatus.Value = 5500
                FrmNet.panNetMsg = "네트워크 단절중"
                
                FrmNetOff.panNetMsg = "네트워크 단절중"
                FrmNetOff.pgbStatus.Value = 1000
                FrmNetOff.pgbStatus.Value = 1500
'                Call MDB_Setting(False)
                gtbPing = False
                gtbNetOn = False
                
'                If gtbOff Then
'                    gtbOff = False
                    FrmMDB_Nothing.Show 1
'                Else
            End If
        End If
    Else
        FrmNetOff.panNetMsg = "네트워크 단절중"
        FrmNetOff.pgbStatus.Value = 1000
        FrmNetOff.pgbStatus.Value = 1500
        
'        Call MDB_Setting(False
        gtbPing = False
        gtbNetOn = False
    End If
    
    Exit Function
ErrHandle:
    Call ErrProcess("User Function - Ping_On()")
End Function

Public Function Data_Copy()
'=========================================================
'   DB를 Open
'   네트워크 연결시는 MDB를 복사
'=========================================================
    Dim ECHO                As ICMP_ECHO_REPLY
    Dim bCheck              As Boolean
    Dim bPing               As Boolean
    Dim i                   As Integer
    Dim sMDBFile            As String

    '=====  네트워크 연결중
    FrmMDB_Copy.timShow.Enabled = True
    FrmMDB_Copy.pgbStatus.Value = 0
    FrmMDB_Copy.panNetMsg = "네트워크 연결중"
    
    If SocketsInitialize() Then
        bCheck = Ping(gtSetting.sNPSIP, "", ECHO)
        '=====  Ping으로 체크
        If bCheck Then
            '=====  마스터 파일 복사
            '=====  복사하기전에 기존에 Open한 MDB를 닫기=>오픈
            FrmMDB_Copy.panNetMsg = "마스터 복사"
            FrmMDB_Copy.pgbStatus.Value = 1000
            gtbPing = True
            Call HJMDB_Copy
            
            '=====  MDB 초기화
            FrmMDB_Copy.panNetMsg = "데이터 재정의"
            FrmMDB_Copy.pgbStatus.Value = 3000
'            Call MDB_Setting(True)
            
            FrmMDB_Copy.pgbStatus.Value = 5000
        Else
            FrmMDB_Copy.panNetMsg = "네트워크 단절중"
            FrmMDB_Copy.pgbStatus.Value = 5000
        End If
    End If
    
End Function

Public Sub Net_Check()
'*******************************************************
'   단절된 상태에서 네트워크를 검색해서 연결
'*******************************************************
    Dim ECHO                As ICMP_ECHO_REPLY
    Dim bCheck              As Long

    On Error GoTo ErrHandle
    
    If SocketsInitialize() Then
        bCheck = Ping(gtSetting.sNPSIP, "", ECHO)
        '=====  Ping으로 체크
        If bCheck Then
            gtbNetOn = True
            FrmNet.Show 1
        Else
            MsgBox "네트워크 상태에 문제가 있습니다.", , "네트워크"
            gtbPing = False
            gtbNetOn = False
        End If
    Else
        MsgBox "네트워크 상태에 문제가 있습니다.", , "네트워크"
        gtbPing = False
        gtbNetOn = False
    End If

    frmMain.txtScan.SetFocus

Exit Sub
ErrHandle:
    Call ErrProcess("User Sub - Net_Check()")
End Sub

'Public Sub MDB_Setting(bPing As Boolean)
'    Dim ECHO            As ICMP_ECHO_REPLY
'    Dim bCheck          As Boolean
'    Dim sMDBFile        As String
'    Dim sFolder         As String
'
'    On Error Resume Next
'
'    If bPing Then
'        If SocketsInitialize() Then
'             bCheck = Ping(gtSetting.sNPSIP, "", ECHO)
'             If bCheck Then
'                gtbPing = True
'                '=====  상품마스터
'                gtDB_N_Commo.Close
'                sMDBFile = gtsServer_Path & "\MASTER\HJ_COMCT.mdb"
'                Set gtDB_N_Commo = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'                '=====  고객마스터
'                gtDB_N_Client.Close
'                sMDBFile = gtsServer_Path & "\MASTER\HJ_CLIEN.mdb"
'                Set gtDB_N_Client = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'                '=====  기타코드 마스터
'                gtDB_N_EtcCode.Close
'                sMDBFile = gtsServer_Path & "\MASTER\HJ_CODES.mdb"
'                Set gtDB_N_EtcCode = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'                '=====  신상품마스터
'                gtDB_N_NewCommo.Close
'                sMDBFile = gtsServer_Path & "\MASTER\HJ_NWCOM.mdb"
'                Set gtDB_N_NewCommo = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'                '=====  환경설정
'                gtDB_N_Setting.Close
'                sMDBFile = gtsServer_Path & "\MASTER\HJ_SETTI.mdb"
'                Set gtDB_N_Setting = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'                '=====  CMS마스터
'                gtDB_N_CMS.Close
'                sMDBFile = gtsServer_Path & "\MASTER\HJ_DISCU.mdb"
'                Set gtDB_N_CMS = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'                Exit Sub
'            End If
'        End If
'    End If
'
'    '=====  상품마스터
'    FrmNetOff.panNetMsg = "데이터 재정의"
'    FrmNetOff.pgbStatus.Value = 2000
'
'    gtDBCommo.Close
'    sMDBFile = app.path & "\MASTER\HJ_COMCT.mdb"
'    Set gtDBCommo = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'    FrmNetOff.pgbStatus.Value = 3000
'
'    '=====  고객마스터
'    gtDBClient.Close
'    sMDBFile = app.path & "\MASTER\HJ_CLIEN.mdb"
'    Set gtDBClient = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'    FrmNetOff.pgbStatus.Value = 3500
'
'    FrmNetOff.panNetMsg = "데이터 재정의 중"
'    '=====  기타코드 마스터
'    gtDBEtcCode.Close
'    sMDBFile = app.path & "\MASTER\HJ_CODES.mdb"
'    Set gtDBEtcCode = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'    gtSQL = "SELECT TBL7TPCD FROM TBHJCL7 WHERE TBL7CTCD='" & 88006611 & "'"
'    Set gtReTemp = gtDBEtcCode.OpenRecordset(gtSQL, dbOpenSnapshot)
'
'    '=====  신상품마스터
'    gtDBNewCommo.Close
'    sMDBFile = app.path & "\MASTER\HJ_NWCOM.mdb"
'    Set gtDBNewCommo = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'    FrmNetOff.pgbStatus.Value = 4000
'    '=====  환경설정
'    gtDBSetting.Close
'    sMDBFile = app.path & "\MASTER\HJ_SETTI.mdb"
'    Set gtDBSetting = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'    '=====  CMS마스터
'    gtDBCMS.Close
'    sMDBFile = app.path & "\MASTER\HJ_DISCU.mdb"
'    Set gtDBCMS = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'    '=====  프린터
'    gtDBPrint.Close
'    sMDBFile = app.path & "\\DATA\FRONT\NP_SETTI.mdb"
'    Set gtDBPrint = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'    FrmNetOff.pgbStatus.Value = 4500
'
'    '=====  판매시점의 고객
''    gtWsClient.Close
''    sMDBFile = app.path & "\MASTER\HJ_CLIEN.mdb"
''    Set gtWsClient = gtWS.OpenDatabase(sMDBFile)
'
'    '=====  판매시점
'    gtDBSale.Close
'    sFolder = app.path & "\DATA\SALE"
'    sMDBFile = app.path & "\DATA\SALE\" & _
'                "S" & Mid(gtSetting.sNPSName, 2, 2) & Jual_Data(Now) & ".mdb"
'    If File_Search(sFolder, sMDBFile, "S", False) Then
'        Set gtDBSale = gtWS.OpenDatabase(sMDBFile)
'    End If
'
'    If NetWork_MDB Then
'        sMDBFile = app.path & "\DATA\DAT\" & "HJ_NET" & Mid(gtSetting.sNPSName, 2, 2) & ".MDB"
'        Set gtDBNet = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'    End If
'
'        '=====  영수증 인사말 변경
'    Call Office_Load
'
'    FrmNetOff.pgbStatus.Value = 5000
'    FrmNetOff.panNetMsg = "데이터 재정의 완료"
'
'End Sub

Public Function MDB_Copy(sDay As String) As Boolean

    Dim sSaleMDB            As String
    Dim sSourceFile         As String
    Dim sFile               As String
    Dim i                   As Integer
    Dim ClsCopy             As ClsSystem
    Dim sFolder             As String
    Dim sToDay              As String
    
    MDB_Copy = False
    
    Set ClsCopy = New ClsSystem
    sToDay = Jual_Data(sDay)

    sSaleMDB = "S" & Mid(gtSetting.sNPSName, 2, 2) & sToDay & ".mdb"
    sFolder = gtsServer_Path & "\DATA\SALE\GRD\" & gtSetting.sNPSName & "\" & _
                sToDay & "\"
    '=====  서버에 판매DB 복사
    '============================================================================
    sSourceFile = App.Path & "\DATA\SALE\" & sSaleMDB

    sFile = sFolder & sSaleMDB
    
    '=====  서버에 판매DB 복사
    For i = 1 To 5
        If ClsCopy.CopyFiles(sSourceFile, sFile, False) Then
            MDB_Copy = True
            Exit For
        End If
    Next i
                
    Set ClsCopy = Nothing
    
End Function

'Private Function Status_On() As Boolean
'
'    FrmMain.wskSave.Close
'    Call FrmMain.wskSave.Connect(gtSetting.sNPSIP, Val("2345"))
'
'    Do While FrmMain.wskSave.State <> sckConnected: DoEvents
'        'wait for socket to connect
'        If FrmMain.wskSave.State = sckError Then
'            Status_On = False
'            Exit Function
'        End If
'        Status_On = True
'    Loop
'
'End Function

Public Sub HJMDB_Copy()
'=================================================
'   현재 Open 된 MDB를 닫고 복사
'   네트워크 연결시 서버DB를 복사
'=================================================
    Dim i                               As Integer
    Dim j                               As Integer
    Dim k                               As Integer
    Dim sFileName()                     As String
    Dim sSource1                        As String
    Dim sSource2                        As String
    Dim sSouForder                      As String
    Dim sDestination1                   As String
    Dim sDestination2                   As String
    Dim sMDBFile                        As String
    Dim sLockFile                       As String
    Dim bLock                           As Boolean
    Dim bCopy                           As Boolean
    Dim ClsCopy                         As ClsSystem
    
    Dim DbTemp                          As Database
    
    On Error Resume Next
    
    bCopy = False
    Set ClsCopy = New ClsSystem
    
    '===========================================================
    '   Lock파일 복사
    '===========================================================
    sLockFile = "\\" & gtsServer_Name & "\C\Windows\System\LOCK.dll"
    If ClsCopy.SearchFile(sLockFile) = False Then
        For i = 1 To 5
            If ClsCopy.CopyFiles(gtsServer_Path & "\MASTER\HJ_NPSLK.mdb", "C:\Windows\System\LOCK.dll", False) Then
                bLock = True
                Exit For
            End If
        Next i
    Else
        For i = 1 To 5
            If ClsCopy.CopyFiles(sLockFile, "C:\Windows\System\LOCK.dll", False) Then
                bLock = True
                Exit For
            End If
        Next i
    End If
    
        ReDim sFileName(9)
        
        sFileName(0) = "HJ_COMCT"           '상품코드
        sFileName(1) = "HJ_CODES"           '기타코드
        sFileName(2) = "HJ_DISCU"           '행사코드
        sFileName(3) = "HJ_MEMBE"           '사원코드
        sFileName(4) = "HJ_SETTI"           '설정
        sFileName(5) = "HJ_CLIEN"           '고객코드
        sFileName(6) = "HJ_CODEA"           '분류코드
        sFileName(7) = "HJ_NWCOM"           '신상품
        sFileName(8) = "HJ_RENTS"           '매대정보
        
    If gtbPing Then
        DoEvents
        sSouForder = "\\VSMS\VisualSMS\MASTER\NPS"
        sDestination1 = App.Path & "\TRANS"
        
        FrmMDB_Copy.panNetMsg = "마스터 복사"
        FrmAdjust.panMsgBox2 = "마스터 복사"
        FrmMDB_Copy.pgbStatus.Value = 500
        FrmAdjust.pgbStatus.Value = 500

        If ClsCopy.SearchFolder(App.Path & "\TRANS\NPS\") = False Then
            If ClsCopy.CreateFolder(App.Path & "\TRANS\NPS\") Then
                bCopy = True
            End If
        Else
            bCopy = True
        End If
        
        If bCopy Then
            'S/C에서 S/C로
            For i = 0 To 8
                sSource1 = "\\VSMS\VisualSMS\MASTER\" & sFileName(i) & ".MDB"
                sSource2 = "\\VSMS\VisualSMS\MASTER\NPS\" & sFileName(i) & ".MDB"
                                
                For k = 1 To 5
                    If ClsCopy.CopyFiles(sSource1, sSource2, False) = True Then
                        Exit For
                    End If
                Next k
            Next i
        
            For j = 1 To 3
                If ClsCopy.CopyFolder(sSouForder, sDestination1 & "\") = True Then
                    FrmMDB_Copy.pgbStatus.Value = 1000
                    FrmAdjust.pgbStatus.Value = 1000
                    '=====  파일 복사
    
                    For i = 0 To 8
                        Select Case i
                            Case 0
                                FrmMDB_Copy.panNetMsg = "상품코드 복사"
                                FrmAdjust.panMsgBox2 = "상품코드 복사"
                                FrmMDB_Copy.pgbStatus.Value = 1200
                                FrmAdjust.pgbStatus.Value = 1200
                            Case 1
                                FrmMDB_Copy.panNetMsg = "기타코드 복사"
                                FrmAdjust.panMsgBox2 = "기타코드 복사"
                                FrmMDB_Copy.pgbStatus.Value = 1400 + (100 * i) * 2
                                FrmAdjust.pgbStatus.Value = 1400 + (100 * i) * 2
                            Case 2
                                FrmMDB_Copy.panNetMsg = "행사코드 복사"
                                FrmAdjust.panMsgBox2 = "행사코드 복사"
                                FrmMDB_Copy.pgbStatus.Value = 1600 + (100 * i) * 2
                                FrmAdjust.pgbStatus.Value = 1600 + (100 * i) * 2
                            Case 3
                                FrmMDB_Copy.panNetMsg = "사원코드 복사"
                                FrmAdjust.panMsgBox2 = "사원코드 복사"
                                FrmMDB_Copy.pgbStatus.Value = 1800 + (100 * i) * 2
                                FrmAdjust.pgbStatus.Value = 1800 + (100 * i) * 2
                            Case 4
                                FrmMDB_Copy.panNetMsg = "환경설정 복사"
                                FrmAdjust.panMsgBox2 = "환경설정 복사"
                                FrmMDB_Copy.pgbStatus.Value = 2000 + (100 * i) * 2
                                FrmAdjust.pgbStatus.Value = 2000 + (100 * i) * 2
                            Case 5
                                FrmMDB_Copy.panNetMsg = "고객코드 복사"
                                FrmAdjust.panMsgBox2 = "고객코드 복사"
                                FrmMDB_Copy.pgbStatus.Value = 2200 + (100 * i) * 2
                                FrmAdjust.pgbStatus.Value = 2200 + (100 * i) * 2
                            Case 6
                                FrmMDB_Copy.panNetMsg = "분류코드 복사"
                                FrmAdjust.panMsgBox2 = "분류코드 복사"
                                FrmMDB_Copy.pgbStatus.Value = 2400 + (100 * i) * 2
                                FrmAdjust.pgbStatus.Value = 2400 + (100 * i) * 2
                            Case 7
                                FrmMDB_Copy.panNetMsg = "신상품 복사"
                                FrmAdjust.panMsgBox2 = "신상품 복사"
                                FrmMDB_Copy.pgbStatus.Value = 2600 + (100 * i) * 2
                                FrmAdjust.pgbStatus.Value = 2600 + (100 * i) * 2
                            Case 8
                                FrmMDB_Copy.panNetMsg = "매대코드 복사"
                                FrmAdjust.panMsgBox2 = "매대코드 복사"
                                FrmMDB_Copy.pgbStatus.Value = 2800 + (100 * i) * 2
                                FrmAdjust.pgbStatus.Value = 2800 + (100 * i) * 2
                        End Select
                        
                        sDestination1 = App.Path & "\TRANS\NPS\" & sFileName(i) & ".MDB"
                        sDestination2 = App.Path & "\MASTER\" & sFileName(i) & ".MDB"
                                        
                        For k = 1 To 5
                            If ClsCopy.CopyFiles(sDestination1, sDestination2, False) = True Then
                                Exit For
                            End If
                        Next k
                    Next i
                End If
            Next j
        Else
        
        End If
    End If
    
    Set ClsCopy = Nothing
    
End Sub

Public Function Client_MDB_Copy() As Boolean
'=================================================
'   현재 Open 된 MDB를 닫고 복사
'   네트워크 연결시 서버DB를 복사
'=================================================
    Dim i                               As Integer
    Dim j                               As Integer
    Dim sSource1                        As String
    Dim sSource2                        As String
    Dim sDestination1                   As String
    Dim sDestination2                   As String
    Dim sMDBFile                        As String
    Dim ClsCopy                         As ClsSystem
    
    On Error Resume Next
    
    Client_MDB_Copy = False
    Set ClsCopy = New ClsSystem
    
    '====================================================
        
        If gtbPing Then
            sSource1 = "\\VSMS\VisualSMS\MASTER\HJ_CLIEN.mdb"
            sSource2 = "\\VSMS\VisualSMS\MASTER\NPS\HJ_CLIEN.mdb"
            sDestination1 = App.Path & "\TRANS\NPS\HJ_CLIEN.MDB"
            sDestination2 = App.Path & "\MASTER\HJ_CLIEN.MDB"
            
            For j = 1 To 5
                If ClsCopy.CopyFiles(sSource1, sSource2, False) = True Then
                    If ClsCopy.CopyFiles(sSource2, sDestination1, False) = True Then
                        If ClsCopy.CopyFiles(sDestination1, sDestination2, False) = True Then
                            Client_MDB_Copy = True
                            Exit For
                        End If
                    End If
                End If
            Next j
        End If
    
    Set ClsCopy = Nothing

End Function

Public Sub NewCommo_Send()
'************************************************************
'   신상품을 스테이터스 전송
'   상품마스터로 이동
'************************************************************

    Dim DbNewCommo              As Database
    Dim ReNewCommo              As Recordset
    
    Dim i                       As Integer
    Dim iCount                  As Integer
    Dim sStatus                 As String
    
    '=====  신상품
    Set DbNewCommo = DBEngine.OpenDatabase(App.Path & "\MASTER\HJ_NWCOM.mdb")
    Set ReNewCommo = DbNewCommo.OpenRecordset("TBHJCNC", dbOpenDynaset)
    
    If ReNewCommo.RecordCount > 0 Then
        ReNewCommo.MoveLast
            iCount = ReNewCommo.RecordCount
        ReNewCommo.MoveFirst
        
        For i = 1 To iCount
            '=====  상품마스터에 등록
'            Call NewCommo_Save(ReNewCommo!TBNCCTCD, ReNewCommo!TBNCCTSP)
            
            '=====  신상품을 전송
            sStatus = NewCommo_Status(ReNewCommo!TBNCCTCD, ReNewCommo!TBNCCTSP)
            If frmMain.wskSave.State = sckConnected Then
                frmMain.wskSave.SendData sStatus
                ReNewCommo.Delete
            End If
            ReNewCommo.MoveNext
        Next i
        
'        DbNewCommo.Execute "DELETE * FROM TBHJCNC"
    End If
    
    ReNewCommo.Close
    DbNewCommo.Close
    
End Sub

Public Function NewCommo_Save(sCommoCode As String, sSellCost As String) As Boolean
'*****************************************************************************************************************************
'    Open 상품 저장
' Parameter         1.sCommoCode            :       상품코드
'                   2.sSellCost             :       판매금액
'*****************************************************************************************************************************
    
    Dim DbCommo                         As Database
    Dim ReCommo                         As Recordset            '상품 레코드
    Dim sCommoName                      As String               '상품명
    
    On Error GoTo ErrHandle
    
    NewCommo_Save = False
    Set DbCommo = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\MASTER\HJ_COMCT.mdb")
    gtSQL = "SELECT * FROM TBHJCWM WHERE TBWMCTCD='" & sCommoCode & "'"
    Set ReCommo = DbCommo.OpenRecordset(gtSQL, dbOpenDynaset)
            
    If ReCommo.RecordCount = 0 Then
        '******************************************************
        '   새로운 상품 등록
        '******************************************************
        ReCommo.AddNew
            '****** 상품코드
            ReCommo.Fields(0) = sCommoCode                              '상품코드
            '****** 상품명
                        
            ReCommo.Fields(1) = Commo_Code_Name(sCommoCode)             '상품명
            '****** 규격
            ReCommo.Fields(2) = ""                                      '규격
            '****** 분류
            ReCommo.Fields(3) = "999999"                                '분류
            '****** 원가
            ReCommo.Fields(4) = 0                                       '원가
            '****** 원가부가세
            ReCommo.Fields(5) = 0
            '****** 판매가
            ReCommo.Fields(6) = sSellCost                               '판매가
            '****** 소비자가
            ReCommo.Fields(7) = 0                                       '소비자가
            '***********    마진가      ****************************
            ReCommo.Fields(9) = Format(MaGin_Moeny(ReCommo.Fields(4), ReCommo.Fields(6), ReCommo.Fields(5)), "#,##0.00")
            '***********    마진율      ****************************
            If ReCommo.Fields(6) = 0 Then
                ReCommo.Fields(10) = 0
            Else
                ReCommo.Fields(10) = Format(MaGin_Paesent(ReCommo.Fields(4), ReCommo.Fields(6), ReCommo.Fields(5)), "#,##0.00")
            End If
            '****** 이동평균원가
            ReCommo.Fields(40) = 0                                      '이동평균원가
            '****** 현재고
            ReCommo.Fields(11) = 0                                      '현재고
            '****** 관리구분
            ReCommo.Fields(13) = 0                                      '관리구분
            '****** 등록일
            ReCommo.Fields(17) = Format(Now, "YYYY-MM-DD")              '등록일
            '****** 거래처
            ReCommo.Fields(18) = "999999"                               '거래처
            '****** 최초매입일자
            ReCommo.Fields(19) = Format(Now, "YYYY-MM-DD")
            '****** 최종매입일자
            ReCommo.Fields(20) = Format(Now, "YYYY-MM-DD")
            '****** 최초매출일자
            ReCommo.Fields(21) = Format(Now, "YYYY-MM-DD")
            '****** 최종매출일자
            ReCommo.Fields(22) = Format(Now, "YYYY-MM-DD")
            '****** 등록장소 여부
            ReCommo.Fields(33) = "1"                                '상품DB에서 오픈
        ReCommo.Update
    Else
        '******************************************************
        '   기존의 상품 수정
        '******************************************************
        ReCommo.Edit
            '****** 상품명
            ReCommo.Fields(1) = Commo_Code_Name(sCommoCode)             '상품명
            '****** 규격
            ReCommo.Fields(2) = ""                                      '규격
            '****** 분류
            ReCommo.Fields(3) = "999999"                                '분류
            '****** 원가
            ReCommo.Fields(4) = 0                                       '원가
            '****** 원가부가세
            ReCommo.Fields(5) = 0
            '****** 판매가
            ReCommo.Fields(6) = sSellCost                               '판매가
            '****** 소비자가
            ReCommo.Fields(7) = 0                                       '소비자가
            '***********    마진가      ****************************
            ReCommo.Fields(9) = Format(MaGin_Moeny(ReCommo.Fields(4), ReCommo.Fields(6), ReCommo.Fields(5)), "#,##0.00")
            '***********    마진율      ****************************
            If ReCommo.Fields(6) = 0 Then
                ReCommo.Fields(10) = 0
            Else
                ReCommo.Fields(10) = Format(MaGin_Paesent(ReCommo.Fields(4), ReCommo.Fields(6), ReCommo.Fields(5)), "#,##0.00")
            End If
            '****** 이동평균원가
            ReCommo.Fields(40) = 0                                      '이동평균원가
            '****** 현재고
            ReCommo.Fields(11) = 0                                      '현재고
            '****** 관리구분
            ReCommo.Fields(13) = 0                                      '관리구분
            '****** 등록일
            ReCommo.Fields(17) = Format(Now, "YYYY-MM-DD")              '등록일
            '****** 거래처
            ReCommo.Fields(18) = "999999"                               '거래처
            '****** 최초매입일자
            ReCommo.Fields(19) = Format(Now, "YYYY-MM-DD")
            '****** 최종매입일자
            ReCommo.Fields(20) = Format(Now, "YYYY-MM-DD")
            '****** 최초매출일자
            ReCommo.Fields(21) = Format(Now, "YYYY-MM-DD")
            '****** 최종매출일자
            ReCommo.Fields(22) = Format(Now, "YYYY-MM-DD")
            '****** 등록장소 여부
            ReCommo.Fields(33) = "1"                                '상품DB에서 오픈
        ReCommo.Update
    End If
    
    ReCommo.Close
    DbCommo.Close
    NewCommo_Save = True
        
    Exit Function
ErrHandle:
    ErrProcess ("User Function - NewCommo_Save()")
End Function

Public Function MaGin_Moeny(sOrig As String, sSellCost As String, iClass As Integer) As String
'*********************************************************************************************
'   마진가를 적용
' Parameter         1. sOrig            :       매입원가
'                   2. sSellCost        :       판매가
'                   3  iClass           :       부가세
'*********************************************************************************************

    Dim bMaGin          As Boolean
    
    On Error GoTo ErrHandle
        
        '=====  마진율 (True : 원가 대비  / False : 판가 대비)
        bMaGin = Money_Class
        
        If sOrig = 0 Then
            MaGin_Moeny = sSellCost
            Exit Function
        End If
        
        If sSellCost = 0 Then
            MaGin_Moeny = 0
            Exit Function
        End If
        
        Select Case iClass
            Case 0                  '포함
                '=====  마진가
                MaGin_Moeny = Format(sSellCost - sOrig, "#,##0.00")
            Case 1                  '별도
                '=====  마진가
                MaGin_Moeny = Format(sSellCost - (sOrig * 1.1), "#,##0.00")
            Case 2                  '면세
                '=====  마진가
                MaGin_Moeny = Format(sSellCost - sOrig, "#,##0.00")
        End Select
    Exit Function
    
ErrHandle:
    If err = 11 Then
'        txtMagin2 = "100.00"
    End If
    ErrProcess ("User Sub - MaGin_Money()")
End Function

Public Function MaGin_Paesent(sOrig As String, sSellCost As String, iClass As Integer) As String
'*********************************************************************************************
'   마진율을 적용
' Parameter         1. sOrig            :       매입원가
'                   2. sSellCost        :       판매가
'                   3. iClass           :       부가세
'*********************************************************************************************

    Dim bMaGin          As Boolean
    
    On Error GoTo ErrHandle
        
        '=====  마진율 (True : 원가 대비  / False : 판가 대비)
        bMaGin = Money_Class
        
        If sOrig = 0 Then
            MaGin_Paesent = "100.00"
            Exit Function
        End If
        
        If sSellCost = 0 Then
            MaGin_Paesent = 0
            Exit Function
        End If
    
        Select Case iClass
            Case 0                  '포함
                '=====  마진율
                If bMaGin Then
                    '=====  원가 대비
                    MaGin_Paesent = Format(((sSellCost - sOrig) / sOrig) * 100, "#,##0.00")
                Else
                    '=====  판가 대비
                    MaGin_Paesent = Format(((sSellCost - sOrig) / sSellCost) * 100, "#,##0.00")
                End If
            Case 1                  '별도
                '=====  마진율
                If bMaGin Then
                    '=====  원가 대비
                    MaGin_Paesent = Format(((sSellCost - (sOrig * 1.1)) / (sOrig * 1.1)) * 100, "#,##0.00")
                Else
                    '====== 판가 대비
                    MaGin_Paesent = Format(((sSellCost - (sOrig * 1.1)) / sSellCost) * 100, "#,##0.00")
                End If
            Case 2                  '면세
                '=====  마진율
                If bMaGin Then
                    '=====  원가 대비
                    MaGin_Paesent = Format(((sSellCost - sOrig) / sOrig) * 100, "#,##0.00")
                Else
                    '=====  판가 대비
                    MaGin_Paesent = Format(((sSellCost - sOrig) / sSellCost) * 100, "#,##0.00")
                End If
        End Select
        
    Exit Function
    
ErrHandle:
    If err = 11 Then
'        txtMagin2 = "100.00"
    End If
    ErrProcess ("User Sub - MaGin()")
End Function

Public Function Money_Class() As Boolean
'***************************************************************
'   원가 대비/판가 대비
' Return        원가대비(True)/판가대비(False)
'****************************************************************

    Dim DbTemp                          As Database
    Dim ReTemp                          As Recordset
    Dim bPing                           As Boolean
    
    bPing = gtbPing
    
'    If bPing Then
'        Set DbTemp = DBEngine.Workspaces(0).OpenDatabase("\\VSMS\VisualSMS\MASTER\HJ_SETTI")
'    Else
'        Set DbTemp = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\MASTER\HJ_SETTI")
'    End If
    
    Set DbTemp = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\MASTER\HJ_SETTI")
    Set ReTemp = DbTemp.OpenRecordset("TBHJCZ2", dbOpenSnapshot)
    
    If ReTemp.RecordCount > 0 Then
        Money_Class = ReTemp!TBZ2MOCL
    Else
        Money_Class = False
    End If
    
    ReTemp.Close
    DbTemp.Close

End Function

Public Sub ChangCommo_Send()
'************************************************************
'   가격변경 내용을 스테이터스로 전송
'************************************************************
    Dim DbChangCommo              As Database
    Dim ReChangCommo              As Recordset
    
    Dim i                       As Integer
    Dim iCount                  As Integer
    Dim sStatus                 As String
    
    '=====  신상품
    Set DbChangCommo = DBEngine.OpenDatabase(App.Path & "\DATA\FRONT\NP_HISTO.mdb")
    Set ReChangCommo = DbChangCommo.OpenRecordset("TBHJCH2", dbOpenDynaset)
    
    If ReChangCommo.RecordCount > 0 Then
        ReChangCommo.MoveLast
            iCount = ReChangCommo.RecordCount
        ReChangCommo.MoveFirst
        
        For i = 1 To iCount
            '=====  가격변경 스테이터스
            sStatus = Chang_Status(ReChangCommo!TBH2CHCD, ReChangCommo!TBH2CHPP, ReChangCommo!TBH2CHNP)
            If frmMain.wskSave.State = sckConnected Then
                frmMain.wskSave.SendData sStatus
                ReChangCommo.Delete
            End If
            ReChangCommo.MoveNext
        Next i
        
'        DbChangCommo.Execute "DELETE * FROM TBHJCH2"
    End If
    
    ReChangCommo.Close
    DbChangCommo.Close

End Sub

Public Function Chang_Status(sCommoCode As String, sOldMoney As String, sMoney As String) As String
'**************************************************************************************************
'   가격변경 스테이터스
'**************************************************************************************************
    
    Dim sSize               As String
    
    Select Case Len(sCommoCode)             '바코드 자리수
        Case 6
            sSize = "6"
            sCommoCode = sCommoCode & "ZZZZZZZ"
        Case 8
            sSize = "8"
            sCommoCode = sCommoCode & "ZZZZZ"
        Case 12
            sSize = "C"
            sCommoCode = sCommoCode & "Z"
        Case 13
            sSize = "D"
    End Select
    
    sOldMoney = Format(sOldMoney, "00000000")
    sMoney = Format(sMoney, "00000000")
                    
                    'CM + 단말기코드(3) + 판매원코드(5) + S + 상품코드(13) + 이전가격(8) + 변경가격(8) + E
    Chang_Status = "CM" & frmMain.panPOS & gtSetting.sNPSMember & sSize & sCommoCode & sOldMoney & sMoney & "E"

End Function

Public Function NewCommo_Status(sCommoCode As String, sMoney As String) As String
'********************************************************************************
'   신상품 스테이터스
'********************************************************************************

    Dim sSize               As String
    
    Select Case Len(sCommoCode)             '바코드 자리수
        Case 6
            sSize = "6"
            sCommoCode = sCommoCode & "ZZZZZZZ"
        Case 8
            sSize = "8"
            sCommoCode = sCommoCode & "ZZZZZ"
        Case 12
            sSize = "C"
            sCommoCode = sCommoCode & "Z"
        Case 13
            sSize = "D"
    End Select
    
    sMoney = Format(sMoney, "00000000")
                    
                    'NC + 단말기코드(3) + 판매원코드(5) + S + 상품코드(13) + 가격(8) + E
    NewCommo_Status = "NC" & frmMain.panPOS & gtSetting.sNPSMember & sSize & sCommoCode & sMoney & "E"

End Function

Public Sub UpData_Init(sFileName As String)
'***************************************************************
'   업데이트 데이터를 삭제
'***************************************************************
    
    Dim DbTemp                  As Database
    Dim sS_MDB                  As String
    
    sS_MDB = gtsServer_Path & "\TRANS\" & gtSetting.sNPSName & "\" & sFileName & ".mdb"
    
    Select Case sFileName
        Case "HJ_COMCT"         '상품마스터
            Set DbTemp = DBEngine.Workspaces(0).OpenDatabase(sS_MDB)
            DbTemp.Execute "DELETE * FROM TBHJCWM"
        Case "HJ_CODES"         '기타코드
            Set DbTemp = DBEngine.Workspaces(0).OpenDatabase(sS_MDB)
            DbTemp.Execute "DELETE * FROM TBHJCL3"              '박스
            DbTemp.Execute "DELETE * FROM TBHJCL4"              '묶음
            DbTemp.Execute "DELETE * FROM TBHJCL5"              '묶음내역
        Case "HJ_DISCU"         '행사코드
            Set DbTemp = DBEngine.Workspaces(0).OpenDatabase(sS_MDB)
            DbTemp.Execute "DELETE * FROM TBHJCO1"              '행사코드
            DbTemp.Execute "DELETE * FROM TBHJCO2"              '행사내역
            DbTemp.Execute "DELETE * FROM TBHJCO3"              'CMS
            DbTemp.Execute "DELETE * FROM TBHJCO4"              'CMS 내역
            DbTemp.Execute "DELETE * FROM TBHJCO5"              '반짝세일
            DbTemp.Execute "DELETE * FROM TBHJCO6"              '홈쿠폰
            DbTemp.Execute "DELETE * FROM TBHJCO7"              '홈쿠폰 내역
        Case "HJ_CLIEN"         '고객코드
            Set DbTemp = DBEngine.Workspaces(0).OpenDatabase(sS_MDB)
            DbTemp.Execute "DELETE * FROM TBHJCU1"            '고객
    End Select

    DbTemp.Close
    
End Sub

Public Sub Login_MDB()
'*******************************************************************
'   서버의 MDB를 로그인
'   1.  네트워크의 연결여부
'   2.  네트워크 드라이브의 존재여부
'       2_1. 네트워크 드라이브 생성(F <=> VisualSMS
'*******************************************************************
    
''    If Ping(True) Then
'        '=====  상품마스터
'        Set gtDbCommo = DBEngine.Workspaces(0).OpenDatabase("\\VSMS\VisualSMS\MASTER\HJ_COMCT.mdb")
'        '=====  신상품마스터
'        Set gtDbNewCommo = DBEngine.Workspaces(0).OpenDatabase("\\VSMS\VisualSMS\MASTER\HJ_NWCOM.mdb")
'
'        '=====  할인행사(CMS/행사/반짝/홈쿠폰)
'        Set gtDbSale = DBEngine.Workspaces(0).OpenDatabase("\\VSMS\VisualSMS\MASTER\HJ_DISCU.mdb")
'
'        '=====  기타코드(박스코드/묶음코드)
'        Set gtDbEtc = DBEngine.Workspaces(0).OpenDatabase("\\VSMS\VisualSMS\MASTER\HJ_CODES.mdb")
'
'        '=====  환경설정(인사말/고객디스플레이/단말기키보드/환경설정)
'        Set gtDbSetting = DBEngine.Workspaces(0).OpenDatabase("\\VSMS\VisualSMS\MASTER\HJ_SETTI.mdb")
'
'        '=====  고객
'        Set gtDbClient = DBEngine.Workspaces(0).OpenDatabase("\\VSMS\VisualSMS\MASTER\HJ_CLIEN.mdb")
''    End If

End Sub

'Public Sub MDB_Setting(bPing As Boolean)
'    Dim ECHO            As ICMP_ECHO_REPLY
'    Dim bCheck          As Boolean
'    Dim sMDBFile        As String
'    Dim sFolder         As String
'
'    On Error Resume Next
'
'    If bPing Then
'        If SocketsInitialize() Then
'             bCheck = Ping(gtSetting.sNPSIP, "", ECHO)
'             If bCheck Then
'                gtbPing = True
'                '=====  상품마스터
'                gtDbN_Commo.Close
'                sMDBFile = "\\VSMS\D\VisualSMS\MASTER\HJ_COMCT.mdb"
'                Set gtDbN_Commo = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'                '=====  신상품마스터
'                gtDbN_Commo.Close
'                sMDBFile = "\\VSMS\D\VisualSMS\MASTER\HJ_NWCOM.mdb"
'                Set gtDbN_NewCommo = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'                '=====  행사
'                gtDbN_Sale.Close
'                sMDBFile = "\\VSMS\D\VisualSMS\MASTER\HJ_DISCU.mdb"
'                Set gtDbN_Sale = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'                '=====  기타코드 마스터
'                gtDbN_Etc.Close
'                sMDBFile = "\\VSMS\D\VisualSMS\MASTER\HJ_CODES.mdb"
'                Set gtDbN_Etc = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'                '=====  환경설정
'                gtDbN_Setting.Close
'                sMDBFile = "\\VSMS\D\VisualSMS\MASTER\HJ_SETTI.mdb"
'                Set gtDbN_Setting = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'                '=====  고객마스터
'                gtDbN_Client.Close
'                sMDBFile = "\\VSMS\D\VisualSMS\MASTER\HJ_CLIEN.mdb"
'                Set gtDbN_Client = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'                Exit Sub
'            End If
'        End If
'    End If
'
'        '=====  상품마스터
'        gtDbCommo.Close
'        sMDBFile = App.Path & "\MASTER\HJ_COMCT.mdb"
'        Set gtDbCommo = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'        '=====  신상품마스터
'        gtDbNewCommo.Close
'        sMDBFile = App.Path & "\MASTER\HJ_NWCOM.mdb"
'        Set gtDbNewCommo = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'        '=====  행사
'        gtDbSale.Close
'        sMDBFile = App.Path & "\MASTER\HJ_DISCU.mdb"
'        Set gtDbSale = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'        '=====  기타코드 마스터
'        gtDbEtc.Close
'        sMDBFile = App.Path & "\MASTER\HJ_CODES.mdb"
'        Set gtDbEtc = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'        '=====  환경설정
'        gtDbSetting.Close
'        sMDBFile = App.Path & "\MASTER\HJ_SETTI.mdb"
'        Set gtDbSetting = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'        '=====  고객마스터
'        gtDbClient.Close
'        sMDBFile = App.Path & "\MASTER\HJ_CLIEN.mdb"
'        Set gtDbN_Client = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'        '=====  프린터
''        gtDBPrint.Close
''        sMDBFile = "C:\VisualNPS\\DATA\FRONT\NP_SETTI.mdb"
''        Set gtDBPrint = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'End Sub
