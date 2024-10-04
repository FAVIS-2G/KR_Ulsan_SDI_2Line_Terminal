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
    ' Ping ������
    ' RoundTripTime = Ping �� ������ �ð� ����
    ' Data = ���ϵ� ����Ÿ
    ' Address = IP �� ���� ��
    ' DataSize = Data �� ũ��
    ' Status = 0
    ' Ping ���н� �����ڵ� ����

    Dim hPort As Long
    Dim dwAddress As Long

    On Error GoTo ErrHandle
    
    ' IP ���� Long ������ ��ȯ�Ѵ�.
    dwAddress = inet_addr(sAddress)

    If dwAddress <> INADDR_NONE Then        ' ��ȿ�� IP Address �̸�
        ' ����Ʋ �����Ѵ�.
        hPort = IcmpCreateFile()

        ' ��Ʈ������ �����ϸ�
        If hPort Then
            ' Ping �õ�
            Call IcmpSendEcho(hPort, dwAddress, sDataToSend, Len(sDataToSend), _
                             0, ECHO, Len(ECHO), PING_TIMEOUT)

            ' Ping �� ���°��� �����Ѵ�.
            If ECHO.status = 0 Then
                If Abs(ECHO.Address) > 0 Then
                    Ping = True
                Else
                    Ping = False
                End If
            Else
                Ping = False
            End If
            Call IcmpCloseHandle(hPort)     ' ��Ʈ�� �ݴ´�
        End If
    Else
        ' �߸��� IP �� ���
        Ping = INADDR_NONE
    End If
    Exit Function
ErrHandle:
'    Call ErrProcess("User Function - Ping()")
End Function

Public Sub SocketsCleanup()
    If WSACleanup() <> 0 Then
        MsgBox "������ ������ CleanUp ����", vbExclamation
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
'   ��Ʈ��ũ ������� MDB ����
'******************************************
        
    Dim sDbPath                 As String
    Dim sFileName               As String
    Dim ClsSys                  As New ClsSystem
    Dim sSystemFolder           As String
    Dim i                       As Integer
    
    sFileName = "HJ_NET" & Mid(gtSetting.sNPSName, 2, 2)
    sDbPath = App.Path & "\DATA\DAT\" & sFileName & ".MDB"
    
    If ClsSys.SearchFile(sDbPath) = False Then
        '=====  �� ����
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
'   �Ǹ� �׸��� ����
'   �Ǹ� ������ ����
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
    '   ��Ʈ��ũ�� ������ ��¥�� Grd�� �Ǹ�DB�� ��� ����
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
                
                '=====  �Ǹ� DB ����
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
                            
                '=====  ���� DB ����
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
'   DB�� Open
'   ��Ʈ��ũ ����ô� ��Ʈ��ũ DB�� �α׿��ؼ� ���
' Parameter         1.bOn       :   ���Ῡ��
'=========================================================
    Dim ECHO                As ICMP_ECHO_REPLY
    Dim bCheck              As Boolean
    Dim bPing               As Boolean
    Dim i                   As Integer
    Dim sMDBFile            As String
    
    Ping_On = False
    
    On Error GoTo ErrHandle
    
    If bOn Then
        '=====  ��Ʈ��ũ ������
        FrmNet.timShow.Enabled = True
        FrmNet.pgbStatus.Value = 0
        FrmNet.panNetMsg = "��Ʈ��ũ ������"
        
        If SocketsInitialize() Then
            bCheck = Ping(gtSetting.sNPSIP, "", ECHO)
            '=====  Ping���� üũ
            If bCheck Then
                Ping_On = True
                '=====  ��ǰ�� UpData
                frmMain.timUpData.Enabled = True
                
                '====================================================
                '   ��Ʈ��ũ�� ���Ῡ�θ� Check
                '====================================================
                '=====  ������ ���� ����
                '=====  �����ϱ����� ������ Open�� MDB�� �ݱ�=>����
'                FrmNet.panNetMsg = "������ ����"
'                FrmNet.pgbStatus.Value = 1000
'                Call HJMDB_Copy

                '=====  MDB �ʱ�ȭ
                FrmNet.panNetMsg = "�Ż�ǰ ����"
                FrmNet.pgbStatus.Value = 1000
'                Call MDB_Setting(True)
                Call NewCommo_Send
                
                FrmNet.panNetMsg = "���ݺ���"
                FrmNet.pgbStatus.Value = 2000
                Call ChangCommo_Send
                
                '=====================================================================
                '   ��Ʈ��ũ�� ������ ������ DB���� �˻��� �ؼ�
                '   ������ ��¥���� ����
                '=====================================================================
                '=====  �Ǹ� ������ ������
                FrmNet.pgbStatus.Value = 3000
                FrmNet.panNetMsg = "�Ǹ� ������ ������"
                '=====  �Ǹ� ������ ����
                '=====  �Ǹ� �׸��� ����
                
                If NetWork_MDB_Load = False Then
                    MsgBox "��Ʈ��ũ�� ������ �ֽ��ϴ�." & Chr(10) & Chr(13) & _
                            "��Ʈ��ũ�� Ȯ���ϰ� �ٽ� �õ��� �ֽʽÿ�.", , "���� ����"

                    FrmNet.pgbStatus.Value = 5500
                    FrmNet.panNetMsg = "��Ʈ��ũ ������"

                    FrmNetOff.panNetMsg = "��Ʈ��ũ ������"
                    FrmNetOff.pgbStatus.Value = 1000
                    FrmNetOff.pgbStatus.Value = 1500

    '                Call MDB_Setting(False)
                    gtbPing = False
                    gtbNetOn = False
                    Ping_On = False
                    Exit Function
                End If
                
                FrmNet.pgbStatus.Value = 3000
'                FrmNet.panNetMsg = "������ ������Ʈ��"
                FrmNet.panNetMsg = "������ ������"
'                Call Login_MDB
                
                gtbPing = True
                Ping_On = True
                FrmNet.pgbStatus.Value = 5000
            Else
                FrmNet.pgbStatus.Value = 5500
                FrmNet.panNetMsg = "��Ʈ��ũ ������"
                
                FrmNetOff.panNetMsg = "��Ʈ��ũ ������"
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
        FrmNetOff.panNetMsg = "��Ʈ��ũ ������"
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
'   DB�� Open
'   ��Ʈ��ũ ����ô� MDB�� ����
'=========================================================
    Dim ECHO                As ICMP_ECHO_REPLY
    Dim bCheck              As Boolean
    Dim bPing               As Boolean
    Dim i                   As Integer
    Dim sMDBFile            As String

    '=====  ��Ʈ��ũ ������
    FrmMDB_Copy.timShow.Enabled = True
    FrmMDB_Copy.pgbStatus.Value = 0
    FrmMDB_Copy.panNetMsg = "��Ʈ��ũ ������"
    
    If SocketsInitialize() Then
        bCheck = Ping(gtSetting.sNPSIP, "", ECHO)
        '=====  Ping���� üũ
        If bCheck Then
            '=====  ������ ���� ����
            '=====  �����ϱ����� ������ Open�� MDB�� �ݱ�=>����
            FrmMDB_Copy.panNetMsg = "������ ����"
            FrmMDB_Copy.pgbStatus.Value = 1000
            gtbPing = True
            Call HJMDB_Copy
            
            '=====  MDB �ʱ�ȭ
            FrmMDB_Copy.panNetMsg = "������ ������"
            FrmMDB_Copy.pgbStatus.Value = 3000
'            Call MDB_Setting(True)
            
            FrmMDB_Copy.pgbStatus.Value = 5000
        Else
            FrmMDB_Copy.panNetMsg = "��Ʈ��ũ ������"
            FrmMDB_Copy.pgbStatus.Value = 5000
        End If
    End If
    
End Function

Public Sub Net_Check()
'*******************************************************
'   ������ ���¿��� ��Ʈ��ũ�� �˻��ؼ� ����
'*******************************************************
    Dim ECHO                As ICMP_ECHO_REPLY
    Dim bCheck              As Long

    On Error GoTo ErrHandle
    
    If SocketsInitialize() Then
        bCheck = Ping(gtSetting.sNPSIP, "", ECHO)
        '=====  Ping���� üũ
        If bCheck Then
            gtbNetOn = True
            FrmNet.Show 1
        Else
            MsgBox "��Ʈ��ũ ���¿� ������ �ֽ��ϴ�.", , "��Ʈ��ũ"
            gtbPing = False
            gtbNetOn = False
        End If
    Else
        MsgBox "��Ʈ��ũ ���¿� ������ �ֽ��ϴ�.", , "��Ʈ��ũ"
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
'                '=====  ��ǰ������
'                gtDB_N_Commo.Close
'                sMDBFile = gtsServer_Path & "\MASTER\HJ_COMCT.mdb"
'                Set gtDB_N_Commo = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'                '=====  ��������
'                gtDB_N_Client.Close
'                sMDBFile = gtsServer_Path & "\MASTER\HJ_CLIEN.mdb"
'                Set gtDB_N_Client = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'                '=====  ��Ÿ�ڵ� ������
'                gtDB_N_EtcCode.Close
'                sMDBFile = gtsServer_Path & "\MASTER\HJ_CODES.mdb"
'                Set gtDB_N_EtcCode = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'                '=====  �Ż�ǰ������
'                gtDB_N_NewCommo.Close
'                sMDBFile = gtsServer_Path & "\MASTER\HJ_NWCOM.mdb"
'                Set gtDB_N_NewCommo = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'                '=====  ȯ�漳��
'                gtDB_N_Setting.Close
'                sMDBFile = gtsServer_Path & "\MASTER\HJ_SETTI.mdb"
'                Set gtDB_N_Setting = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'                '=====  CMS������
'                gtDB_N_CMS.Close
'                sMDBFile = gtsServer_Path & "\MASTER\HJ_DISCU.mdb"
'                Set gtDB_N_CMS = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'                Exit Sub
'            End If
'        End If
'    End If
'
'    '=====  ��ǰ������
'    FrmNetOff.panNetMsg = "������ ������"
'    FrmNetOff.pgbStatus.Value = 2000
'
'    gtDBCommo.Close
'    sMDBFile = app.path & "\MASTER\HJ_COMCT.mdb"
'    Set gtDBCommo = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'    FrmNetOff.pgbStatus.Value = 3000
'
'    '=====  ��������
'    gtDBClient.Close
'    sMDBFile = app.path & "\MASTER\HJ_CLIEN.mdb"
'    Set gtDBClient = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'    FrmNetOff.pgbStatus.Value = 3500
'
'    FrmNetOff.panNetMsg = "������ ������ ��"
'    '=====  ��Ÿ�ڵ� ������
'    gtDBEtcCode.Close
'    sMDBFile = app.path & "\MASTER\HJ_CODES.mdb"
'    Set gtDBEtcCode = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'    gtSQL = "SELECT TBL7TPCD FROM TBHJCL7 WHERE TBL7CTCD='" & 88006611 & "'"
'    Set gtReTemp = gtDBEtcCode.OpenRecordset(gtSQL, dbOpenSnapshot)
'
'    '=====  �Ż�ǰ������
'    gtDBNewCommo.Close
'    sMDBFile = app.path & "\MASTER\HJ_NWCOM.mdb"
'    Set gtDBNewCommo = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'    FrmNetOff.pgbStatus.Value = 4000
'    '=====  ȯ�漳��
'    gtDBSetting.Close
'    sMDBFile = app.path & "\MASTER\HJ_SETTI.mdb"
'    Set gtDBSetting = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'    '=====  CMS������
'    gtDBCMS.Close
'    sMDBFile = app.path & "\MASTER\HJ_DISCU.mdb"
'    Set gtDBCMS = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'    '=====  ������
'    gtDBPrint.Close
'    sMDBFile = app.path & "\\DATA\FRONT\NP_SETTI.mdb"
'    Set gtDBPrint = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'    FrmNetOff.pgbStatus.Value = 4500
'
'    '=====  �ǸŽ����� ��
''    gtWsClient.Close
''    sMDBFile = app.path & "\MASTER\HJ_CLIEN.mdb"
''    Set gtWsClient = gtWS.OpenDatabase(sMDBFile)
'
'    '=====  �ǸŽ���
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
'        '=====  ������ �λ縻 ����
'    Call Office_Load
'
'    FrmNetOff.pgbStatus.Value = 5000
'    FrmNetOff.panNetMsg = "������ ������ �Ϸ�"
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
    '=====  ������ �Ǹ�DB ����
    '============================================================================
    sSourceFile = App.Path & "\DATA\SALE\" & sSaleMDB

    sFile = sFolder & sSaleMDB
    
    '=====  ������ �Ǹ�DB ����
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
'   ���� Open �� MDB�� �ݰ� ����
'   ��Ʈ��ũ ����� ����DB�� ����
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
    '   Lock���� ����
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
        
        sFileName(0) = "HJ_COMCT"           '��ǰ�ڵ�
        sFileName(1) = "HJ_CODES"           '��Ÿ�ڵ�
        sFileName(2) = "HJ_DISCU"           '����ڵ�
        sFileName(3) = "HJ_MEMBE"           '����ڵ�
        sFileName(4) = "HJ_SETTI"           '����
        sFileName(5) = "HJ_CLIEN"           '���ڵ�
        sFileName(6) = "HJ_CODEA"           '�з��ڵ�
        sFileName(7) = "HJ_NWCOM"           '�Ż�ǰ
        sFileName(8) = "HJ_RENTS"           '�Ŵ�����
        
    If gtbPing Then
        DoEvents
        sSouForder = "\\VSMS\VisualSMS\MASTER\NPS"
        sDestination1 = App.Path & "\TRANS"
        
        FrmMDB_Copy.panNetMsg = "������ ����"
        FrmAdjust.panMsgBox2 = "������ ����"
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
            'S/C���� S/C��
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
                    '=====  ���� ����
    
                    For i = 0 To 8
                        Select Case i
                            Case 0
                                FrmMDB_Copy.panNetMsg = "��ǰ�ڵ� ����"
                                FrmAdjust.panMsgBox2 = "��ǰ�ڵ� ����"
                                FrmMDB_Copy.pgbStatus.Value = 1200
                                FrmAdjust.pgbStatus.Value = 1200
                            Case 1
                                FrmMDB_Copy.panNetMsg = "��Ÿ�ڵ� ����"
                                FrmAdjust.panMsgBox2 = "��Ÿ�ڵ� ����"
                                FrmMDB_Copy.pgbStatus.Value = 1400 + (100 * i) * 2
                                FrmAdjust.pgbStatus.Value = 1400 + (100 * i) * 2
                            Case 2
                                FrmMDB_Copy.panNetMsg = "����ڵ� ����"
                                FrmAdjust.panMsgBox2 = "����ڵ� ����"
                                FrmMDB_Copy.pgbStatus.Value = 1600 + (100 * i) * 2
                                FrmAdjust.pgbStatus.Value = 1600 + (100 * i) * 2
                            Case 3
                                FrmMDB_Copy.panNetMsg = "����ڵ� ����"
                                FrmAdjust.panMsgBox2 = "����ڵ� ����"
                                FrmMDB_Copy.pgbStatus.Value = 1800 + (100 * i) * 2
                                FrmAdjust.pgbStatus.Value = 1800 + (100 * i) * 2
                            Case 4
                                FrmMDB_Copy.panNetMsg = "ȯ�漳�� ����"
                                FrmAdjust.panMsgBox2 = "ȯ�漳�� ����"
                                FrmMDB_Copy.pgbStatus.Value = 2000 + (100 * i) * 2
                                FrmAdjust.pgbStatus.Value = 2000 + (100 * i) * 2
                            Case 5
                                FrmMDB_Copy.panNetMsg = "���ڵ� ����"
                                FrmAdjust.panMsgBox2 = "���ڵ� ����"
                                FrmMDB_Copy.pgbStatus.Value = 2200 + (100 * i) * 2
                                FrmAdjust.pgbStatus.Value = 2200 + (100 * i) * 2
                            Case 6
                                FrmMDB_Copy.panNetMsg = "�з��ڵ� ����"
                                FrmAdjust.panMsgBox2 = "�з��ڵ� ����"
                                FrmMDB_Copy.pgbStatus.Value = 2400 + (100 * i) * 2
                                FrmAdjust.pgbStatus.Value = 2400 + (100 * i) * 2
                            Case 7
                                FrmMDB_Copy.panNetMsg = "�Ż�ǰ ����"
                                FrmAdjust.panMsgBox2 = "�Ż�ǰ ����"
                                FrmMDB_Copy.pgbStatus.Value = 2600 + (100 * i) * 2
                                FrmAdjust.pgbStatus.Value = 2600 + (100 * i) * 2
                            Case 8
                                FrmMDB_Copy.panNetMsg = "�Ŵ��ڵ� ����"
                                FrmAdjust.panMsgBox2 = "�Ŵ��ڵ� ����"
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
'   ���� Open �� MDB�� �ݰ� ����
'   ��Ʈ��ũ ����� ����DB�� ����
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
'   �Ż�ǰ�� �������ͽ� ����
'   ��ǰ�����ͷ� �̵�
'************************************************************

    Dim DbNewCommo              As Database
    Dim ReNewCommo              As Recordset
    
    Dim i                       As Integer
    Dim iCount                  As Integer
    Dim sStatus                 As String
    
    '=====  �Ż�ǰ
    Set DbNewCommo = DBEngine.OpenDatabase(App.Path & "\MASTER\HJ_NWCOM.mdb")
    Set ReNewCommo = DbNewCommo.OpenRecordset("TBHJCNC", dbOpenDynaset)
    
    If ReNewCommo.RecordCount > 0 Then
        ReNewCommo.MoveLast
            iCount = ReNewCommo.RecordCount
        ReNewCommo.MoveFirst
        
        For i = 1 To iCount
            '=====  ��ǰ�����Ϳ� ���
'            Call NewCommo_Save(ReNewCommo!TBNCCTCD, ReNewCommo!TBNCCTSP)
            
            '=====  �Ż�ǰ�� ����
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
'    Open ��ǰ ����
' Parameter         1.sCommoCode            :       ��ǰ�ڵ�
'                   2.sSellCost             :       �Ǹűݾ�
'*****************************************************************************************************************************
    
    Dim DbCommo                         As Database
    Dim ReCommo                         As Recordset            '��ǰ ���ڵ�
    Dim sCommoName                      As String               '��ǰ��
    
    On Error GoTo ErrHandle
    
    NewCommo_Save = False
    Set DbCommo = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\MASTER\HJ_COMCT.mdb")
    gtSQL = "SELECT * FROM TBHJCWM WHERE TBWMCTCD='" & sCommoCode & "'"
    Set ReCommo = DbCommo.OpenRecordset(gtSQL, dbOpenDynaset)
            
    If ReCommo.RecordCount = 0 Then
        '******************************************************
        '   ���ο� ��ǰ ���
        '******************************************************
        ReCommo.AddNew
            '****** ��ǰ�ڵ�
            ReCommo.Fields(0) = sCommoCode                              '��ǰ�ڵ�
            '****** ��ǰ��
                        
            ReCommo.Fields(1) = Commo_Code_Name(sCommoCode)             '��ǰ��
            '****** �԰�
            ReCommo.Fields(2) = ""                                      '�԰�
            '****** �з�
            ReCommo.Fields(3) = "999999"                                '�з�
            '****** ����
            ReCommo.Fields(4) = 0                                       '����
            '****** �����ΰ���
            ReCommo.Fields(5) = 0
            '****** �ǸŰ�
            ReCommo.Fields(6) = sSellCost                               '�ǸŰ�
            '****** �Һ��ڰ�
            ReCommo.Fields(7) = 0                                       '�Һ��ڰ�
            '***********    ������      ****************************
            ReCommo.Fields(9) = Format(MaGin_Moeny(ReCommo.Fields(4), ReCommo.Fields(6), ReCommo.Fields(5)), "#,##0.00")
            '***********    ������      ****************************
            If ReCommo.Fields(6) = 0 Then
                ReCommo.Fields(10) = 0
            Else
                ReCommo.Fields(10) = Format(MaGin_Paesent(ReCommo.Fields(4), ReCommo.Fields(6), ReCommo.Fields(5)), "#,##0.00")
            End If
            '****** �̵���տ���
            ReCommo.Fields(40) = 0                                      '�̵���տ���
            '****** �����
            ReCommo.Fields(11) = 0                                      '�����
            '****** ��������
            ReCommo.Fields(13) = 0                                      '��������
            '****** �����
            ReCommo.Fields(17) = Format(Now, "YYYY-MM-DD")              '�����
            '****** �ŷ�ó
            ReCommo.Fields(18) = "999999"                               '�ŷ�ó
            '****** ���ʸ�������
            ReCommo.Fields(19) = Format(Now, "YYYY-MM-DD")
            '****** ������������
            ReCommo.Fields(20) = Format(Now, "YYYY-MM-DD")
            '****** ���ʸ�������
            ReCommo.Fields(21) = Format(Now, "YYYY-MM-DD")
            '****** ������������
            ReCommo.Fields(22) = Format(Now, "YYYY-MM-DD")
            '****** ������ ����
            ReCommo.Fields(33) = "1"                                '��ǰDB���� ����
        ReCommo.Update
    Else
        '******************************************************
        '   ������ ��ǰ ����
        '******************************************************
        ReCommo.Edit
            '****** ��ǰ��
            ReCommo.Fields(1) = Commo_Code_Name(sCommoCode)             '��ǰ��
            '****** �԰�
            ReCommo.Fields(2) = ""                                      '�԰�
            '****** �з�
            ReCommo.Fields(3) = "999999"                                '�з�
            '****** ����
            ReCommo.Fields(4) = 0                                       '����
            '****** �����ΰ���
            ReCommo.Fields(5) = 0
            '****** �ǸŰ�
            ReCommo.Fields(6) = sSellCost                               '�ǸŰ�
            '****** �Һ��ڰ�
            ReCommo.Fields(7) = 0                                       '�Һ��ڰ�
            '***********    ������      ****************************
            ReCommo.Fields(9) = Format(MaGin_Moeny(ReCommo.Fields(4), ReCommo.Fields(6), ReCommo.Fields(5)), "#,##0.00")
            '***********    ������      ****************************
            If ReCommo.Fields(6) = 0 Then
                ReCommo.Fields(10) = 0
            Else
                ReCommo.Fields(10) = Format(MaGin_Paesent(ReCommo.Fields(4), ReCommo.Fields(6), ReCommo.Fields(5)), "#,##0.00")
            End If
            '****** �̵���տ���
            ReCommo.Fields(40) = 0                                      '�̵���տ���
            '****** �����
            ReCommo.Fields(11) = 0                                      '�����
            '****** ��������
            ReCommo.Fields(13) = 0                                      '��������
            '****** �����
            ReCommo.Fields(17) = Format(Now, "YYYY-MM-DD")              '�����
            '****** �ŷ�ó
            ReCommo.Fields(18) = "999999"                               '�ŷ�ó
            '****** ���ʸ�������
            ReCommo.Fields(19) = Format(Now, "YYYY-MM-DD")
            '****** ������������
            ReCommo.Fields(20) = Format(Now, "YYYY-MM-DD")
            '****** ���ʸ�������
            ReCommo.Fields(21) = Format(Now, "YYYY-MM-DD")
            '****** ������������
            ReCommo.Fields(22) = Format(Now, "YYYY-MM-DD")
            '****** ������ ����
            ReCommo.Fields(33) = "1"                                '��ǰDB���� ����
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
'   �������� ����
' Parameter         1. sOrig            :       ���Կ���
'                   2. sSellCost        :       �ǸŰ�
'                   3  iClass           :       �ΰ���
'*********************************************************************************************

    Dim bMaGin          As Boolean
    
    On Error GoTo ErrHandle
        
        '=====  ������ (True : ���� ���  / False : �ǰ� ���)
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
            Case 0                  '����
                '=====  ������
                MaGin_Moeny = Format(sSellCost - sOrig, "#,##0.00")
            Case 1                  '����
                '=====  ������
                MaGin_Moeny = Format(sSellCost - (sOrig * 1.1), "#,##0.00")
            Case 2                  '�鼼
                '=====  ������
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
'   �������� ����
' Parameter         1. sOrig            :       ���Կ���
'                   2. sSellCost        :       �ǸŰ�
'                   3. iClass           :       �ΰ���
'*********************************************************************************************

    Dim bMaGin          As Boolean
    
    On Error GoTo ErrHandle
        
        '=====  ������ (True : ���� ���  / False : �ǰ� ���)
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
            Case 0                  '����
                '=====  ������
                If bMaGin Then
                    '=====  ���� ���
                    MaGin_Paesent = Format(((sSellCost - sOrig) / sOrig) * 100, "#,##0.00")
                Else
                    '=====  �ǰ� ���
                    MaGin_Paesent = Format(((sSellCost - sOrig) / sSellCost) * 100, "#,##0.00")
                End If
            Case 1                  '����
                '=====  ������
                If bMaGin Then
                    '=====  ���� ���
                    MaGin_Paesent = Format(((sSellCost - (sOrig * 1.1)) / (sOrig * 1.1)) * 100, "#,##0.00")
                Else
                    '====== �ǰ� ���
                    MaGin_Paesent = Format(((sSellCost - (sOrig * 1.1)) / sSellCost) * 100, "#,##0.00")
                End If
            Case 2                  '�鼼
                '=====  ������
                If bMaGin Then
                    '=====  ���� ���
                    MaGin_Paesent = Format(((sSellCost - sOrig) / sOrig) * 100, "#,##0.00")
                Else
                    '=====  �ǰ� ���
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
'   ���� ���/�ǰ� ���
' Return        �������(True)/�ǰ����(False)
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
'   ���ݺ��� ������ �������ͽ��� ����
'************************************************************
    Dim DbChangCommo              As Database
    Dim ReChangCommo              As Recordset
    
    Dim i                       As Integer
    Dim iCount                  As Integer
    Dim sStatus                 As String
    
    '=====  �Ż�ǰ
    Set DbChangCommo = DBEngine.OpenDatabase(App.Path & "\DATA\FRONT\NP_HISTO.mdb")
    Set ReChangCommo = DbChangCommo.OpenRecordset("TBHJCH2", dbOpenDynaset)
    
    If ReChangCommo.RecordCount > 0 Then
        ReChangCommo.MoveLast
            iCount = ReChangCommo.RecordCount
        ReChangCommo.MoveFirst
        
        For i = 1 To iCount
            '=====  ���ݺ��� �������ͽ�
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
'   ���ݺ��� �������ͽ�
'**************************************************************************************************
    
    Dim sSize               As String
    
    Select Case Len(sCommoCode)             '���ڵ� �ڸ���
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
                    
                    'CM + �ܸ����ڵ�(3) + �Ǹſ��ڵ�(5) + S + ��ǰ�ڵ�(13) + ��������(8) + ���氡��(8) + E
    Chang_Status = "CM" & frmMain.panPOS & gtSetting.sNPSMember & sSize & sCommoCode & sOldMoney & sMoney & "E"

End Function

Public Function NewCommo_Status(sCommoCode As String, sMoney As String) As String
'********************************************************************************
'   �Ż�ǰ �������ͽ�
'********************************************************************************

    Dim sSize               As String
    
    Select Case Len(sCommoCode)             '���ڵ� �ڸ���
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
                    
                    'NC + �ܸ����ڵ�(3) + �Ǹſ��ڵ�(5) + S + ��ǰ�ڵ�(13) + ����(8) + E
    NewCommo_Status = "NC" & frmMain.panPOS & gtSetting.sNPSMember & sSize & sCommoCode & sMoney & "E"

End Function

Public Sub UpData_Init(sFileName As String)
'***************************************************************
'   ������Ʈ �����͸� ����
'***************************************************************
    
    Dim DbTemp                  As Database
    Dim sS_MDB                  As String
    
    sS_MDB = gtsServer_Path & "\TRANS\" & gtSetting.sNPSName & "\" & sFileName & ".mdb"
    
    Select Case sFileName
        Case "HJ_COMCT"         '��ǰ������
            Set DbTemp = DBEngine.Workspaces(0).OpenDatabase(sS_MDB)
            DbTemp.Execute "DELETE * FROM TBHJCWM"
        Case "HJ_CODES"         '��Ÿ�ڵ�
            Set DbTemp = DBEngine.Workspaces(0).OpenDatabase(sS_MDB)
            DbTemp.Execute "DELETE * FROM TBHJCL3"              '�ڽ�
            DbTemp.Execute "DELETE * FROM TBHJCL4"              '����
            DbTemp.Execute "DELETE * FROM TBHJCL5"              '��������
        Case "HJ_DISCU"         '����ڵ�
            Set DbTemp = DBEngine.Workspaces(0).OpenDatabase(sS_MDB)
            DbTemp.Execute "DELETE * FROM TBHJCO1"              '����ڵ�
            DbTemp.Execute "DELETE * FROM TBHJCO2"              '��系��
            DbTemp.Execute "DELETE * FROM TBHJCO3"              'CMS
            DbTemp.Execute "DELETE * FROM TBHJCO4"              'CMS ����
            DbTemp.Execute "DELETE * FROM TBHJCO5"              '��¦����
            DbTemp.Execute "DELETE * FROM TBHJCO6"              'Ȩ����
            DbTemp.Execute "DELETE * FROM TBHJCO7"              'Ȩ���� ����
        Case "HJ_CLIEN"         '���ڵ�
            Set DbTemp = DBEngine.Workspaces(0).OpenDatabase(sS_MDB)
            DbTemp.Execute "DELETE * FROM TBHJCU1"            '��
    End Select

    DbTemp.Close
    
End Sub

Public Sub Login_MDB()
'*******************************************************************
'   ������ MDB�� �α���
'   1.  ��Ʈ��ũ�� ���Ῡ��
'   2.  ��Ʈ��ũ ����̺��� ���翩��
'       2_1. ��Ʈ��ũ ����̺� ����(F <=> VisualSMS
'*******************************************************************
    
''    If Ping(True) Then
'        '=====  ��ǰ������
'        Set gtDbCommo = DBEngine.Workspaces(0).OpenDatabase("\\VSMS\VisualSMS\MASTER\HJ_COMCT.mdb")
'        '=====  �Ż�ǰ������
'        Set gtDbNewCommo = DBEngine.Workspaces(0).OpenDatabase("\\VSMS\VisualSMS\MASTER\HJ_NWCOM.mdb")
'
'        '=====  �������(CMS/���/��¦/Ȩ����)
'        Set gtDbSale = DBEngine.Workspaces(0).OpenDatabase("\\VSMS\VisualSMS\MASTER\HJ_DISCU.mdb")
'
'        '=====  ��Ÿ�ڵ�(�ڽ��ڵ�/�����ڵ�)
'        Set gtDbEtc = DBEngine.Workspaces(0).OpenDatabase("\\VSMS\VisualSMS\MASTER\HJ_CODES.mdb")
'
'        '=====  ȯ�漳��(�λ縻/�����÷���/�ܸ���Ű����/ȯ�漳��)
'        Set gtDbSetting = DBEngine.Workspaces(0).OpenDatabase("\\VSMS\VisualSMS\MASTER\HJ_SETTI.mdb")
'
'        '=====  ��
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
'                '=====  ��ǰ������
'                gtDbN_Commo.Close
'                sMDBFile = "\\VSMS\D\VisualSMS\MASTER\HJ_COMCT.mdb"
'                Set gtDbN_Commo = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'                '=====  �Ż�ǰ������
'                gtDbN_Commo.Close
'                sMDBFile = "\\VSMS\D\VisualSMS\MASTER\HJ_NWCOM.mdb"
'                Set gtDbN_NewCommo = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'                '=====  ���
'                gtDbN_Sale.Close
'                sMDBFile = "\\VSMS\D\VisualSMS\MASTER\HJ_DISCU.mdb"
'                Set gtDbN_Sale = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'                '=====  ��Ÿ�ڵ� ������
'                gtDbN_Etc.Close
'                sMDBFile = "\\VSMS\D\VisualSMS\MASTER\HJ_CODES.mdb"
'                Set gtDbN_Etc = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'                '=====  ȯ�漳��
'                gtDbN_Setting.Close
'                sMDBFile = "\\VSMS\D\VisualSMS\MASTER\HJ_SETTI.mdb"
'                Set gtDbN_Setting = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'                '=====  ��������
'                gtDbN_Client.Close
'                sMDBFile = "\\VSMS\D\VisualSMS\MASTER\HJ_CLIEN.mdb"
'                Set gtDbN_Client = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'                Exit Sub
'            End If
'        End If
'    End If
'
'        '=====  ��ǰ������
'        gtDbCommo.Close
'        sMDBFile = App.Path & "\MASTER\HJ_COMCT.mdb"
'        Set gtDbCommo = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'        '=====  �Ż�ǰ������
'        gtDbNewCommo.Close
'        sMDBFile = App.Path & "\MASTER\HJ_NWCOM.mdb"
'        Set gtDbNewCommo = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'        '=====  ���
'        gtDbSale.Close
'        sMDBFile = App.Path & "\MASTER\HJ_DISCU.mdb"
'        Set gtDbSale = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'        '=====  ��Ÿ�ڵ� ������
'        gtDbEtc.Close
'        sMDBFile = App.Path & "\MASTER\HJ_CODES.mdb"
'        Set gtDbEtc = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'        '=====  ȯ�漳��
'        gtDbSetting.Close
'        sMDBFile = App.Path & "\MASTER\HJ_SETTI.mdb"
'        Set gtDbSetting = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'        '=====  ��������
'        gtDbClient.Close
'        sMDBFile = App.Path & "\MASTER\HJ_CLIEN.mdb"
'        Set gtDbN_Client = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'        '=====  ������
''        gtDBPrint.Close
''        sMDBFile = "C:\VisualNPS\\DATA\FRONT\NP_SETTI.mdb"
''        Set gtDBPrint = DBEngine.Workspaces(0).OpenDatabase(sMDBFile)
'
'End Sub
