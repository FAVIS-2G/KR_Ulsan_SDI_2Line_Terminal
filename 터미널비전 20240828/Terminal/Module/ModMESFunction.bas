Attribute VB_Name = "ModMESFunction"
'/***********************************************************
' ��Ʈ�p ���۰��� ���
'/***********************************************************
Public Type NETRESOURCE
   dwScope       As Long
   dwType        As Long
   dwDisplayType As Long
   dwUsage       As Long
   lpLocalName   As String
   lpRemoteName  As String
   lpComment     As String
   lpProvider    As String
End Type
Public Declare Function WNetOpenEnum Lib "mpr.dll" _
Alias "WNetOpenEnumA" _
(ByVal dwScope As Long, _
ByVal dwType As Long, _
ByVal dwUsage As Long, _
lpNetResource As NETRESOURCE, _
lphEnum As Long) As Long

Public Declare Function WNetAddConnection2 Lib "mpr" _
    Alias "WNetAddConnection2A" _
   (lpNetResource As NETRESOURCE, _
    ByVal lpPassword As String, _
    ByVal lpUserName As String, _
    ByVal dwFlags As Long) As Long
       
Public Declare Function WNetCancelConnection2 Lib "mpr" _
    Alias "WNetCancelConnection2A" _
   (ByVal lpName As String, _
    ByVal dwFlags As Long, _
    ByVal fForce As Long) As Long
Public Declare Function WNetGetConnection Lib "mpr.dll" _
    Alias "WNetGetConnectionA" _
    (ByVal lpszLocalName As String, _
    ByVal lpszRemoteName As String, _
    cbRemoteName As Long) As Long
Public Declare Function IsNetworkAlive Lib "Sensapi.dll" (dwFlags As Long) As Long  '�����̺� ������� üũ
Public Declare Function WNetConnectionDialog Lib "mpr" _
   (ByVal hWnd As Long, ByVal dwType As Long) As Long
Public Declare Function WNetDisconnectDialog Lib "mpr" _
   (ByVal hWnd As Long, ByVal dwType As Long) As Long
Public Declare Function WNetCancelConnection Lib "mpr.dll" Alias "WNetCancelConnectionA" (ByVal lpszName As String, ByVal bForce As Long) As Long

'���׽�Ʈ - �ϴ��ϴ� ���� ���ϳ�
Public Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Public Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Public Declare Function IcmpSendEcho Lib "icmp.dll" _
       (ByVal IcmpHandle As Long, _
        ByVal DestinationAddress As Long, _
        ByVal RequestData As String, _
        ByVal RequestSize As Long, _
        ByVal RequestOptions As Long, _
        ReplyBuffer As ICMP_ECHO_REPLY, _
        ByVal ReplySize As Long, _
        ByVal Timeout As Long) As Long
Public Declare Function inet_addr Lib "WSOCK32.DLL" (ByVal s As String) As Long    '�����Ǹ� ������ �ٲ��ֳ�..? -_-;
Public Const PING_TIMEOUT As Long = 500
Public Type ICMP_OPTIONS
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

Public Const IP_SUCCESS As Long = 0
Public Const IP_STATUS_BASE As Long = 11000
Public Const IP_BUF_TOO_SMALL As Long = (11000 + 1)
Public Const IP_DEST_NET_UNREACHABLE As Long = (11000 + 2)
Public Const IP_DEST_HOST_UNREACHABLE As Long = (11000 + 3)
Public Const IP_DEST_PROT_UNREACHABLE As Long = (11000 + 4)
Public Const IP_DEST_PORT_UNREACHABLE As Long = (11000 + 5)
Public Const IP_NO_RESOURCES As Long = (11000 + 6)
Public Const IP_BAD_OPTION As Long = (11000 + 7)
Public Const IP_HW_ERROR As Long = (11000 + 8)
Public Const IP_PACKET_TOO_BIG As Long = (11000 + 9)
Public Const IP_REQ_TIMED_OUT As Long = (11000 + 10)
Public Const IP_BAD_REQ As Long = (11000 + 11)
Public Const IP_BAD_ROUTE As Long = (11000 + 12)
Public Const IP_TTL_EXPIRED_TRANSIT As Long = (11000 + 13)
Public Const IP_TTL_EXPIRED_REASSEM As Long = (11000 + 14)
Public Const IP_PARAM_PROBLEM As Long = (11000 + 15)
Public Const IP_SOURCE_QUENCH As Long = (11000 + 16)
Public Const IP_OPTION_TOO_BIG As Long = (11000 + 17)
Public Const IP_BAD_DESTINATION As Long = (11000 + 18)
Public Const IP_ADDR_DELETED As Long = (11000 + 19)
Public Const IP_SPEC_MTU_CHANGE As Long = (11000 + 20)
Public Const IP_MTU_CHANGE As Long = (11000 + 21)
Public Const IP_UNLOAD As Long = (11000 + 22)
Public Const IP_ADDR_ADDED As Long = (11000 + 23)
Public Const IP_GENERAL_FAILURE As Long = (11000 + 50)
Public Const MAX_IP_STATUS As Long = (11000 + 50)
Public Const IP_PENDING As Long = (11000 + 255)
Public Const WS_VERSION_REQD As Long = &H101
Public Const MIN_SOCKETS_REQD As Long = 1
Public Const SOCKET_ERROR As Long = -1
Public Const MAX_WSADescription As Long = 256
Public Const MAX_WSASYSStatus As Long = 128

Public Const INADDR_NONE As Long = &HFFFFFFFF
Public Const ERROR_SUCCESS = 0
Public Const CONNECT_UPDATE_PROFILE = &H1
Public Const RESOURCETYPE_DISK = &H1
Public Const RESOURCETYPE_PRINT = &H2
Public Const RESOURCETYPE_ANY = &H0
Public Const RESOURCE_GLOBALNET = &H2
Public Const RESOURCEDISPLAYTYPE_SHARE = &H3
Public Const RESOURCEUSAGE_CONNECTABLE = &H1

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function SHFileOperation Lib "Shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type
Public Const FO_COPY = &H2
Public Const FOF_ALLOWUNDO = &H40

Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" _
              (ByVal lpApplicationName As Long, _
               ByVal lpCommandLine As String, _
               ByVal lpProcessAttributes As Long, _
               ByVal lpThreadAttributes As Long, _
               ByVal bInheritHandles As Long, _
               ByVal dwCreationFlags As Long, _
               ByVal lpEnvironment As Long, _
               ByVal lpCurrentDriectory As Long, _
               lpStartupInfo As STARTUPINFO, _
               lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lptitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type
Public Const NORMAL_PRIORITY_CLASS = &H20&
Public Const INFINITE = -1&
Public Const STARTF_USESHOWWINDOW = &H1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWDEFAULT = 10


'MES ����
Public sMESEquipCode As String              'MES �����ڵ�
Public sMESEquipName As String              'MES �����̸�
Public sMESLineNum As String                'MES ���γѹ�
Public sMESProgressCode As String           'MES �����ڵ�
Public sMESProcess As String                'MES ���μ���
Public sMESFileSavePath As String           'MES ����������   (PC)
Public sMESFileSendPath As String           'MES ���Ϻ����°�� (��Ʈ��ũ����̺� S:\)
Public sMESLogSavePath As String            'MES �α�������   (PC)
Public iDataRow As Integer                  'Data File ������ Data ��� ���� ����
Public bPathSelect(0 To 2) As Boolean       '������ ���� boolean

Public ms_Source As String                  '�߽��� ID
Public ms_Destination As String             '������ ID

Public bMESReply As Boolean                'MES Reply ����
Public iTmrRecipe As Integer               '������ Ÿ�̸� ī��Ʈ
Public iTmrRecipePM As Integer             '������ �Ķ���� Ÿ�̸� ī��Ʈ
Public iTmrLogin As Integer                '�α��� Ÿ�̸� ī��Ʈ
Public iTmrDate As Integer                 '�ð����� Ÿ�̸� ī��Ʈ
Public sDateTimeCheck As String            'QCP ���ϰ� JPG ������ �ð����� ���� �����ϰ� ��û�ؼ� QCP �����Ҷ� �ð��� ������ ���� �ִ´�.

Public sRecipeID(1 To 10) As String
Public sNowRecipeID As String
Public iNowRecipeID As Integer
Public sBufRecipeID As String              '������߰���
Public iBufRecipeID As Integer             '������߰���
Public sRecipeComment(1 To 10) As String
Public iRecipeIDcount As Integer
Public iMESGridClick As Integer                 '�׸��� Ŭ�� ������ ��ȣ
Public iMESGridClickIdx As Integer              '�׸��� Ŭ�� �׸��� ��ȣ

Public sMsgString(0 To 999) As String       'MES �޼���
Public sPathName As String                  '���
Public sMESDate As String                   'MES �� ���� ���� ��¥�� �ð�
Public iMesSysbyte As Integer               'PC -> MES ���� �ý��� ����Ʈ
Public iMesSysbyteR As Integer              'Reply �ý��� ����Ʈ (MES -> PC �ι��� �ý��۹���Ʈ�� �״�� ��������)
Public iParamCount(1 To 10) As Integer               'Parameter Count
Public sParamCount(1 To 10) As String
Public sParamName_SV(1 To 10, 0 To 99) As String        'Sv Code Name
Public sParamName_SVsj As String        'Sv Code Name
Public sParamName_PV(1 To 10, 0 To 99) As String     'Pv Code Name
Public sParamName_NG(1 To 10, 0 To 99) As String     'NG Code Name

Public sParamValue(1 To 10, 0 To 99) As String
Public sParamMinValue(1 To 10, 0 To 99) As String
Public sParamMaxValue(1 To 10, 0 To 99) As String

Public dParamValue(1 To 10, 0 To 99) As Double
Public dParamMinValue(1 To 10, 0 To 99) As Double
Public dParamMaxValue(1 To 10, 0 To 99) As Double

Public sMesUserID As String                 'MES ����� ID
Public sMesUserPass As String               'MES ����� Password

'��Ʈ��ũ ����̺� �������
Public bNetDriveConnect As Boolean          'True ����
Public bNetDriveExist As Boolean            '�ݵ���̺� ���� ���� (true ����)
Public sMesPCIP As String                     '�ݵ���̺� PC ������
Public sMesPCID As String                 '�ݵ���̺� PC ����
Public sMesPCPW As String                 '�ݵ���̺� PC ��ȣ


Public Sub MES_ServerOpen()
On Error Resume Next
    frmMain.WinsockMES.Close
    frmMain.WinsockMES.LocalPort = CLng(sMESPort)
    frmMain.WinsockMES.Listen
    'frmMain.lstMESSocket.Clear
    frmMain.shpMESSock.BackColor = vbYellow
    'frmMain.lblMESServer.Caption = "���� OPEN"
End Sub
'������ ���� �������� �Ѵ�.. (����ö���� �ҽ� ����)
Public Sub ChangeViewSection(frmSection As Form)
    With frmMESMain
        frmSection.Top = 0
        frmSection.Left = 0
        Call SetParent(frmSection.hWnd, .picSection.hWnd)
    End With
    frmSection.Move 0, 0, frmSection.Width, frmSection.Height
    frmSection.Show
End Sub

Public Function MES_MakingDataMsg(msgid As String, pItem As String) As String
On Error GoTo err:
Dim tmpStrDefaultMsg As String  '����Ʈ �޼���
Dim tmpStrDataMsg As String     '������ �޼���
Dim tmpStrMidMsg As String      '����Ʈ + ������
Dim tmpStrHeaderMsg As String   '��� �޼���
Dim tmpStrLastMsg As String     '���� �޼���
Dim tmpMsgLen As Integer        '�޼��� ����
Dim tmpS_IDcode As String       '���̵� �ڵ� �ӽ� ���� ����
Dim i As Integer

    If sIDCode(0) = "NOID" Then
        tmpS_IDcode = ""
    Else
        tmpS_IDcode = sIDCode(0)
    End If
    
    tmpStrDefaultMsg = ""
    tmpStrDefaultMsg = tmpStrDefaultMsg & "   <DEFAULT>" & vbCrLf
    If msgid = "DATE_REPLY" Or msgid = "LINKTEST_REPLY" Then
        tmpStrDefaultMsg = tmpStrDefaultMsg & "      <SYSTEM_BYTES>" & iMesSysbyteR & "</SYSTEM_BYTES>" & vbCrLf '1~9999
    Else
        If iMesSysbyte = 9999 Then
            iMesSysbyte = 0
        End If
        iMesSysbyte = iMesSysbyte + 1
        tmpStrDefaultMsg = tmpStrDefaultMsg & "      <SYSTEM_BYTES>" & CStr(iMesSysbyte) & "</SYSTEM_BYTES>" & vbCrLf '1~9999
    End If
    
    If g_Timeout = 0 Then
        g_LastMesSystemByte = CStr(iMesSysbyte)
    End If
    
    tmpStrDefaultMsg = tmpStrDefaultMsg & "      <EQUIP_ID>" & sMESEquipCode & "</EQUIP_ID>" & vbCrLf
    tmpStrDefaultMsg = tmpStrDefaultMsg & "      <LOT_ID>" & tmpS_IDcode & "</LOT_ID>" & vbCrLf
    tmpStrDefaultMsg = tmpStrDefaultMsg & "      <FROM>" & sMESEquipCode & "</FROM>" & vbCrLf
    tmpStrDefaultMsg = tmpStrDefaultMsg & "      <TO>MES</TO>" & vbCrLf
    tmpStrDefaultMsg = tmpStrDefaultMsg & "      <RECIPE_ID>" & sModelName & "</RECIPE_ID>" & vbCrLf
    tmpStrDefaultMsg = tmpStrDefaultMsg & "      <USER_ID>" & sMesUserID & "</USER_ID>" & vbCrLf
    tmpStrDefaultMsg = tmpStrDefaultMsg & "      <OPER>" & sMESProgressCode & "</OPER>" & vbCrLf
    tmpStrDefaultMsg = tmpStrDefaultMsg & "      <PROCESS>" & sMESProcess & "</PROCESS>" & vbCrLf
    tmpStrDefaultMsg = tmpStrDefaultMsg & "      <DATE>" & Format(Now, "YYYY-MM-DD HH:MM:SS") & "</DATE>" & vbCrLf
    tmpStrDefaultMsg = tmpStrDefaultMsg & "   </DEFAULT>" & vbCrLf
    
    tmpStrDataMsg = ""
    tmpStrDataMsg = tmpStrDataMsg & Space(3) & "<DATA>" & vbCrLf
    
    frmMain.tmrMesTimeout.Enabled = True
    frmMain.tmrMesTimeout.Interval = g_TimeoutInterval
    
    Select Case msgid
        
        'Equipment State
        Case "EQ_STATE_EVENT"
            tmpStrDataMsg = tmpStrDataMsg & Space(6) & "<READY>READY</READY>" & vbCrLf
            'AUTO
            If pItem = "AUTO" Then
                tmpStrDataMsg = tmpStrDataMsg & Space(6) & "<STATE>AUTO</STATE>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(6) & "<PROCESS_STATE>IDLE</PROCESS_STATE>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(6) & "<MAINT_STATE></MAINT_STATE>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(6) & "<MAINT_CODE></MAINT_CODE>" & vbCrLf
            ElseIf pItem = "PROCESS" Then
                tmpStrDataMsg = tmpStrDataMsg & Space(6) & "<STATE>AUTO</STATE>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(6) & "<PROCESS_STATE>PROCESSING</PROCESS_STATE>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(6) & "<MAINT_STATE></MAINT_STATE>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(6) & "<MAINT_CODE></MAINT_CODE>" & vbCrLf
            Else           '���� ���� �� ����Ʈ ó���ʿ�
                tmpStrDataMsg = tmpStrDataMsg & Space(6) & "<STATE>MANUAL</STATE>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(6) & "<PROCESS_STATE>DOWN</PROCESS_STATE>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(6) & "<MAINT_STATE></MAINT_STATE>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(6) & "<MAINT_CODE></MAINT_CODE>" & vbCrLf
            End If
            
        Case "LOGIN_EVENT"
            tmpStrDataMsg = tmpStrDataMsg & Space(6) & "<ID>" & sMesUserID & "</ID>" & vbCrLf
            tmpStrDataMsg = tmpStrDataMsg & Space(6) & "<PASSWORD>" & sMesUserPass & "</PASSWORD>" & vbCrLf
            
        Case "RECIPE_EVENT"
        
        Case "RECIPE_CHANGE_EVENT"
            'RECIPE ����
            tmpStrDataMsg = tmpStrDataMsg & Space(6) & "<PARAM_COUNT>" & iToolCount / 2 & "</PARAM_COUNT>" & vbCrLf
            
            For i = 1 To iToolCount / 2
                tmpStrDataMsg = tmpStrDataMsg & Space(6) & "<PARAM_DATA>" & vbCrLf
                'tmpStrDataMsg = tmpStrDataMsg & Space(9) & "<PARAM_NAME>" & sParamName_SV(iNowRecipeID, i) & "</PARAM_NAME>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(9) & "<PARAM_VALUE>" & dParamValue(iNowRecipeID, i) & "</PARAM_VALUE>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(9) & "<PARAM_MINVALUE>" & dParamMinValue(iNowRecipeID, i) & "</PARAM_MINVALUE>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(9) & "<PARAM_MAXVALUE>" & dParamMaxValue(iNowRecipeID, i) & "</PARAM_MAXVALUE>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(6) & "</PARAM_DATA>" & vbCrLf
            Next i
            
        Case "RECIPE_SV_CHANGE_EVENT"
            'RECIPE �� ����
            tmpStrDataMsg = tmpStrDataMsg & Space(6) & "<PARAM_COUNT>" & iToolCount / 2 & "</PARAM_COUNT>" & vbCrLf
            
            For i = 1 To iToolCount / 2
                tmpStrDataMsg = tmpStrDataMsg & Space(6) & "<PARAM_DATA>" & vbCrLf
                'tmpStrDataMsg = tmpStrDataMsg & Space(9) & "<PARAM_NAME>" & sParamName_SV(iNowRecipeID, i) & "</PARAM_NAME>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(9) & "<PARAM_VALUE>" & frmMESRecipePM.txtPCValueOri(i).Text & "</PARAM_VALUE>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(9) & "<PARAM_MINVALUE>" & frmMESRecipePM.txtPCValueMin(i) & "</PARAM_MINVALUE>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(9) & "<PARAM_MAXVALUE>" & frmMESRecipePM.txtPCValueMax(i) & "</PARAM_MAXVALUE>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(6) & "</PARAM_DATA>" & vbCrLf
            Next i
        
        Case "NG_PRODUCT_EVENT"
            ' MES ����� ��û���� loss count�� ������ 1 ������
            Dim NGCount As Integer
            Dim NSDCount As Integer
            
            NGCount = 0
            For i = 0 To 3
                If g_Judge(i) = False Then
                    NGCount = NGCount + 1
                End If
            Next i
            
            NSDCount = 0
            For i = 0 To 5
                If g_Judge(i + 4) = False Then
                    NSDCount = NSDCount + 1
                End If
            Next i
            
            If NSDCount > 0 Or g_Judge(11) = False Then
                NGCount = NGCount + 1
            End If
            
            tmpStrDataMsg = tmpStrDataMsg & Space(6) & "     <LOSS_COUNT>" & CStr(NGCount) & "</LOSS_COUNT>" & vbCrLf
                       
            For i = 0 To 3
                If g_Judge(i) = False Then
                    tmpStrDataMsg = tmpStrDataMsg & Space(6) & "     <LOSS_DATA>" & vbCrLf
                    tmpStrDataMsg = tmpStrDataMsg & Space(6) & "          <LOT_ID>" & tmpS_IDcode & "</LOT_ID>" & vbCrLf
                    tmpStrDataMsg = tmpStrDataMsg & Space(6) & "          <LOSS_CD>" & sParamName_NG(iNowRecipeID, i + 1) & "</LOSS_CD>" & vbCrLf
                    tmpStrDataMsg = tmpStrDataMsg & Space(6) & "     </LOSS_DATA>" & vbCrLf
                End If
            Next i
            
            If NSDCount > 0 Or g_Judge(11) = False Then
                tmpStrDataMsg = tmpStrDataMsg & Space(6) & "     <LOSS_DATA>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(6) & "          <LOT_ID>" & tmpS_IDcode & "</LOT_ID>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(6) & "          <LOSS_CD>" & sParamName_NG(iNowRecipeID, 5) & "</LOSS_CD>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(6) & "     </LOSS_DATA>" & vbCrLf
            End If
                
            
        Case "QMS_EVENT"
            Dim TempPvName As String
            
            tmpStrDataMsg = tmpStrDataMsg & Space(6) & "<TRAY_ID></TRAY_ID>" & vbCrLf
            tmpStrDataMsg = tmpStrDataMsg & Space(6) & "<LOT_COUNT></LOT_COUNT>" & vbCrLf
            tmpStrDataMsg = tmpStrDataMsg & Space(6) & "<SV>" & vbCrLf
'            For i = 0 To 9
'                tmpStrDataMsg = tmpStrDataMsg & Space(9) & "<PARAM_DATA>" & vbCrLf
'                tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<SLOT_ID></SLOT_ID>" & vbCrLf
'                tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<PARAM_NAME>" & Left$(sParamName_SV(iNowRecipeID, i + 1), Len(sParamName_SV(iNowRecipeID, i + 1)) - 2) & Format(CInt(Right(sParamName_SV(iNowRecipeID, i + 1), 2)) + 0, "00") & "</PARAM_NAME>" & vbCrLf
'                tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<PARAM_VALUE>" & dSpecOri(i) & "</PARAM_VALUE>" & vbCrLf
'                tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<STEP_NAME></STEP_NAME>" & vbCrLf
'                tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<STEP_CONDITION></STEP_CONDITION>" & vbCrLf
'                tmpStrDataMsg = tmpStrDataMsg & Space(9) & "</PARAM_DATA>" & vbCrLf
'                tmpStrDataMsg = tmpStrDataMsg & Space(9) & "<PARAM_DATA>" & vbCrLf
'                tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<SLOT_ID></SLOT_ID>" & vbCrLf
'                tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<PARAM_NAME>" & Left$(sParamName_SV(iNowRecipeID, i + 1), Len(sParamName_SV(iNowRecipeID, i + 1)) - 2) & Format(CInt(Right(sParamName_SV(iNowRecipeID, i + 1), 2)) + 1, "00") & "</PARAM_NAME>" & vbCrLf
'                tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<PARAM_VALUE>" & dSpecMax(i) & "</PARAM_VALUE>" & vbCrLf
'                tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<STEP_NAME></STEP_NAME>" & vbCrLf
'                tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<STEP_CONDITION></STEP_CONDITION>" & vbCrLf
'                tmpStrDataMsg = tmpStrDataMsg & Space(9) & "</PARAM_DATA>" & vbCrLf
'                tmpStrDataMsg = tmpStrDataMsg & Space(9) & "<PARAM_DATA>" & vbCrLf
'                tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<SLOT_ID></SLOT_ID>" & vbCrLf
'                tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<PARAM_NAME>" & Left$(sParamName_SV(iNowRecipeID, i + 1), Len(sParamName_SV(iNowRecipeID, i + 1)) - 2) & Format(CInt(Right(sParamName_SV(iNowRecipeID, i + 1), 2)) + 2, "00") & "</PARAM_NAME>" & vbCrLf
'                tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<PARAM_VALUE>" & dSpecMin(i) & "</PARAM_VALUE>" & vbCrLf
'                tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<STEP_NAME></STEP_NAME>" & vbCrLf
'                tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<STEP_CONDITION></STEP_CONDITION>" & vbCrLf
'                tmpStrDataMsg = tmpStrDataMsg & Space(9) & "</PARAM_DATA>" & vbCrLf
'            Next i
            tmpStrDataMsg = tmpStrDataMsg & Space(6) & "</SV>" & vbCrLf
            
            tmpStrDataMsg = tmpStrDataMsg & Space(6) & "<PV>" & vbCrLf
            For i = 0 To 3
                tmpStrDataMsg = tmpStrDataMsg & Space(9) & "<PARAM_DATA>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<SLOT_POSITION></SLOT_POSITION>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<CELL_ID>" & tmpS_IDcode & "</CELL_ID>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<PARAM_NAME>" & sParamName_PV(iNowRecipeID, i + 1) & "</PARAM_NAME>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<PARAM_VALUE>" & Format(g_Distance(i), "#0.00") & "</PARAM_VALUE>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<STEP_NAME></STEP_NAME>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<STEP_CONDITION></STEP_CONDITION>" & vbCrLf
                tmpStrDataMsg = tmpStrDataMsg & Space(9) & "</PARAM_DATA>" & vbCrLf
            Next i
            
            For i = 0 To 5
                If Not (i = 1 Or i = 3) Then
                    tmpStrDataMsg = tmpStrDataMsg & Space(9) & "<PARAM_DATA>" & vbCrLf
                    tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<SLOT_POSITION></SLOT_POSITION>" & vbCrLf
                    tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<CELL_ID>" & tmpS_IDcode & "</CELL_ID>" & vbCrLf
                    tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<PARAM_NAME>" & sParamName_PV(iNowRecipeID, i + 5) & "</PARAM_NAME>" & vbCrLf
                    tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<PARAM_VALUE>" & Format(g_NsdDistance(i), "#0.00") & "</PARAM_VALUE>" & vbCrLf
                    tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<STEP_NAME></STEP_NAME>" & vbCrLf
                    tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<STEP_CONDITION></STEP_CONDITION>" & vbCrLf
                    tmpStrDataMsg = tmpStrDataMsg & Space(9) & "</PARAM_DATA>" & vbCrLf
                End If
            Next i
            
            'NSD ����
            tmpStrDataMsg = tmpStrDataMsg & Space(9) & "<PARAM_DATA>" & vbCrLf
            tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<SLOT_POSITION></SLOT_POSITION>" & vbCrLf
            tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<CELL_ID>" & tmpS_IDcode & "</CELL_ID>" & vbCrLf
            TempPvName = Left(sParamName_PV(iNowRecipeID, 10), Len(sParamName_PV(iNowRecipeID, 10)) - 3) & Format(Right(sParamName_PV(iNowRecipeID, 10), 3) + 10, "000")
            tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<PARAM_NAME>" & TempPvName & "</PARAM_NAME>" & vbCrLf
            tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<PARAM_VALUE>" & IIf(g_Judge(11) = True, "1", "2") & "</PARAM_VALUE>" & vbCrLf
            tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<STEP_NAME></STEP_NAME>" & vbCrLf
            tmpStrDataMsg = tmpStrDataMsg & Space(12) & "<STEP_CONDITION></STEP_CONDITION>" & vbCrLf
            tmpStrDataMsg = tmpStrDataMsg & Space(9) & "</PARAM_DATA>" & vbCrLf
            
            tmpStrDataMsg = tmpStrDataMsg & Space(6) & "</PV>" & vbCrLf
        Case "TIMEOUT_EVENT"
            tmpStrDataMsg = tmpStrDataMsg & Space(6) & "<TO_SYSTEM_BYTES>" & CStr(g_TTemp) & "</TO_SYSTEM_BYTES>" & vbCrLf  '1~9999
        Case Else
            tmpStrDataMsg = tmpStrDataMsg & Space(6) & "<RETURN_VALUE>1</RETURN_VALUE>" & vbCrLf
            tmpStrDataMsg = tmpStrDataMsg & Space(6) & "<ERROR_MSG></ERROR_MSG>" & vbCrLf
            frmMain.tmrMesTimeout.Enabled = False
    End Select
    
    tmpStrDataMsg = tmpStrDataMsg & Space(3) & "</DATA>" & vbCrLf
    
    '============XML �������� ��ȯ============
    'Default
    Call lpSetXmlEngine(tmpStrDefaultMsg)
    'Data
    Call lpSetXmlEngine(tmpStrDataMsg)
    '=========================================

    tmpStrMidMsg = ""
    tmpStrMidMsg = tmpStrMidMsg & tmpStrDefaultMsg & vbCrLf
    tmpStrMidMsg = tmpStrMidMsg & tmpStrDataMsg & vbCrLf
    tmpStrMidMsg = tmpStrMidMsg & "</MESSAGE>" & vbCrLf
    
    tmpMsgLen = Len(tmpStrMidMsg)
    
    tmpStrHeaderMsg = ""
    tmpStrHeaderMsg = tmpStrHeaderMsg & "<MESSAGE>" & vbCrLf
    tmpStrHeaderMsg = tmpStrHeaderMsg & "   <HEADER>" & vbCrLf
    tmpStrHeaderMsg = tmpStrHeaderMsg & "      <MSG_ID>" & msgid & "</MSG_ID>" & vbCrLf
    tmpStrHeaderMsg = tmpStrHeaderMsg & "      <MSG_LEN>" & CStr(tmpMsgLen) & "</MSG_LEN>" & vbCrLf
    tmpStrHeaderMsg = tmpStrHeaderMsg & "   </HEADER>" & vbCrLf
    
    'STX + Header �� Data, Default �޼��� ���� + ETX �߰�
    tmpStrLastMsg = ""
    tmpStrLastMsg = tmpStrLastMsg & Chr(2) & vbCrLf
    tmpStrLastMsg = tmpStrLastMsg & tmpStrHeaderMsg + tmpStrMidMsg
    tmpStrLastMsg = tmpStrLastMsg & Chr(3)
    
    MES_MakingDataMsg = tmpStrLastMsg
Exit Function
err:
    ListBox_Append "XML Data ������ ������ �߻� �Ͽ����ϴ�.", 1
End Function

Public Function DJ_DataFileADD(Index As Integer) As String
Dim i As Integer
Dim tempstr As String
Dim tempStrSV As String
Dim tempStrPV As String
Dim tempStrTotal As String
Dim sDate As String
Dim stime As String
Dim sVBTab As String
    sVBTab = vbTab
    stime = Format(Time, "HH:MM:SS")
    
    For i = 0 To 9
        'tempStrSV = tempStrSV & sParamName_SV(iNowRecipeID, i + 1) & "=" & dSpecOri(i) & vbCrLf
        Dim lstr_tempSVcode As String
        '���ذ� ���ϱ�
        If Len(CStr(i * 3 + 1)) = 1 Then
            lstr_tempSVcode = "0" & CStr(i * 3 + 1)
        Else
            lstr_tempSVcode = CStr(i * 3 + 1)
        End If
        tempStrSV = tempStrSV & sParamName_SV(iNowRecipeID, i + 1) & "=" & dSpecOri(i) & vbCrLf
        
        '�÷��� ���� ���ϱ�
        If Len(CStr(i * 3 + 2)) = 1 Then
            lstr_tempSVcode = "0" & CStr(i * 3 + 2)
        Else
            lstr_tempSVcode = CStr(i * 3 + 2)
        End If
        tempStrSV = tempStrSV & sParamName_SV(iNowRecipeID, i + 1) & "=" & dSpecMax(i) & vbCrLf
        
        '���̳ʽ� ���� ���ϱ�
        If Len(CStr(i * 3 + 3)) = 1 Then
            lstr_tempSVcode = "0" & CStr(i * 3 + 3)
        Else
            lstr_tempSVcode = CStr(i * 3 + 3)
        End If
        tempStrSV = tempStrSV & sParamName_SV(iNowRecipeID, i + 1) & "=" & dSpecMin(i) & vbCrLf
        
    Next i
    
    For i = 0 To 3
        tempStrPV = tempStrPV & sParamName_PV(iNowRecipeID, i + 1) & "=" & Format(g_Distance(i), "#0.00") & vbCrLf
    Next i
    
    For i = 0 To 5
        tempStrPV = tempStrPV & sParamName_PV(iNowRecipeID, i + 5) & "=" & Format(g_NsdDistance(i), "#0.00") & vbCrLf
    Next i
    
    tempstr = "[SV]" & vbCrLf & tempStrSV & "[PV]" & vbCrLf & tempStrPV
    
    DJ_DataFileADD = tempstr
    
End Function

Public Sub LOG_TACK(cellID As String, PLC As Long, GRAB As Long, INSPECTION As Long, JUDGEMENT As Long, SCREENSHOT As Long, MES As Long, INSPECTIONEND As Long)
On Error GoTo ErrorHandle

Dim ff As Integer
Dim LogPath As String
Dim HEADER As Boolean


    ff = FreeFile
    LogPath = "D:\LOG\TACKTIME\" & Format(Date, "YYYY") & "\" & Format(Date, "MM") & "\" & Format(Date, "DD") & "\" & Format(Date, "YYYYMMDD") & "_" & Format(Time, "HH") & ".csv"
    
    Call Create_DIR("D:\LOG\TACKTIME\" & Format(Date, "YYYY") & "\" & Format(Date, "MM") & "\" & Format(Date, "DD"))
    
    
    
    If Len(Dir$(LogPath)) = 0 Then
        HEADER = True
    End If
    
    Open LogPath For Append As ff
        If HEADER Then
            Print #ff, "�ð�, ID, ���峻���ޱ�, �����Կ�, �˻�, ����, ��ũ����, MES����, �˻�Ϸ�"
        End If
        Print #ff, Format(Now, "YYYY-MM-DD hh:mm:ss"), ",", cellID, ",", CStr(PLC), ",", CStr(GRAB), ",", CStr(INSPECTION), ",", CStr(JUDGEMENT), ",", CStr(SCREENSHOT), ",", CStr(MES), ",", CStr(INSPECTIONEND)
    Close #ff


    Exit Sub
ErrorHandle:
    
    Close #ff
    

End Sub


Public Sub DataFileSave(Index As Integer, Data As String, Path As String)
On Error GoTo err:
    Dim i As Integer
    Dim ff As Integer
    
    Dim sDate_Y As String
    Dim sDate_M As String
    Dim sDate_D As String
    Dim sDate_TOT As String
    
    Dim SDataRow As String
    Dim tempstr(0 To 9) As String
    Dim sidtemp As String
    Dim sttime As String
    
    ff = FreeFile
    'iDataRow = CInt(frmMESFunction.txtDataRow.Text)
    SDataRow = Format((iToolCount + 2), "000000")
    tempstr(0) = "01"
    tempstr(1) = "1"

    TTime = Format(Time, "HHMMSS")
    Tdate = Format(Date, "YYYYMMDD")
    sttime = Format(Time, "HH:MM:SS")

    sDate_Y = Left(Tdate, 4)
    sDate_M = Mid(Tdate, 5, 2)
    sDate_D = Right(Tdate, 2)
    sDate_TOT = sDate_Y & "/" & sDate_M & "/" & sDate_D
    
    
        sPathName = Path
        Open sPathName For Append As ff
            Print #ff, sMESEquipCode
            Print #ff, sDate_TOT
            Print #ff, sttime
            If sIDCode(Index) = "NOID" Then
                sidtemp = ""
            Else
                sidtemp = sIDCode(Index)
            End If
            Print #ff, sidtemp
            Print #ff, sNowRecipeID
            Print #ff, sMesUserID
            Print #ff, sMESLineNum
            Print #ff, sMESProgressCode
            Print #ff, tempstr(0)
            Print #ff, tempstr(1)
            Print #ff, tempstr(2)
            Print #ff, tempstr(3)
            Print #ff, SDataRow
            Print #ff, Data
        Close #ff
Exit Sub

err:

End Sub
Public Sub MES_ImageFile_Send()
On Error GoTo err:

Dim SourceFile As String
Dim DesFile As String
Dim ret As Boolean
    
    SourceFile = sMESFileSavePath & "\*.*"
    DesFile = "\\" & sMesPCIP '�ݵ���̺긦 ������� �ʰ� ���� ��η� �����Ͽ� ��� sMESFileSendPath & "\"
    ret = ConnectThisNetworkDrive(DesFile, "", sMesPCID, sMesPCPW)
    If ret = 85 Or ret = 0 Or ret = True Then
        Call SynchronizedShell("xcopy" & " " & SourceFile & " " & DesFile, 0)
    ElseIf ret = 1326 Then
        MsgBox "Mes ȯ�漳������ ���̵� �� ��й�ȣ�� Ȯ���ϼ���"
    End If
    'Call SynchronizedShell("xcopy" & " " & SourceFile & " " & DesFile, 0)

Exit Sub
err:
    'Call MES_NetDriveConnect
End Sub
Public Sub SynchronizedShell(command As String, windowsStyle As VbAppWinStyle)
Dim tm As Double
Dim lRet As Long
Dim lRet2 As Long
Dim vProc As PROCESS_INFORMATION
Dim vStart As STARTUPINFO
Dim vRv As Long

    vStart.cb = Shell(command, windowsStyle)
    Dlay_T (2)
    vStart.dwFlags = STARTF_USESHOWWINDOW
    vStart.wShowWindow = SW_SHOWDEFAULT 'SW_SHOWMAXIMIZED

    ' Process ����
    vRv = CreateProcess(0&, RunCommand, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, vStart, vProc)

''''    lRet = Shell(command, windowsStyle)
''''    lRet = CreateProcess(&H100000, False, lngPID)
    tm = Timer + 1
    Do
        'DoEvents
        vRv = WaitForSingleObject(vProc.hProcess, INFINITE)
    Loop Until vRv <> 0 Or tm < Timer
    Dlay_T (0.2)
    If vRv <> 0 Then
        Kill sMESFileSavePath & "\*.*"
    End If
    ' Process ����
    'vRv = CloseHandle(vProc.hProcess) ���μ��� ���ᰡ �� ���� �ʾ� ��ü
    Shell "tskill xcopy"
End Sub
Public Sub DJ_MakeBat()
On Error GoTo err
Dim ff As Integer
Dim i As Integer
    ff = FreeFile
    Open "D:" & "\FILECOPY" & ".bat" For Output As ff
        Print #ff, "copy " & Chr(34) & "D:\MES\SEND\*.*" & Chr(34) & " " & Chr(34) & "S:"; Chr(34)
        Print #ff, "DEL " & Chr(34) & "D:\MES\SEND\*.*" & Chr(34) & "/q"
    Close ff
    Exit Sub
err:
    Close ff
End Sub
Public Sub FilrCopyShow(fromFile As String, toFile As String)
    Dim SHFile As SHFILEOPSTRUCT
    With SHFile
        .hWnd = 0
        .wFunc = FO_COPY
        .pFrom = fromFile
        .pTo = toFile
        .fFlags = FOF_ALLOWUNDO
    End With
    SHFileOperation SHFile
End Sub
Public Sub MES_FindNetDrive(sDrive As String)
Dim fso As New FileSystemObject
Dim strDrive As String

    strDrive = Chr(Asc(sDrive))
    If fso.DriveExists(strDrive) = True Then
        bNetDriveExist = True
    Else
        bNetDriveExist = False
    End If

End Sub
Public Sub MES_NetDriveConnect()
On Error Resume Next
Dim errInfo As Long
Dim cbBuff As Long
Dim IpBuff As String
Dim bretValue As Boolean
Dim lretValue As Long
Dim lipadress As Long
Dim itemp As Integer
Dim ECHO As ICMP_ECHO_REPLY
    cbBuff = 255
    IpBuff = String$(cbBuff, Chr$(0))

Dlay_T (0.5)
    lipadress = inet_addr(DJSJ_XMLData_Find(1, "", "\", sMesPCIP, itemp))           'sMesPCIP �� IP + ���� ����̱� ������ IP �� ©��ͼ� Long ������ ��ȯ�Ѵ�.
    If lipadress <> INADDR_NONE Then       ' -1 �� �ƴϸ�
        bretValue = MES_PingCheck(lipadress, "", ECHO)         'Ping Test �� �����Ѵ�.
    End If
    If bretValue = True Then                                   'Ping �� ���������
        frmMain.shpMESNetDrive.BackColor = vbGreen
'        Call MES_ImageFile_Send
        Call WriteMelsec(frmMain.ActEasyIF, sMelsecAddrAlarmNetDrive, 0)
    Else                 '�� �� ���� �ȵ������� ������ �õ� �Ѵ� (���������� �״��� �˻綧 �Ѳ����� �����Ѵ�)

        ListBox_Append Time & "  ��Ʈ�p �����߿����� �߻��Ͽ����ϴ�!.....��Ʈ�p Ȯ���ϼ���!", 1
        frmMain.shpMESNetDrive.BackColor = vbRed
        bNetDriveConnect = False
        Call WriteMelsec(frmMain.ActEasyIF, sMelsecAddrAlarmNetDrive, 1)
    End If

End Sub

Public Function MES_PingCheck(sipAdress As Long, sdata As String, ECHO As ICMP_ECHO_REPLY) As Boolean
Dim bPort As Long
Dim stemp As String

    ' ����Ʋ �����Ѵ�.
    bPort = IcmpCreateFile()
    ' ��Ʈ������ �����ϸ�
    If bPort <> 0 Then
        ' Ping �õ�
        Call IcmpSendEcho(bPort, sipAdress, sdata, CLng(Len(sdata)), 0, ECHO, Len(ECHO), PING_TIMEOUT)
        ' Ping �� ���°��� �����Ѵ�.
        If ECHO.status = 0 Then
            If Abs(ECHO.Address) > 0 Then
                MES_PingCheck = True
            Else
                MES_PingCheck = False
            End If
        Else
            MES_PingCheck = False
        End If
        Call IcmpCloseHandle(bPort)     ' ��Ʈ�� �ݴ´�
    End If
End Function
Public Sub MES_SendData(sdata As String)
On Error GoTo err:
Dim i As Integer
Dim stemp As String
    If frmMain.WinsockMES.State <> sckClosed Then
        frmMain.WinsockMES.SendData sdata
    End If
Exit Sub
err:
    'MsgBox "MES�� ������ ���������ϴ�." & vbCrLf & "��Ż��¸� Ȯ���ϼ���", vbCritical, "��� ����"
    
End Sub
'��Ʈ��ũ ����̺� ����
Public Function ConnectThisNetworkDrive(sServer As String, sDrv As String, sUserID As String, sPass As String) As Long
On Error GoTo err:

    Dim NETR As NETRESOURCE
    Dim errInfo As Long
   
    With NETR
        .dwScope = RESOURCE_GLOBALNET
        .dwType = RESOURCETYPE_DISK
        .dwDisplayType = RESOURCEDISPLAYTYPE_SHARE
        .dwUsage = RESOURCEUSAGE_CONNECTABLE
        .lpRemoteName = sServer
        .lpLocalName = ""   'sDrv
    End With
   
    errInfo = WNetAddConnection2(NETR, sPass, sUserID, CONNECT_UPDATE_PROFILE)
    ConnectThisNetworkDrive = errInfo
Exit Function

err:
    
End Function
'XML ���Ŀ��� ���ϴ� String ã�ƿ��� (ex : StartPoint -> �˻������� �ڸ��� , StartStr -> <HEADER> , EndStr -> </HEADER> , ToTalData -> XML DATA ��ü , end_LenStr �� EndStr �� ������ �ڸ����� ��ȯ)
'DJSJ_XMLData_Find ���� <HEADER> �� </HEADER> ���̿� �ִ� ���ڿ��� ��ȯ �Ѵ�.
Public Function DJSJ_XMLData_Find(StartPoint As Integer, StartStr As String, EndStr As String, TotalData As String, ByRef end_LenStr As Integer) As String
On Error GoTo err:

    Dim tempStrMsg As String
    Dim tempNstart As Integer
    Dim tempNend As Integer
    
    tempNstart = InStr(StartPoint, TotalData, StartStr) + Len(StartStr)
    tempNend = InStr(StartPoint, TotalData, EndStr)
    tempStrMsg = Mid(TotalData, tempNstart, tempNend - tempNstart)
    
    DJSJ_XMLData_Find = tempStrMsg
    end_LenStr = tempNend
Exit Function
err:

End Function
Public Sub MES_DATASEND_FUNC(msgid As String, msgid2 As String, ReplySys As String, Optional flag As Boolean = False)
Dim tempStrMsg As String
    If ReplySys = "" Then
        'iMesSysbyte = iMesSysbyte + 1
    End If
    
    If flag = False Then
        g_LastMesMsgId = msgid
        g_LastMesMsgItem1 = msgid2
        g_LastMesMsgItem2 = ReplySys
        g_TTemp = iMesSysbyte + 1
    End If
    
    frmMESMain.txtSendMES.Text = ""
    bMESReply = False
    tempStrMsg = ""
    tempStrMsg = MES_MakingDataMsg(msgid, msgid2)
    If flag = False Then
        g_LastMesMsg = tempStrMsg
    End If
    frmMESMain.txtSendMES.Text = tempStrMsg
    Call DJ_MESmsgLogSave(tempStrMsg)
    Call MES_SendData(tempStrMsg)
End Sub
Public Function MESLOGWrite(LogStr As String)
On Error Resume Next
' ����
    Dim FileNameT As String
    Dim FileNumberT As Integer
    FileNameT = sMESLogSavePath & "\" & Date & "_MESLOG.txt"
    FileNumberT = FreeFile
    Open FileNameT For Append As FileNumberT
        Print #FileNumberT, Date & " - " & Time & "  :  " & LogStr
    Close #FileNumberT
End Function
'XML ���� ��ȯ
Private Sub lpSetXmlEngine(ByRef pMsg As String)
    
    On Error GoTo err
    
    'XML ���� ��ȯ(Data)
    Set clsXmlEngine = New XmlEngine
    With clsXmlEngine
        .InitializeBeforeParsing
        .BuildTreeDuringParse = True
        .AppendAndParse pMsg
        .CleanupAfterParsing
        
        'Space(3) CLASS ���� �����ؾ� �Ǵµ� �����Ƽ� ~~
        pMsg = ""
        pMsg = Space(3) & .RootElement.ToXml
    End With
    
    If Not clsXmlEngine Is Nothing Then
        Set clsXmlEngine = Nothing
    End If
    
    Exit Sub
    '--------------------------------------------------------------------------------------------
err:
    Call MESLOGWrite("Error, lpSetXmlEngine " & err.Description)
    
    If Not clsXmlEngine Is Nothing Then
        Set clsXmlEngine = Nothing
    End If
    
    Resume Next
    
End Sub

Public Sub DJ_EquipSpecLoad(sPname As String, iCount As Integer, Rcpindex As Integer)
Dim i As Integer
    'If sPname = sParamName_SV(Rcpindex, iCount) Then
        frmMESRecipePM.lblMESParameter(iCount).Caption = sSpecName(iCount - 1)
        frmMESRecipePM.txtPCValueOri(iCount).Text = dSpecOri(iCount - 1)
        frmMESRecipePM.txtPCValueMin(iCount).Text = dSpecOriMin(iCount - 1)
        frmMESRecipePM.txtPCValueMax(iCount).Text = dSpecOriMax(iCount - 1)
    'End If

End Sub
Public Sub DJ_ComparePN(sPname As String, iCount As Integer, Rcpindex As Integer)  'sPname �� �Ķ���� ���� , icount �� �Ķ���� ���� , rcpindex�� �������ε���
Dim i As Integer

    'If sPname = sParamName_SV(Rcpindex, iCount) Then
        frmMESRecipePM.lblMESParameter(iCount).Caption = sSpecName(iCount - 1)
        frmMESRecipePM.txtPCValueOri(iCount).Text = sParamValue(Rcpindex, iCount)
        frmMESRecipePM.txtPCValueMin(iCount).Text = sParamMinValue(Rcpindex, iCount)
        frmMESRecipePM.txtPCValueMax(iCount).Text = sParamMaxValue(Rcpindex, iCount)
    'End If
    
End Sub
Public Sub DJ_EquipSpecApply_OK()
Dim i As Integer
    For i = 0 To iParamCount(iNowRecipeID) - 1
        dSpecOri(i) = frmMESRecipePM.txtPCValueOri(i + 1).Text
        dSpecOriMin(i) = frmMESRecipePM.txtPCValueMin(i + 1).Text
        dSpecOriMax(i) = frmMESRecipePM.txtPCValueMin(i + 1).Text
        dSpecMin(i) = frmMESRecipePM.txtPCValueOri(i + 1).Text - frmMESRecipePM.txtPCValueMin(i + 1).Text
        dSpecMax(i) = frmMESRecipePM.txtPCValueMax(i + 1).Text - frmMESRecipePM.txtPCValueOri(i + 1).Text
        
'        frmMain.txtSpecOri(i).Text = dSpecOri(i)
'        frmMain.txtSpecMin(i).Text = dSpecMin(i)
'        frmMain.txtSpecMax(i).Text = dSpecMax(i)
    Next i
    
End Sub
Public Sub DJ_EquipSpecApply_NG()        '�Լ����� NG �� �پ������� ���� ������ ��������� �ٽ� �ҷ����� �Լ� �̴�. ������ ü������ ��� (������߰��� �ּ�)
On Error GoTo err:
Dim i As Integer
    For i = 0 To iParamCount(iNowRecipeID) - 1
        dSpecOri(i) = dParamValue(iNowRecipeID, i + 1)
        dSpecOriMin(i) = dParamMinValue(iNowRecipeID, i + 1)
        dSpecOriMax(i) = dParamMaxValue(iNowRecipeID, i + 1)
        dSpecMin(i) = dSpecOri(i) - dSpecOriMin(i)
        dSpecMax(i) = dSpecOriMax(i) - dSpecOri(i)
        
'        frmMain.txtSpecOri(i).Text = dSpecOri(i)
'        frmMain.txtSpecMin(i).Text = dSpecMin(i)
'        frmMain.txtSpecMax(i).Text = dSpecMax(i)
    Next i
Exit Sub
err:
    ListBox_Append Time & " �����Ͱ� �ùٸ��� �ʽ��ϴ�.", 1
End Sub

Public Sub MESWriteGrid(iCount As Integer, iRownum As Integer)
    frmMESRecipe.MSFlexGrid1.Rows = iRownum
    frmMESRecipe.MSFlexGrid1.TextMatrix(iRownum - 1, 0) = iCount
    frmMESRecipe.MSFlexGrid1.TextMatrix(iRownum - 1, 1) = sRecipeID(iCount)
    frmMESRecipe.MSFlexGrid1.TextMatrix(iRownum - 1, 2) = "TEST"
    
    frmMESRecipe.MSFlexGrid1.Row = 1
    frmMESRecipe.MSFlexGrid1.Col = 0
    frmMESRecipe.MSFlexGrid1.Sort = 4
    
End Sub
Public Sub MESRecipeAllshow()
On Error Resume Next

Dim Rownum As Integer
Dim i As Integer
    
    For i = 1 To iRecipeIDcount - 1
        frmMESRecipe.MSFlexGrid1.Rows = i + 1
        frmMESRecipe.MSFlexGrid1.TextMatrix(i, 0) = iRecipeIDcount - i
        frmMESRecipe.MSFlexGrid1.TextMatrix(i, 1) = sRecipeID(iRecipeIDcount - i)
        frmMESRecipe.MSFlexGrid1.TextMatrix(i, 2) = sRecipeComment(iRecipeIDcount - i)
    Next i
    If frmMESRecipe.MSFlexGrid1.Row >= 1 Then
        frmMESRecipe.MSFlexGrid1.Row = 1
        frmMESRecipe.MSFlexGrid1.Col = 0
        frmMESRecipe.MSFlexGrid1.Sort = 4
    End If
End Sub
Public Sub MESRecipeRecieve()
On Error GoTo err:
Dim Rownum As Integer

    Rownum = frmMESRecipe.MSFlexGrid1.Rows
    If frmMESRecipe.MSFlexGrid1.Rows >= 3001 Then
        frmMESRecipe.MSFlexGrid1.Clear
        frmMESRecipe.MSFlexGrid1.Rows = 1
        frmMESRecipe.MSFlexGrid1.Cols = 3
        frmMESRecipe.MSFlexGrid1.FormatString = "^No.    |" & "^RECIPE ID                |" & "^COMMENT                  "
    
        frmMESRecipe.MSFlexGrid1.RowHeight(0) = 500
        frmMESRecipe.MSFlexGrid1.ColWidth(0) = 1000
        frmMESRecipe.MSFlexGrid1.ColWidth(1) = 5100
        frmMESRecipe.MSFlexGrid1.ColWidth(2) = 3400
        Rownum = 1
        Rownum = Rownum + 1
    Else
        Rownum = Rownum + 1
    End If
    Call MESWriteGrid(iRecipeIDcount, Rownum)
Exit Sub
err:
End Sub

Public Sub MESRecipeShift(Index As Integer)
On Error GoTo err
Dim i As Integer
    For i = Index To iRecipeIDcount - 2

        sRecipeID(i) = sRecipeID(i + 1)
        sParamCount(i) = sParamCount(i + 1)
        iParamCount(i) = iParamCount(i + 1)
        sRecipeComment(i) = sRecipeComment(i + 1)
        For j = 1 To 10         'J�� �Ķ���� �ڵ�
            'sParamName_SV(i, j) = sParamName_SV(i + 1, j)
            sParamValue(i, j) = sParamValue(i + 1, j)
            sParamMinValue(i, j) = sParamMinValue(i + 1, j)
            sParamMaxValue(i, j) = sParamMaxValue(i + 1, j)
            
        Next j
    Next i
Exit Sub
err:

End Sub

Public Sub MESRecipeChange_OK()                      '������߰���
    sNowRecipeID = sBufRecipeID
    iNowRecipeID = iBufRecipeID
    Call DJ_MESMowRecipeSave
    frmMESRecipe.lblNowRecipeName.Caption = sNowRecipeID
    
    Call DJ_MESRecipeLoad(iNowRecipeID)
    For i = 0 To 2
        For j = 1 To iRecipeIDcount - 1
            frmMESRecipe.MSFlexGrid1.Col = i
            frmMESRecipe.MSFlexGrid1.Row = j
            frmMESRecipe.MSFlexGrid1.CellBackColor = vbWhite
        Next j
        frmMESRecipe.MSFlexGrid1.Col = i
        frmMESRecipe.MSFlexGrid1.Row = iRecipeIDcount - iNowRecipeID
        frmMESRecipe.MSFlexGrid1.CellBackColor = vbGreen
    Next i
End Sub
Public Sub MESRecipeChange_NG()                      '������߰���

    frmMESRecipe.lblNowRecipeName.Caption = sNowRecipeID
    
    Call DJ_MESRecipeLoad(iNowRecipeID)
    For i = 0 To 2
        For j = 1 To iRecipeIDcount - 1
            frmMESRecipe.MSFlexGrid1.Col = i
            frmMESRecipe.MSFlexGrid1.Row = j
            frmMESRecipe.MSFlexGrid1.CellBackColor = vbWhite
        Next j
        frmMESRecipe.MSFlexGrid1.Col = i
        frmMESRecipe.MSFlexGrid1.Row = iRecipeIDcount - iNowRecipeID
        frmMESRecipe.MSFlexGrid1.CellBackColor = vbGreen
    Next i
End Sub
