Attribute VB_Name = "ModHeader"
Public Declare Function GetProcessHeap Lib "kernel32" () As Long
Public Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Public Declare Function HeapDestroy Lib "kernel32" (ByVal hHeap As Long) As Boolean
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub keybd_event Lib "USER32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function SetParent Lib "USER32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'�� �����ϰ� �ϱ�======================================================
Public Const LWA_COLORKEY As Long = &H1
Public Const GWL_EXSTYLE As Long = -20
Public Const WS_EX_LAYERED  As Long = &H80000

Public Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crkey As Long, ByVal bAlpha As Byte, _
                        ByVal dwFlags As Long) As Long
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'======================================================================


'���� ��ư ������ ��� true - �ڵ� �˻簡 �Ϸ簡 ������ Ǯ���ٰ� �ؼ� ���� ��ư�� ���� ���� �ƴϸ�
'do loop �ʱ�� ���ư��� �ϱ����� ����
Public mb_IfstopBtnClked As Boolean

' bVk - �����Ű�ڵ尪..
' bScan - �ϵ���� ��ĵ�ڵ尪..
'        1 �̸� Active ȭ�鸸...
'        0 �̸� ȭ����ü��...
Public Const VK_SNAPSHOT = &H2C
Public Declare Function GetDiskFreeSpaceEx Lib "kernel32" _
  Alias "GetDiskFreeSpaceExA" _
  (ByVal lpRootPathName As String, _
  lpFreeBytesAvailableToCaller As Currency, _
  lpTotalNumberOfBytes As Currency, _
  lpTotalNumberOfFreeBytes As Currency) As Long

'Picture Box �� �� �׸��� ����
Public Enum SUBFORM_NAME
    nInspect = 100
    nLastNG
    nToolSetup
    nParameter
End Enum
Public g_NowForm As SUBFORM_NAME

Public Const pi = 3.1415926535

'ī�޶� �ػ�
Public Const XRES = 2560
Public Const YRES = 1920

Public Const MAX_CAMERA = 2

'ī�޶� �ػ�
'Public Const XRES = 2592
'Public Const YRES = 1944

'Cognex Library


'Favis Library
Public fvImageBuf(0 To kMaxCamera - 1) As Long
Public nImageRotation(0 To 3) As FvImageRotationConstants
'
'Edge Region (����ü�� ����� ������ ����ص� ������ Setting �� �Ǽ��� ���� ���尪�� �ٲ�°� �����ϱ� ���� ������ ����)
'double �� Long �� �ĺ� ���̺귯�� �޴����� ����� �������

Public dEdgeCenterX(0 To 3, 0 To 29) As Double
Public dEdgeCenterY(0 To 3, 0 To 29) As Double
Public lEdgeSideX(0 To 3, 0 To 29) As Long
Public lEdgeSideY(0 To 3, 0 To 29) As Long
Public dEdgeRotation(0 To 3, 0 To 29) As Double

Public dFixEdgeCenterX(0 To 3, 0 To 3) As Double
Public dFixEdgeCenterY(0 To 3, 0 To 3) As Double
Public lFixEdgeSideX(0 To 3, 0 To 3) As Long
Public lFixEdgeSideY(0 To 3, 0 To 3) As Long
Public dFixEdgeRotation(0 To 3, 0 To 3) As Double

Public dCalEdgeCenterX(0 To 3, 0 To 3) As Double
Public dCalEdgeCenterY(0 To 3, 0 To 3) As Double
Public lCalEdgeSideX(0 To 3, 0 To 3) As Long
Public lCalEdgeSideY(0 To 3, 0 To 3) As Long
Public dCalEdgeRotation(0 To 3, 0 To 3) As Double

'Blob Region
Public lBlobCenterX(0 To 3, 0 To 29) As Long
Public lBlobCenterY(0 To 3, 0 To 29) As Long
Public lBlobSideX(0 To 3, 0 To 29) As Long
Public lBlobSideY(0 To 3, 0 To 29) As Long
Public dBlobRotation(0 To 3, 0 To 29) As Double
'Region ȸ�� (0 : �ƹ��͵� ���� , 1 : �̸��극�̼� �̸��� , 2 : �Ƚ��� �̸��� , 3 : �˻� �̸���)
Public iRegionidx As Integer

'PLC �� MES ����
Public sPLCIP As String              'PLC IP
Public sMESIP As String              'MES IP
Public sPLCPort As String            'PLC Port
Public sMESPort As String            'MES Port
Public bWinsock(0 To 1) As Boolean   '��Ż��� boolean
Public sIDCode(0 To 4) As String     'PLC �� ���� �޴� ID �������
Public sIDCode_Q71(0 To 9) As Long  'Q71 ������ �ޱ�
Public str_sIDCode_Q71(0 To 9) As String  'Q71 ������ �ޱ�

Public bWriteReadPLC As Boolean      'True �� Write False �� read
Public sWriteReadPLC As String       '"I"�� read "O" write "D" Data read

'�˻���� �Լ�
Public bAutoRunOn As Boolean        '�ڵ��������� Ȯ���ϴ� boolean ���� ��
Public bLiveOn As Boolean           '���̺� �� ����
Public bTriggerOn As Boolean        'Ʈ���� On = true
Public bDArrival As Boolean         '������ ����̹� ��
Public bSpecChangeOn As Boolean     '����ü���� on = true
Public bJobChangeOn As Boolean     '����ü���� on = true
Public nPreJobNum As Integer           '���� �� ��ȣ

Public iProName As Integer          'Program Name ����
Public iCamNumber As Integer        'ī�޶� ���� ( 0 , 1)
Public iCamNumberS As Integer       'ī�޶� ����(����ȭ��) ( 0 , 1)
Public CamIdx As Integer

Public dCaliMM(0 To kMaxCamera - 1) As Double               '�̸��극�̼� mm ��
Public dCaliPX(0 To kMaxCamera - 1) As Double '�̸��극�̼� Pixel ��
Public dCaliPXY(0 To kMaxCamera - 1) As Double
Public lInspectionNum As Long                  '�˻�Count

Public bSettingManualRun As Boolean            'Setting â���� �޴��� �Ҷ� (true)

'Fixture ����
Public Fix_UseMode(0 To 3) As Integer                   '�Ƚ��� ��� ��� (0 : ������ , 1 : X ,Y �� ��� , 2 : ������ ���)

Public Fix_PointX(0 To 3) As Double                     '�Ƚ��� Origin ��ġ X
Public Fix_PointY(0 To 3) As Double                     '�Ƚ��� Origin ��ġ X
Public Fix_PointAngle(0 To 3) As Double                 '�Ƚ��� Origin ����

Public Fix_PointRunX(0 To 3) As Double                  '�Ƚ��� �˻� ��ġ X
Public Fix_PointRunY(0 To 3) As Double                  '�Ƚ��� �˻� ��ġ Y
Public Fix_PointAngleRunX(0 To 3, 0 To 1) As Double     '�Ƚ��� �˻� ��ġ X (������)
Public Fix_PointAngleRunY(0 To 3, 0 To 1) As Double     '�Ƚ��� �˻� ��ġ Y (������)

Public Fix_ShiftPointX(0 To 3) As Double                'shift ���Ѿ��� �Ƚ���X��
Public Fix_ShiftPointY(0 To 3) As Double                'shift ���Ѿ��� �Ƚ���Y��
Public Fix_ShiftPointAngle(0 To 3) As Double            'shift ���Ѿ��� ������

'Model ����
Public sModelName As String                   'Model Name
Public sNowModelName As String                'Now Model Name
Public sTemp_Modelname As String
Public bModelName As Boolean
Public iNowModelNo As Integer
Public sModelRoom(1 To 100) As String         '�𵨰���â���� �𵨹� ���̸� ����
Public sLastModelstr As String

'Tool Number
Public CNum As Integer        '�̸��� �ѹ�
Public BNum As Integer        '�� �ѹ�
Public iToolCount As Integer  '������ ����
Public iBlobToolCount As Integer    '�� ������ ����
Public ispecFalse As Integer  '���庰 �ҷ� ǥ�� (������ �������� ���� ���� �ҷ��̸� 1 �ƴϸ� 0)

'Tool Results
Public dCaliperX1(0 To 3, 0 To 29) As Double         'PosX1
Public dCaliperY1(0 To 3, 0 To 29) As Double         'PosY1
Public dCaliperX2(0 To 3, 0 To 29) As Double         'PosX2
Public dCaliperY2(0 To 3, 0 To 29) As Double         'PosY2
Public dCaliperCX(0 To 3, 0 To 29) As Double         'ResultX
Public dCaliperCY(0 To 3, 0 To 29) As Double         'ResultY
Public dBlobArea(0 To 3, 0 To 29) As Double
Public dTextPointX(0 To 3, 0 To 29) As Double        'ImageBox �� ����� ���ִ� POintX
Public dTextPointY(0 To 3, 0 To 29) As Double        'ImageBox �� ����� ���ִ� POintY

'ResultData �� ����
Public bResultJudge(0 To 3) As Boolean                          '��ü ����
Public bResultJudge_Spec(0 To 3, 0 To 19) As Boolean            '�׸� ����
Public iResultJudge_BlobCnt As Integer            '�� �׸� ����
Public bResultjudge_cnt As Integer
Public dInspectResult_mm(0 To 3, 0 To 19) As Double
Public dInspectResult_Pixel(0 To 3, 0 To 19) As Double

'GD
'20141021
Public bResultJudge_Blob(0 To 1) As Boolean
Public dResultJudge_BlobArea(0 To 1) As Double
Public dResultBlobArea(0 To 1) As Double

Public lOKCount As Long
Public lNGCount As Long
Public lToTalCount As Long

'Spec ����
Public sBlobName(0 To 29) As String       'SpecName
Public sSpecName(0 To 14) As String       'SpecName
Public bSpecPass(0 To 14) As Boolean      'Spec �� Pass    'True �� Pass

Public dSpecOriMin(0 To 14) As Double     'Ori - Min
Public dSpecOriMax(0 To 14) As Double     'Ori + Max
Public dSpecMin(0 To 14) As Double        '-����
Public dSpecMax(0 To 14) As Double        '+����
Public dSpecOri(0 To 14) As Double        '���ذ�
Public dSpecOffset(0 To 56) As Double     '�ɼ�
Public dCelluse(0 To 3) As Integer        'Cell����

Public bCamPass As Boolean                'True �� Pass
Public bOKimageSave As Boolean            'OK �̹��� ���� True
Public bNGimageSave As Boolean            'NG �̹��� ���� True
Public bWriteDataSave As Boolean          '������ ���� True

Public iImageFileMode As Integer           'Image ���� ���
Public OverHdd As Boolean                  '�ϵ� �뷮 �ʰ�

'��й�ȣ ����
Public sNewPassWord As String              '���ο� ��й�ȣ (����ȭ��)
Public sNowPassWord As String              '���� ��й�ȣ (����ȭ��)



'Melsec ����
Public nMelsecChannel As Integer
Public nMelsecMode As Integer

'Melsec �ּ� ����
Public lMelsecAddrInput As String                 'InputIO
Public lMelsecAddrOutput As String                'OutputIO
Public lMelsecAddrCellID As String                'CellID
Public sMelsecAddrCellID(0 To 9) As String
Public sMelsecAddrModelNumber As String             '�𵨹�ȣ
Public lMelsecAddrInspection(0 To 3, 0 To 9) As String  '�˻� ���
Public lMelsecAddrNgCode As String                '�ҷ��ڵ�
Public lMelsecAddrBaseSpec(0 To 10, 0 To 2) As String      '0: ���ذ�, 1: ���� ��, 2: ���� ��
Public lMelsecAddrCelluse(0 To 3) As String
Public sMelsecAddrAlarm As String
Public sMelsecAddrZigID As String

Public sMelsecAddrAcqDone As String
Public sMelsecAddrAutoRemove As String

Public sMelsecAddrAlarmHDD As String
Public sMelsecAddrAlarmCIM As String
Public sMelsecAddrAlarmNetDrive As String
Public sMelsecAddrAlarmCamera As String



Public sZigID As String

Public dScore As Double
Public dResultScore(0 To kMaxCamera - 1) As Double

Public nBlobNGCount As Integer

