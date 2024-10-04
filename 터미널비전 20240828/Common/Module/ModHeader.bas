Attribute VB_Name = "ModHeader"
Public Declare Function GetProcessHeap Lib "kernel32" () As Long
Public Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Public Declare Function HeapDestroy Lib "kernel32" (ByVal hHeap As Long) As Boolean
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Sub keybd_event Lib "USER32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function SetParent Lib "USER32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

'폼 투명하게 하기======================================================
Public Const LWA_COLORKEY As Long = &H1
Public Const GWL_EXSTYLE As Long = -20
Public Const WS_EX_LAYERED  As Long = &H80000

Public Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crkey As Long, ByVal bAlpha As Byte, _
                        ByVal dwFlags As Long) As Long
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'======================================================================


'정지 버튼 눌렀을 경우 true - 자동 검사가 하루가 지나면 풀린다고 해서 정지 버튼을 누른 것이 아니면
'do loop 초기로 돌아가게 하기위한 변수
Public mb_IfstopBtnClked As Boolean

' bVk - 버츄얼키코드값..
' bScan - 하드웨어 스캔코드값..
'        1 이면 Active 화면만...
'        0 이면 화면전체를...
Public Const VK_SNAPSHOT = &H2C
Public Declare Function GetDiskFreeSpaceEx Lib "kernel32" _
  Alias "GetDiskFreeSpaceExA" _
  (ByVal lpRootPathName As String, _
  lpFreeBytesAvailableToCaller As Currency, _
  lpTotalNumberOfBytes As Currency, _
  lpTotalNumberOfFreeBytes As Currency) As Long

'Picture Box 에 폼 그리기 관련
Public Enum SUBFORM_NAME
    nInspect = 100
    nLastNG
    nToolSetup
    nParameter
End Enum
Public g_NowForm As SUBFORM_NAME

Public Const pi = 3.1415926535

'카메라 해상도
Public Const XRES = 2560
Public Const YRES = 1920

Public Const MAX_CAMERA = 2

'카메라 해상도
'Public Const XRES = 2592
'Public Const YRES = 1944

'Cognex Library


'Favis Library
Public fvImageBuf(0 To kMaxCamera - 1) As Long
Public nImageRotation(0 To 3) As FvImageRotationConstants
'
'Edge Region (구조체에 저장된 값으로 사용해도 되지만 Setting 시 실수로 영역 저장값이 바뀌는걸 방지하기 위해 변수를 만듬)
'double 과 Long 은 파비스 라이브러리 메뉴얼을 참고로 만들었음

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
'Region 회전 (0 : 아무것도 안함 , 1 : 켈리브레이션 켈리퍼 , 2 : 픽스쳐 켈리퍼 , 3 : 검사 켈리퍼)
Public iRegionidx As Integer

'PLC 및 MES 관련
Public sPLCIP As String              'PLC IP
Public sMESIP As String              'MES IP
Public sPLCPort As String            'PLC Port
Public sMESPort As String            'MES Port
Public bWinsock(0 To 1) As Boolean   '통신상태 boolean
Public sIDCode(0 To 4) As String     'PLC 로 부터 받는 ID 저장공간
Public sIDCode_Q71(0 To 9) As Long  'Q71 데이터 받기
Public str_sIDCode_Q71(0 To 9) As String  'Q71 데이터 받기

Public bWriteReadPLC As Boolean      'True 면 Write False 면 read
Public sWriteReadPLC As String       '"I"는 read "O" write "D" Data read

'검사관련 함수
Public bAutoRunOn As Boolean        '자동상태인지 확인하는 boolean 변수 임
Public bLiveOn As Boolean           '라이브 블린 변수
Public bTriggerOn As Boolean        '트리거 On = true
Public bDArrival As Boolean         '데이터 어라이벌 블린
Public bSpecChangeOn As Boolean     '스펙체인지 on = true
Public bJobChangeOn As Boolean     '스펙체인지 on = true
Public nPreJobNum As Integer           '이전 잡 번호

Public iProName As Integer          'Program Name 변수
Public iCamNumber As Integer        '카메라 변수 ( 0 , 1)
Public iCamNumberS As Integer       '카메라 변수(세팅화면) ( 0 , 1)
Public CamIdx As Integer

Public dCaliMM(0 To kMaxCamera - 1) As Double               '켈리브레이션 mm 값
Public dCaliPX(0 To kMaxCamera - 1) As Double '켈리브레이션 Pixel 값
Public dCaliPXY(0 To kMaxCamera - 1) As Double
Public lInspectionNum As Long                  '검사Count

Public bSettingManualRun As Boolean            'Setting 창에서 메뉴얼런 할때 (true)

'Fixture 관련
Public Fix_UseMode(0 To 3) As Integer                   '픽스쳐 사용 모드 (0 : 사용안함 , 1 : X ,Y 만 사용 , 2 : 각도만 사용)

Public Fix_PointX(0 To 3) As Double                     '픽스쳐 Origin 위치 X
Public Fix_PointY(0 To 3) As Double                     '픽스쳐 Origin 위치 X
Public Fix_PointAngle(0 To 3) As Double                 '픽스쳐 Origin 각도

Public Fix_PointRunX(0 To 3) As Double                  '픽스쳐 검사 위치 X
Public Fix_PointRunY(0 To 3) As Double                  '픽스쳐 검사 위치 Y
Public Fix_PointAngleRunX(0 To 3, 0 To 1) As Double     '픽스쳐 검사 위치 X (각도용)
Public Fix_PointAngleRunY(0 To 3, 0 To 1) As Double     '픽스쳐 검사 위치 Y (각도용)

Public Fix_ShiftPointX(0 To 3) As Double                'shift 시켜야할 픽스쳐X값
Public Fix_ShiftPointY(0 To 3) As Double                'shift 시켜야할 픽스쳐Y값
Public Fix_ShiftPointAngle(0 To 3) As Double            'shift 시켜야할 각도값

'Model 관련
Public sModelName As String                   'Model Name
Public sNowModelName As String                'Now Model Name
Public sTemp_Modelname As String
Public bModelName As Boolean
Public iNowModelNo As Integer
Public sModelRoom(1 To 100) As String         '모델관리창에서 모델방 모델이름 저장
Public sLastModelstr As String

'Tool Number
Public CNum As Integer        '켈리퍼 넘버
Public BNum As Integer        '블럽 넘버
Public iToolCount As Integer  '툴개수 설정
Public iBlobToolCount As Integer    '블럽 툴개수 설정
Public ispecFalse As Integer  '스펙별 불량 표시 (툴개수 설정으로 인해 만듬 불량이면 1 아니면 0)

'Tool Results
Public dCaliperX1(0 To 3, 0 To 29) As Double         'PosX1
Public dCaliperY1(0 To 3, 0 To 29) As Double         'PosY1
Public dCaliperX2(0 To 3, 0 To 29) As Double         'PosX2
Public dCaliperY2(0 To 3, 0 To 29) As Double         'PosY2
Public dCaliperCX(0 To 3, 0 To 29) As Double         'ResultX
Public dCaliperCY(0 To 3, 0 To 29) As Double         'ResultY
Public dBlobArea(0 To 3, 0 To 29) As Double
Public dTextPointX(0 To 3, 0 To 29) As Double        'ImageBox 에 결과값 써주는 POintX
Public dTextPointY(0 To 3, 0 To 29) As Double        'ImageBox 에 결과값 써주는 POintY

'ResultData 및 판정
Public bResultJudge(0 To 3) As Boolean                          '전체 판정
Public bResultJudge_Spec(0 To 3, 0 To 19) As Boolean            '항목별 판정
Public iResultJudge_BlobCnt As Integer            '블럽 항목별 판정
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

'Spec 관련
Public sBlobName(0 To 29) As String       'SpecName
Public sSpecName(0 To 14) As String       'SpecName
Public bSpecPass(0 To 14) As Boolean      'Spec 별 Pass    'True 면 Pass

Public dSpecOriMin(0 To 14) As Double     'Ori - Min
Public dSpecOriMax(0 To 14) As Double     'Ori + Max
Public dSpecMin(0 To 14) As Double        '-공차
Public dSpecMax(0 To 14) As Double        '+공차
Public dSpecOri(0 To 14) As Double        '기준값
Public dSpecOffset(0 To 56) As Double     '옵셋
Public dCelluse(0 To 3) As Integer        'Cell유무

Public bCamPass As Boolean                'True 면 Pass
Public bOKimageSave As Boolean            'OK 이미지 저장 True
Public bNGimageSave As Boolean            'NG 이미지 저장 True
Public bWriteDataSave As Boolean          '데이터 저장 True

Public iImageFileMode As Integer           'Image 저장 모드
Public OverHdd As Boolean                  '하드 용량 초과

'비밀번호 관련
Public sNewPassWord As String              '새로운 비밀번호 (메인화면)
Public sNowPassWord As String              '현재 비밀번호 (메인화면)



'Melsec 연결
Public nMelsecChannel As Integer
Public nMelsecMode As Integer

'Melsec 주소 번지
Public lMelsecAddrInput As String                 'InputIO
Public lMelsecAddrOutput As String                'OutputIO
Public lMelsecAddrCellID As String                'CellID
Public sMelsecAddrCellID(0 To 9) As String
Public sMelsecAddrModelNumber As String             '모델번호
Public lMelsecAddrInspection(0 To 3, 0 To 9) As String  '검사 결과
Public lMelsecAddrNgCode As String                '불량코드
Public lMelsecAddrBaseSpec(0 To 10, 0 To 2) As String      '0: 기준값, 1: 상한 값, 2: 하한 값
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

