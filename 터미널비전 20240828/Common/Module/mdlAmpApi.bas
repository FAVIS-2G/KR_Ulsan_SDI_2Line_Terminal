Attribute VB_Name = "commonAmpApi"
        
Option Explicit

'조명제어보드(OPTLC_100SD) 활성화
'초기(활성화)화에 성공한 보드의 개수를 0,1,2... 등으로 리턴함
'실패하면 -1 을 리턴
Public Declare Function OpenDAQDevice Lib "Pci_Pwm02.dll" () As Long
'조명제어보드(OPTLC_100SD) 해제
Public Declare Function CloseDAQDevice Lib "Pci_Pwm02.dll" () As Boolean
'조명제어보드(OPTLC_100SD) 보드 초기화
Public Declare Function ResetBoard Lib "Pci_Pwm02.dll" (ByVal nBoard As Long) As Boolean

'조명관련 설정값 0으로 초기화
Public Declare Function Pwm_Reset Lib "Pci_Pwm02.dll" (ByVal nCh As Long) As Boolean

'Strobe 모드동작 설정, Pwm_Enable이 설정된 상태에서 동작함
Public Declare Function Set_Mode Lib "Pci_Pwm02.dll" (ByVal nCh As Long, ByVal nMode As Long) As Boolean
'Strobe 모드동작 설정해제
Public Declare Function Get_Mode Lib "Pci_Pwm02.dll" (ByVal nCh As Long) As Long

'Strobe 설정값에 따른 출력 시물레이션 동작반복
'반복동작을 하기 위해서는 Strobe Signal 입력이 있어야함
Public Declare Function Set_Cont Lib "Pci_Pwm02.dll" (ByVal nCh As Long, ByVal nCont As Long) As Boolean
'Strobe 설정값에 따른 출력 시물레이션 동작해제
Public Declare Function Get_Cont Lib "Pci_Pwm02.dll" (ByVal nCh As Long) As Long


'조명 밝기설정, 0~255
Public Declare Function Set_Pwm Lib "Pci_Pwm02.dll" (ByVal nCh As Long, ByVal nNum As Long) As Boolean
'조명 밝기설정값 일기, 0~255
Public Declare Function Get_Pwm Lib "Pci_Pwm02.dll" (ByVal nCh As Long) As Long

'조명출력 ON
Public Declare Function Pwm_Enable Lib "Pci_Pwm02.dll" (ByVal nCh As Integer) As Boolean
'조명출력 OFF
Public Declare Function Pwm_Disable Lib "Pci_Pwm02.dll" (ByVal nCh As Integer) As Boolean

'Strobe 트리거 수신후, 조명출력 까지의 지연시간설정값 설정하기,
'입력범위 : 1~1000000 usec(마이크로세크), 1초=1000000 usec
Public Declare Function Set_Delay Lib "Pci_Pwm02.dll" (ByVal nCh As Long, ByVal nTime As Long) As Boolean
'Strobe 트리거 수신후, 조명출력 까지의 지연시간설정값 가져오기
'입력범위 : 1~1000000 usec(마이크로세크), 1초=1000000 usec
Public Declare Function Get_Delay Lib "Pci_Pwm02.dll" (ByVal nCh As Long) As Long

'Strobe 조명출력 시간 설정하기
Public Declare Function Set_Period Lib "Pci_Pwm02.dll" (ByVal nCh As Long, ByVal nTime As Long) As Boolean
'Strobe 조명출력 시간설정값 가져오기
Public Declare Function Get_Period Lib "Pci_Pwm02.dll" (ByVal nCh As Long) As Long

'DO(DigitalOutput) 출력을 Hex값으로 설정(출력)
Public Declare Function Set_Dout Lib "Pci_Pwm02.dll" (ByVal dout As Long) As Boolean
'DO(DigitalOutput) 출력값 Hex값으로 읽기
Public Declare Function Get_Dout Lib "Pci_Pwm02.dll" () As Long
'DI(DigitalInput) 출력을 Hex값으로 설정(출력)
Public Declare Function Set_Din Lib "Pci_Pwm02.dll" () As Long
'DI(DigitalInput) 출력값 Hex값으로 읽기
Public Declare Function Get_Din Lib "Pci_Pwm02.dll" () As Long


Public Const HIGH = 1
Public Const LOW = 0


'***********************************************************************************
' 작성 Eng' : 이상훈
' 작성 일자 : 2007 / 10 / 18
' 내용      : OPTLC-100SD(조명제어보드) 적용방법
'-----------------------------------------------------------------------------------
' OPTLC100SD (조명제어보드) DLL 파일적용
'
' 파일명    : pci_pwm02.dll => 조명밝기조절 및 On/Off
' 적용방법  : pci_pwm02.dll 파일을 C:\Windows\system32 폴더에 복사



'***********************************************************************************
' 작성 Eng' : 이상훈
' 작성 일자 : 2007 / 10 / 18
' 내용      : 디지털 출력부
'-----------------------------------------------------------------------------------
' port      : 출력신호 인덱스(비트번호), 0부터 15까지 제공
' act       : 출력 ON시에는 HIGH, 출력 OFF시는 LOW를 정의
' 사용예    : OutPortOnOff 0, HIGH      => 0번비트 출력ON
'             OutPortOnOff 0, LOW       => 0번비트 출력OFF
'***********************************************************************************
Public Sub OutPortOnOff(ByVal BitNo As Integer, ByVal Act As Integer)
On Error GoTo ErrorHandler

    Dim intGetOutPort       As Integer
    Dim intGetOutBit        As Integer
    Dim mBit                As Long
    Dim mVal                As Integer
    Dim mBuf                As Long
    
    intGetOutPort = Get_Dout()
    
    
    If Act = HIGH Then
        mBit = (2 ^ BitNo)
        mVal = intGetOutPort Or mBit
        Set_Dout (mVal)
        
    Else
        mBit = (2 ^ BitNo) * &H1
        mBuf = 65535 Xor mBit
        mVal = intGetOutPort And mBuf
        Set_Dout (mVal)
    
    End If
    
    
Exit Sub
ErrorHandler:
    Debug.Print "~OutPortOnOff " & err.Description
    
End Sub

'***********************************************************************************
' 작성 Eng' : 이상훈
' 작성 일자 : 2007 / 10 / 18
' 내용      : 디지털 출력확인
'-----------------------------------------------------------------------------------
' BitNo     : 출력확인 비트번호, 0부터 11까지 제공
' Return    : 해당비트의 출력이 있다면 True가 리턴됨
' 사용예    :
'                If OutPortCheck(0) = True Then
'                    MsgBox "0번 비트 출력신호가 감지되었습니다."
'                Else
'                    MsgBox "0번 비트 출력신호가 없습니다."
'                End If
'
'***********************************************************************************
Public Function OutPortCheck(ByVal BitNo As Integer) As Boolean
On Error GoTo ErrorHandler
    Dim intGetOutPort       As Integer
    Dim intGetOutBit        As Integer
    Dim mBit                As Long
    Dim mVal                As Integer
    Dim mBuf                As Long
    intGetOutPort = Get_Dout()
    
    mBit = (2 ^ BitNo)
    mVal = intGetOutPort And mBit
    
    If mVal = mBit Then
        OutPortCheck = True
    Else
        OutPortCheck = False
    End If
    
Exit Function
ErrorHandler:
    Debug.Print "~OutPortCheck " & err.Description
    
End Function



'***********************************************************************************
' 작성 Eng' : 이상훈
' 작성 일자 : 2007 / 10 / 18
' 내용      : 디지털 입력부
'-----------------------------------------------------------------------------------
' BitNo     : 입력확인 비트번호, 0부터 6까지 제공
' Return    : 해당비트의 입력이 있다면 True가 리턴됨
' 사용예    :
'                If InPortCheck(0) = True Then
'                    MsgBox "0번 비트로 입력신호가 감지되었습니다."
'                Else
'                    MsgBox "0번 비트 입력신호가 없습니다."
'                End If
'
'***********************************************************************************
Public Function InPortCheck(ByVal BitNo As Integer) As Boolean
On Error GoTo ErrorHandler

    Dim lngGetInPort        As Long
    Dim lngGetInBit         As Long
    Dim mData               As Long
   
    lngGetInPort = Get_Din()
    
    Select Case BitNo
        Case 0
            mData = 1
        Case 1
            mData = 2
        Case 2
            mData = 4
        Case 3
            mData = 8
        Case 4
            mData = 16
        Case 5
            mData = 32
    End Select

    lngGetInBit = lngGetInPort And mData

    If mData = lngGetInBit Then
        InPortCheck = True
    Else
        InPortCheck = False
    End If
   
    
Exit Function
ErrorHandler:
    Debug.Print "~InPortCheck " & err.Description
    
End Function

