Attribute VB_Name = "commonAmpApi"
        
Option Explicit

'���������(OPTLC_100SD) Ȱ��ȭ
'�ʱ�(Ȱ��ȭ)ȭ�� ������ ������ ������ 0,1,2... ������ ������
'�����ϸ� -1 �� ����
Public Declare Function OpenDAQDevice Lib "Pci_Pwm02.dll" () As Long
'���������(OPTLC_100SD) ����
Public Declare Function CloseDAQDevice Lib "Pci_Pwm02.dll" () As Boolean
'���������(OPTLC_100SD) ���� �ʱ�ȭ
Public Declare Function ResetBoard Lib "Pci_Pwm02.dll" (ByVal nBoard As Long) As Boolean

'������� ������ 0���� �ʱ�ȭ
Public Declare Function Pwm_Reset Lib "Pci_Pwm02.dll" (ByVal nCh As Long) As Boolean

'Strobe ��嵿�� ����, Pwm_Enable�� ������ ���¿��� ������
Public Declare Function Set_Mode Lib "Pci_Pwm02.dll" (ByVal nCh As Long, ByVal nMode As Long) As Boolean
'Strobe ��嵿�� ��������
Public Declare Function Get_Mode Lib "Pci_Pwm02.dll" (ByVal nCh As Long) As Long

'Strobe �������� ���� ��� �ù����̼� ���۹ݺ�
'�ݺ������� �ϱ� ���ؼ��� Strobe Signal �Է��� �־����
Public Declare Function Set_Cont Lib "Pci_Pwm02.dll" (ByVal nCh As Long, ByVal nCont As Long) As Boolean
'Strobe �������� ���� ��� �ù����̼� ��������
Public Declare Function Get_Cont Lib "Pci_Pwm02.dll" (ByVal nCh As Long) As Long


'���� ��⼳��, 0~255
Public Declare Function Set_Pwm Lib "Pci_Pwm02.dll" (ByVal nCh As Long, ByVal nNum As Long) As Boolean
'���� ��⼳���� �ϱ�, 0~255
Public Declare Function Get_Pwm Lib "Pci_Pwm02.dll" (ByVal nCh As Long) As Long

'������� ON
Public Declare Function Pwm_Enable Lib "Pci_Pwm02.dll" (ByVal nCh As Integer) As Boolean
'������� OFF
Public Declare Function Pwm_Disable Lib "Pci_Pwm02.dll" (ByVal nCh As Integer) As Boolean

'Strobe Ʈ���� ������, ������� ������ �����ð������� �����ϱ�,
'�Է¹��� : 1~1000000 usec(����ũ�μ�ũ), 1��=1000000 usec
Public Declare Function Set_Delay Lib "Pci_Pwm02.dll" (ByVal nCh As Long, ByVal nTime As Long) As Boolean
'Strobe Ʈ���� ������, ������� ������ �����ð������� ��������
'�Է¹��� : 1~1000000 usec(����ũ�μ�ũ), 1��=1000000 usec
Public Declare Function Get_Delay Lib "Pci_Pwm02.dll" (ByVal nCh As Long) As Long

'Strobe ������� �ð� �����ϱ�
Public Declare Function Set_Period Lib "Pci_Pwm02.dll" (ByVal nCh As Long, ByVal nTime As Long) As Boolean
'Strobe ������� �ð������� ��������
Public Declare Function Get_Period Lib "Pci_Pwm02.dll" (ByVal nCh As Long) As Long

'DO(DigitalOutput) ����� Hex������ ����(���)
Public Declare Function Set_Dout Lib "Pci_Pwm02.dll" (ByVal dout As Long) As Boolean
'DO(DigitalOutput) ��°� Hex������ �б�
Public Declare Function Get_Dout Lib "Pci_Pwm02.dll" () As Long
'DI(DigitalInput) ����� Hex������ ����(���)
Public Declare Function Set_Din Lib "Pci_Pwm02.dll" () As Long
'DI(DigitalInput) ��°� Hex������ �б�
Public Declare Function Get_Din Lib "Pci_Pwm02.dll" () As Long


Public Const HIGH = 1
Public Const LOW = 0


'***********************************************************************************
' �ۼ� Eng' : �̻���
' �ۼ� ���� : 2007 / 10 / 18
' ����      : OPTLC-100SD(���������) ������
'-----------------------------------------------------------------------------------
' OPTLC100SD (���������) DLL ��������
'
' ���ϸ�    : pci_pwm02.dll => ���������� �� On/Off
' ������  : pci_pwm02.dll ������ C:\Windows\system32 ������ ����



'***********************************************************************************
' �ۼ� Eng' : �̻���
' �ۼ� ���� : 2007 / 10 / 18
' ����      : ������ ��º�
'-----------------------------------------------------------------------------------
' port      : ��½�ȣ �ε���(��Ʈ��ȣ), 0���� 15���� ����
' act       : ��� ON�ÿ��� HIGH, ��� OFF�ô� LOW�� ����
' ��뿹    : OutPortOnOff 0, HIGH      => 0����Ʈ ���ON
'             OutPortOnOff 0, LOW       => 0����Ʈ ���OFF
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
' �ۼ� Eng' : �̻���
' �ۼ� ���� : 2007 / 10 / 18
' ����      : ������ ���Ȯ��
'-----------------------------------------------------------------------------------
' BitNo     : ���Ȯ�� ��Ʈ��ȣ, 0���� 11���� ����
' Return    : �ش��Ʈ�� ����� �ִٸ� True�� ���ϵ�
' ��뿹    :
'                If OutPortCheck(0) = True Then
'                    MsgBox "0�� ��Ʈ ��½�ȣ�� �����Ǿ����ϴ�."
'                Else
'                    MsgBox "0�� ��Ʈ ��½�ȣ�� �����ϴ�."
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
' �ۼ� Eng' : �̻���
' �ۼ� ���� : 2007 / 10 / 18
' ����      : ������ �Էº�
'-----------------------------------------------------------------------------------
' BitNo     : �Է�Ȯ�� ��Ʈ��ȣ, 0���� 6���� ����
' Return    : �ش��Ʈ�� �Է��� �ִٸ� True�� ���ϵ�
' ��뿹    :
'                If InPortCheck(0) = True Then
'                    MsgBox "0�� ��Ʈ�� �Է½�ȣ�� �����Ǿ����ϴ�."
'                Else
'                    MsgBox "0�� ��Ʈ �Է½�ȣ�� �����ϴ�."
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

