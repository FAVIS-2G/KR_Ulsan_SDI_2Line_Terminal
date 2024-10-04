Attribute VB_Name = "ModlLedControl_Multi_v20"
Option Explicit

'********************************************************************************************************
' Function  : Single & Muli �����Լ�
' Date      : 2009-07-05
' Eng'      : �̻���
' Desc.     : API Version UP (Multi Type)�� ���� ����
'********************************************************************************************************

'Function   : ���������(OPTLC_100SD) Ȱ��ȭ
'Return     : Success = 0~3, �ʱ�(Ȱ��ȭ)ȭ�� ������ ������ ����, Fail=-1
Public Declare Function OpenDAQDevice Lib "Pci_Pwm02.dll" () As Long
'���������(OPTLC_100SD) ����
Public Declare Function CloseDAQDevice Lib "Pci_Pwm02.dll" () As Boolean
'���������(OPTLC_100SD) ���� �ʱ�ȭ
Public Declare Function ResetBoard Lib "Pci_Pwm02.dll" (ByVal nBoard As Long) As Boolean

'********************************************************************************************************
' Function  : Multi Type v2.0
' Date      : 2009-07-05
' Eng'      : �̻���
' Desc.     : API Version UP (Multi Type)
'********************************************************************************************************

'������� ������ 0���� �ʱ�ȭ
'Parameter  : nBoard(Board Number), nCh(Channel Number)
'Return     : Success = True, Fail=False
Public Declare Function Pwm_Reset_Mul Lib "Pci_Pwm02.dll" (ByVal nBoard As Long, ByVal nCh As Long) As Boolean

'Strobe ��嵿�� ����, Pwm_Enable�� ������ ���¿��� ������
'Parameter  : nBoard(Board Number), nCh(Channel Number), nMode (0:Normal, Others:Tirgger Mode)
'Return     : Success = True, Fail=False
Public Declare Function Set_Mode_Mul Lib "Pci_Pwm02.dll" (ByVal nBoard As Long, ByVal nCh As Long, ByVal nMode As Long) As Boolean

'Strobe ��嵿�� ��������
Public Declare Function Get_Mode_Mul Lib "Pci_Pwm02.dll" (ByVal nBoard As Long, ByVal nCh As Long) As Long

'Strobe �������� ���� ��� �ù����̼� ���۹ݺ�
'�ݺ������� �ϱ� ���ؼ��� Strobe Signal �Է��� �־����
'Return     : Success = True, Fail=False
Public Declare Function Set_Cont_Mul Lib "Pci_Pwm02.dll" (ByVal nBoard As Long, ByVal nCh As Long, ByVal nCont As Long) As Boolean

'Strobe �������� ���� ��� �ù����̼� ��������
'Return     : 0=one-shot Mode, others=Continus Trigger Mode
Public Declare Function Get_Cont_Mul Lib "Pci_Pwm02.dll" (ByVal nBoard As Long, ByVal nCh As Long) As Long


'���� ��⼳��, nNum = 0~255
'Return : Success=True
Public Declare Function Set_Pwm_Mul Lib "Pci_Pwm02.dll" (ByVal nBoard As Long, ByVal nCh As Long, ByVal nNum As Long) As Boolean

'���� ��⼳���� ��������, 0~255
'Return : Success=0~255, Fail=-1
Public Declare Function Get_Pwm_Mul Lib "Pci_Pwm02.dll" (ByVal nBoard As Long, ByVal nCh As Long) As Long

'������� ON
'Return     : Success=True, Fail=False
Public Declare Function Pwm_Enable_Mul Lib "Pci_Pwm02.dll" (ByVal nBoard As Long, ByVal nCh As Integer) As Boolean

'������� OFF
'Return     : Success=True, Fail=False
Public Declare Function Pwm_Disable_Mul Lib "Pci_Pwm02.dll" (ByVal nBoard As Long, ByVal nCh As Integer) As Boolean

'Function   : Strobe Ʈ���� ������, ������� ������ �����ð������� �����ϱ�,
'Parameter  : nBoard=0~3, nCh=0~3, nTime=1~65,535 msec , 1 sec = 1,000 msec
'Return     : Success=True, Fail=False
Public Declare Function Set_Delay_Mul Lib "Pci_Pwm02.dll" (ByVal nBoard As Long, ByVal nCh As Long, ByVal nTime As Long) As Boolean

'Function   : Strobe Ʈ���� ������, ������� ������ �����ð������� ��������
'Parameter  : nBoard=0~3, nCh=0~3
'Return     : Success=1~65,535, Fail=-1
Public Declare Function Get_Delay_Mul Lib "Pci_Pwm02.dll" (ByVal nBoard As Long, ByVal nCh As Long) As Long

'Function   : Strobe ������� �ð� �����ϱ�
'Parameter  : nBoard=0~3, nCh=0~3, nTime=1~65,535 msec , 1 sec = 1,000 msec
'Return     : Success=True, Fail=False
Public Declare Function Set_Period_Mul Lib "Pci_Pwm02.dll" (ByVal nBoard As Long, ByVal nCh As Long, ByVal nTime As Long) As Boolean

'Function   : Strobe ������� �ð������� ��������
'Parameter  : nBoard=0~3, nCh=0~3
'Return     : Success=1~65,535, Fail=-1
Public Declare Function Get_Period_Mul Lib "Pci_Pwm02.dll" (ByVal nBoard As Long, ByVal nCh As Long) As Long

'Function   : DO(DigitalOutput) ����� Hex������ ����(���)
'Parameter  : nBoard=0~3, dout=Digital Write Value
'Return     : Success=True, Fail=False
Public Declare Function Set_Dout_Mul Lib "Pci_Pwm02.dll" (ByVal nBoard As Long, ByVal dout As Long) As Boolean

'DO(DigitalOutput) ��°� Hex������ �б�
'Return     : Success=Digital Read Value, Fail=-1
Public Declare Function Get_Dout_Mul Lib "Pci_Pwm02.dll" (ByVal nBoard As Long) As Long

'DI(DigitalInput) ��°� Hex������ �б�
Public Declare Function Get_Din_Mul Lib "Pci_Pwm02.dll" (ByVal nBoard As Long) As Long


'********************************************************************************************************
' Define
Public Const HIGH           As Integer = 1
Public Const LOW            As Integer = 0

Public gint_LC_BoardNo      As Integer


'***********************************************************************************
' �ۼ� Eng' : �̻���
' �ۼ� ���� : 2009-07-05
' ����      : OPTLC-100SDM v2.0(���������) ������
'-----------------------------------------------------------------------------------
' OPTLC100SDM_v2.0 (���������) DLL ��������
'
' ���ϸ�    : pci_pwm02.dll => ���������� �� On/Off
' ������  : pci_pwm02.dll ������ C:\Windows\system32 ������ ����
'           : SingleBoard DLL ���ϰ� ���ϸ��� �����ϴ� ��������� ���ǹٶ�



'***********************************************************************************
' �ۼ� Eng' : �̻���
' �ۼ� ���� : 2009-07-05
' ����      : ������ ��º� - MultiBoard
'-----------------------------------------------------------------------------------
' port      : ��½�ȣ �ε���(��Ʈ��ȣ), 0���� 15���� ����
' act       : ��� ON�ÿ��� HIGH, ��� OFF�ô� LOW�� ����
' ��뿹    : call OutPortOnOff (0, HIGH)      => 0����Ʈ ���ON
'             call OutPortOnOff (0, LOW)       => 0����Ʈ ���OFF
'***********************************************************************************
Public Sub OutPortOnOff_Mul(ByVal lBoardNo As Long, ByVal BitNo As Integer, ByVal Act As Integer)
On Error GoTo ErrorHandler

    Dim intGetOutPort       As Integer
    Dim intGetOutBit        As Integer
    Dim mBit                As Long
    Dim mVal                As Integer
    Dim mBuf                As Long
    Dim bRetSetDout         As Boolean
    
    intGetOutPort = Get_Dout_Mul(lBoardNo)
    
    
    If Act = HIGH Then
        mBit = (2 ^ BitNo)
        mVal = intGetOutPort Or mBit
        bRetSetDout = Set_Dout_Mul(lBoardNo, mVal)
        
    Else
        mBit = (2 ^ BitNo) * &H1
        mBuf = 65535 Xor mBit
        mVal = intGetOutPort And mBuf
        bRetSetDout = Set_Dout_Mul(lBoardNo, mVal)
    
    End If
    
    
Exit Sub
ErrorHandler:
    Debug.Print "~OutPortOnOff " & err.Description
    
End Sub

'***********************************************************************************
' �ۼ� Eng' : �̻���
' �ۼ� ���� : 2009-07-05
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
Public Function OutPortCheck_Mul(ByVal lBoardNo As Long, ByVal BitNo As Integer) As Boolean
On Error GoTo ErrorHandler
    Dim intGetOutPort       As Integer
    Dim intGetOutBit        As Integer
    Dim mBit                As Long
    Dim mVal                As Integer
    Dim mBuf                As Long
    
    intGetOutPort = Get_Dout_Mul(lBoardNo)
    
    mBit = (2 ^ BitNo)
    mVal = intGetOutPort And mBit
    
    If mVal = mBit Then
        OutPortCheck_Mul = True
    Else
        OutPortCheck_Mul = False
    End If
    
Exit Function
ErrorHandler:
    Debug.Print "~OutPortCheck " & err.Description
    
End Function



'***********************************************************************************
' �ۼ� Eng' : �̻���
' �ۼ� ���� : 2009-07-05
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
Public Function InPortCheck_Mul(ByVal lBoardNo As Long, ByVal BitNo As Integer) As Boolean
On Error GoTo ErrorHandler

    Dim lngGetInPort        As Long
    Dim lngGetInBit         As Long
    Dim mData               As Long
   
    lngGetInPort = Get_Din_Mul(lBoardNo)
    
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
        InPortCheck_Mul = True
    Else
        InPortCheck_Mul = False
    End If
   
    
Exit Function
ErrorHandler:
    Debug.Print "~InPortCheck " & err.Description
    
End Function


