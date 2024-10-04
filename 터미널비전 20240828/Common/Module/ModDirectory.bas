Attribute VB_Name = "ModDirectory"
Option Explicit

Public Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Secinfo As SECURITY_ATTRIBUTES
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Public Function Create_DIR(DirPathName As String) As Boolean
'----------------------------------------------
'directory create
'model info save folder make code
'model image save folder make code
On Error GoTo LPerr
Dim Secinfo As SECURITY_ATTRIBUTES
'Dim Dirname As String
Dim Rtn As Long
    
    Dim tempStrMsg As String
    Dim tempNstart As Integer
    Dim tempNend As Integer
    
    Dim Pos As Integer
    Dim EndPos As Integer
    Dim Path As String
    
    Pos = 0
    Do
        Pos = InStr(Pos + 1, DirPathName, "\")
        Path = Mid(DirPathName, 1, Pos)
        
        If Path = "" Then
            Path = DirPathName
        End If
        
        'Debug.Print Path
        If Dir(Path, vbDirectory) = "" Then
            Call CreateDirectory(Path, Secinfo)
        End If
    Loop Until Pos = 0
    
    Create_DIR = True
    
Exit Function

LPerr:
    Create_DIR = False

End Function

Public Function Model_Create_DIR(TempName As String) As Boolean
'----------------------------------------------
'directory create
'model info save folder make code
'model image save folder make code
On Error GoTo err
Dim Secinfo As SECURITY_ATTRIBUTES
Dim Dirname As String
Dim Rtn As Long
    '모델폴더 작성
    If Dir(App.Path & "\" & "model" & "\" & TempName, vbNormal) = "" Then
        Dirname = App.Path & "\" & "model" & "\" & TempName
        Rtn = CreateDirectory(Dirname, Secinfo)
            If Rtn = 0 Then
                Beep
            End If
    End If
    Model_Create_DIR = True
Exit Function
    
err:
    Model_Create_DIR = False
    MsgBox "현재 입력하신 모델 이름이 폴더를 생성하기에 적합하지 않습니다..!!", vbCritical, "폴더 이름 오류"
    
End Function

Public Sub Create_Date_DIR()
'----------------------------------------------
'directory create
'model info save folder make code
'model image save folder make code
On Error GoTo err
Dim Secinfo As SECURITY_ATTRIBUTES
Dim Dirname As String
Dim Rtn As Long
    If Dir(App.Path & "\Data\" & CStr(Date), vbNormal) = "" Then
        Dirname = App.Path & "\ResultData\" & CStr(Date)
        Rtn = CreateDirectory(Dirname, Secinfo)
'            If Rtn = 0 Then
'                Beep
'            End If
    End If
Exit Sub

err:

End Sub
