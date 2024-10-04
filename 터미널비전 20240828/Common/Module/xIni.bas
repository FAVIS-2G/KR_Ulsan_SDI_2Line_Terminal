Attribute VB_Name = "xIni"
Option Explicit

' ini 颇老 包访 Library
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

'#######################
'ini 颇老 Reading
'#######################
Public Function ReadINI(iniFile As String, Section As String, Key As String, Optional default As String = "") As String

    Dim retVal As String
    Dim Worked As Integer
    Dim arrip() As String
    Dim repstConv As String
    
    retVal = String(500, 0)
    
    Worked = GetPrivateProfileString(UCase(Section), UCase(Key), default, retVal, Len(retVal), iniFile)
    
    repstConv = Replace(Left(retVal, Worked), Chr(0), "")
    
    ReadINI = repstConv
    
End Function

'##################################
'ini 颇老 Writing
'##################################
Public Sub WriteINI(iniFile As String, Section As String, Key As String, W_KEY As String)
    Dim Worked As Integer
    Dim dblText As String
    
    'AppName = App.Path & "\" & iniFile 'info.ini"
    
    Worked = WritePrivateProfileString(UCase(Section), UCase(Key), W_KEY, iniFile)
End Sub


