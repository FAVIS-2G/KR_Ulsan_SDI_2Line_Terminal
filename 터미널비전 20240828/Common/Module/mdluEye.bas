Attribute VB_Name = "mdluEye"
Option Explicit

Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long


Public m_PixelMemory As Long ' Address of allocated pixel memoryac




'Public Function AllocateNew(ByVal Width As Long, ByVal Height As Long)
'
'On Error GoTo ErrHandler
'
'    If m_PixelMemory <> 0 Then
'      'err.Raise -1, , "Memory is already allocated"
'      Exit Function
'    End If
'
'    ' Allocate a block of memory from a heap.
'    m_PixelMemory = HeapAlloc(GetProcessHeap, 0, Width * Height)
'
'    Exit Function
'
'ErrHandler:
'
'
'End Function


' This allocates memory for storing pixels and returns a new
' CogImage8Root object
Public Function Allocate(ByVal Width As Long, ByVal Height As Long) As CogImage8Root

On Error GoTo ErrHandler

    Dim tme As IDisposable

    ' If memory is already allocated, it's an error. You need to create
    ' a new instance of this class each time you want a new block of memory.
    '''''  If m_PixelMemory <> 0 Then
    '''''    Err.Raise -1, , "Memory is already allocated"
    '''''  End If
    '''''
    '''''  ' Allocate a block of memory from a heap.
    '''''  m_PixelMemory = HeapAlloc(GetProcessHeap, 0, Width * Height)
    
    ' Create a root buffer to pass to CogImage8Grey.SetRoot
    ' Buffer will hold raw 8-bit pixel data of an image.
    Dim Buffer As ICogImage8RootBuffer
    Set Buffer = New CogImage8Root
    
    m_PixelMemory = frmMain.uEyeCam.GetImageMem
    
    'm_PixelMemory(CamNum) = frmMain.uEyeCam1(0).GetImageMem
      
    ' Intialize the buffer, giving it the image dimensions and a reference
    ' back to this object so it can call Dispose when it's done with
    ' the pixel memory.
    
    Buffer.Initialize Width, Height, m_PixelMemory, Width, tme
    
    ' Return the buffer to the caller. Note that it's important NOT to store
    ' a reference to the Buffer in this class, because the Buffer already
    ' has a reference to this class's IDisposable interface. If a reference
    ' to the Buffer were stored in this class it would create a circular
    ' reference and the two objects would never get freed up.
    Set Allocate = Buffer
    Set Buffer = Nothing
    
    Exit Function

ErrHandler:
    MsgBox "Memory 할당 에러", vbCritical
    
End Function




' Free the pixel memory if it has been allocated
Public Function FreeAllocatedMemory(Index As Integer)

On Error GoTo ErrHandler

    If m_PixelMemory <> 0 Then
        HeapFree GetProcessHeap, 0, m_PixelMemory
        m_PixelMemory = 0
    End If
    
    Exit Function
    
ErrHandler:

End Function

' This function gets called when the last reference to the CogImage8Root
' is released.
Public Function IDisposable_Dispose()

On Error Resume Next

    FreeAllocatedMemory 1
    FreeAllocatedMemory 2
    
End Function

' Normally the memory will already have been freed up by IDisposable_Dispose
' before Class_Terminate runs, but just in case something went wrong this
' will check again and free the memory if necessary.
Public Function Camera_Memory_Terminate()

On Error GoTo ErrHandler
    
    FreeAllocatedMemory 1

    Exit Function
    
ErrHandler:

End Function



Public Function pfunc_Acqurie() As CogImage8Grey
On Error GoTo ErrHandler

    Dim IdsImage    As New CogImage8Grey
'
'    frmMain.cogDisp.StaticGraphics.Clear
'    frmMain.cogDisp.InteractiveGraphics.Clear

'''    Set favImage(Index) = Nothing
'''    Set favImage(Index) = New CogImage8Grey

    If IS_SUCCESS = frmMain.uEyeCam.FreezeImage(IS_WAIT) Then
        IdsImage.SetRoot Allocate(1280, 512)
        Set pfunc_Acqurie = IdsImage.Copy
    Else
        Set pfunc_Acqurie = Nothing
        
    End If
    
Exit Function
ErrHandler:
    Set pfunc_Acqurie = Nothing

End Function




Public Function pfuncb_uEye_Init(Optional vOcxNo As Integer) As Boolean
On Error GoTo ErrorHandler

    Dim lngCamReturn            As Long
    Dim lngTempPixelClock       As Long
    Dim dblTempFrameRate        As Double
    Dim dblTempExposure         As Double
    Dim strTempCameraType       As String

    Call AllocateNew(1280, 512)
    If frmMain.uEyeCam.InitCamera(1) = IS_SUCCESS Then
        Call psub_LogWriteTxt("Operation", "Success Camera Init (CamID=1)")
'    If lngCamReturn <> 3 Then
'        lngTempPixelClock = frmMain.uEyeCam.GetPixelClock
'        dblTempFrameRate = frmMain.uEyeCam.GetFrameRate
'        dblTempExposure = frmMain.uEyeCam.GetExposureTime
'
'        frmMain.uEyeCam.SetPixelClock lngTempPixelClock
'        frmMain.uEyeCam.SetFrameRate dblTempFrameRate
'        frmMain.uEyeCam.SetExposureTime dblTempExposure
'
'        frmMain.uEyeCam.AllowPopupMenu = True
'        strTempCameraType = frmMain.uEyeCam.GetCameraTyp
'
'    End If
    
    Else
        pfuncb_uEye_Init = False
        
    End If
    
    pfuncb_uEye_Init = True
    
Exit Function
ErrorHandler:
    pfuncb_uEye_Init = False
    Call psub_LogWriteTxt("Operation", "~uEye_Init " & err.Description)
    
End Function


Public Function pfunc_LoadCameraParamDefault() As Boolean
On Error GoTo ErrorHandler

    Dim FilePath        As String
    Dim strResult       As String * 1024
    Dim lngApiReturn    As Long
    Dim szSection       As String
    Dim i               As Integer

    'FilePath = App.Path & "\CONFIG\" & Current_Model & "\Camera.ini"
    FilePath = App.Path & "\CONFIG\" & "\Camera.ini"
    
    pfunc_LoadCameraParamDefault = True
    
Exit Function
ErrorHandler:
    pfunc_LoadCameraParamDefault = False

End Function

    

