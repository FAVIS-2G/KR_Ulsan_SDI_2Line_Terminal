VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BaslerCamera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Exposure As Long
Public gain As Long
Public m_nID As Long

' The handle of the currently used pylon device.
Public m_hDevice As PYLON_DEVICE_HANDLE
Public DeviceIndex As Integer
Public SelectDevice As Integer

Public m_hCallback As PYLON_DEVICECALLBACK_HANDLE



Public Function Initialize() As Boolean
On Error GoTo err

Dim hDeviceInfo As PYLON_DEVICE_INFO_HANDLE
Dim numDevices As Long
Dim i As Long


    m_hDevice = cPylonInvalidHandle
    m_hCallback = cPylonInvalidHandle
    
    
    
    ' Enumerate all devices attached to this system.

    
    numDevices = PylonEnumerateDevices

    ' Add each found device to the listbox.
    For i = 0 To numDevices - 1
        Dim name As String

        hDeviceInfo = cPylonInvalidHandle

        ' Get the handle to the device info object.
        hDeviceInfo = PylonGetDeviceInfoHandle(i)

        ' Retrieve ModelName.
        name = PylonDeviceInfoGetPropertyValueByName(hDeviceInfo, cPylonDeviceInfoUserDefinedNameKey)

        If CLng(name - 2) = m_nID Then

            DeviceIndex = m_nID
            Exit For

        End If

    Next i

    m_hDevice = PylonCreateDeviceByIndex(DeviceIndex)

    If m_hDevice = 0 Then
        MsgBox "Cam Init failed!!!"
        Exit Function
    End If
    
    ' Open the device.
    PylonDeviceOpen m_hDevice, (cPylonAccessModeExclusive Or cPylonAccessModeStream Or cPylonAccessModeControl)
    ' Extend the heartbeat timeout to make debuggig easier.
    'SetHeartbeatTimeoutFromCommandLine m_hDevice
    ' Set up the camera (will raise an error if the camera isn't suitable).
    ConfigureDevice m_hDevice

    Initialize = True
Exit Function
err:
    Initialize = False
    
End Function

Public Sub Release()
    
    If (m_hDevice <> cPylonInvalidHandle) Then
        
        ' If something goes wrong
        ' we cannot do anything about it.
        On Error Resume Next
        
        ' Deregister the removal callback.
        If m_hCallback <> cPylonInvalidHandle Then
            PylonDeviceDeregisterRemovalCallbackHwnd m_hDevice, m_hCallback
            m_hCallback = cPylonInvalidHandle
        End If
        
        ' Close and free the device.
        If (PylonDeviceIsOpen(m_hDevice)) Then
            PylonDeviceClose m_hDevice
        End If
        
        PylonDestroyDevice (m_hDevice)
        m_hDevice = cPylonInvalidHandle
        
    End If
 
    ' Uninitialize pylon.
    PylonTerminate
    
End Sub

Public Function ImageGrab() As Boolean
On Error GoTo err
' When grabbing a single frame we use only one buffer.
Dim OneBuffer() As Byte
' It will be filled with information about the grabbed image data.
Dim GrabResult As PylonGrabResult_t

      Dim tmpWidth As Long
    Dim tmpHeight As Long
    Dim tmp As Long
    
    tmpWidth = PylonDeviceGetIntegerFeature(m_hDevice, "Width")
    tmpHeight = PylonDeviceGetIntegerFeature(m_hDevice, "Height")
    DoEvents
    If PylonDeviceGrabSingleFrame(m_hDevice, 0, OneBuffer, GrabResult, 200) Then
        DoEvents
        If GrabResult.status = EPylonGrabStatus.Grabbed Then
            'MsgBox "±×·¦¼º°ø", vbInformation
            'CopyMemory m_PixelMemory(m_nID), GetAddrOf(OneBuffer(0)), tmpWidth * tmpHeight
            Sleep 1
            '//// cognex
            Dim Image As New CogImage8Grey

           ' fvImageBuf(m_nID) = GetAddrOf(OneBuffer(0))
            m_favFlip.ClearOperation
            m_favFlip.InputImage = GetAddrOf(OneBuffer(0))
            m_favFlip.ImageWidth = XRES
            m_favFlip.ImageHeight = YRES
            m_favFlip.Run
            m_PixelMemory(m_nID) = m_favFlip.OutputImage
            'm_PixelMemory(m_nID) = GetAddrOf(OneBuffer(0))
            fvImageBuf(m_nID) = m_PixelMemory(m_nID)
            
            frmMain.FavisImageBoxMain(m_nID).DeleteAll
            frmMain.FavisImageBoxMain(m_nID).DrawImage fvImageBuf(m_nID)
            frmMain.FavisImageBoxMain(m_nID).UpdateDraw
'            Call C8Grey_Acquire(m_nID, m_PixelMemory(m_nID))
'            Set m_Image(m_nID) = m_Image(m_nID).Copy(cogImageCopyModeCopyPixels)
'            Set frmSetting.cogSettingDisplay(0).Image = m_Image(m_nID)
    
            'Set g_CogImageBuffer(m_nID) = Nothing
            
            ImageGrab = True
        ElseIf (GrabResult.status = EPylonGrabStatus.Failed) Then
        
            ' Grab error. Display the error code from the GrabResult.
            'MsgBox "Failed to grab image. Errorcode = " & Hex(GrabResult.ErrorCode), vbExclamation
            ImageGrab = False
        End If
        
    Else
        ImageGrab = False
        m_PixelMemory(m_nID) = 0
        'MsgBox "Timeout while acquiring the image.", vbExclamation
    End If

    
Exit Function
err:
    ImageGrab = False
    m_PixelMemory(m_nID) = 0
End Function
Public Function ImageGrabSetting() As Boolean
On Error GoTo err
' When grabbing a single frame we use only one buffer.
Dim OneBuffer() As Byte
' It will be filled with information about the grabbed image data.
Dim GrabResult As PylonGrabResult_t

    Dim tmpWidth As Long
    Dim tmpHeight As Long
    Dim tmp As Long
    
    tmpWidth = PylonDeviceGetIntegerFeature(m_hDevice, "Width")
    tmpHeight = PylonDeviceGetIntegerFeature(m_hDevice, "Height")
    DoEvents
    If PylonDeviceGrabSingleFrame(m_hDevice, 0, OneBuffer, GrabResult, 200) Then
        DoEvents
        If GrabResult.status = EPylonGrabStatus.Grabbed Then
            'MsgBox "±×·¦¼º°ø", vbInformation
            'CopyMemory m_PixelMemory(m_nID), GetAddrOf(OneBuffer(0)), tmpWidth * tmpHeight
            Sleep 1
            '//// cognex
            Dim Image As New CogImage8Grey

           ' fvImageBuf(m_nID) = GetAddrOf(OneBuffer(0))
            m_favFlip.ClearOperation
            m_favFlip.InputImage = GetAddrOf(OneBuffer(0))
            m_favFlip.ImageWidth = XRES
            m_favFlip.ImageHeight = YRES
            m_favFlip.Run
            m_PixelMemory(m_nID) = m_favFlip.OutputImage
            'm_PixelMemory(m_nID) = GetAddrOf(OneBuffer(0))
            fvImageBuf(m_nID) = m_PixelMemory(m_nID)
            
            frmSetting.FavisImageBoxSetting.DeleteAll
            frmSetting.FavisImageBoxSetting.DrawImage fvImageBuf(m_nID)
            frmSetting.FavisImageBoxSetting.UpdateDraw
'            Call C8Grey_Acquire(m_nID, m_PixelMemory(m_nID))
'            Set m_Image(m_nID) = m_Image(m_nID).Copy(cogImageCopyModeCopyPixels)
'            Set frmSetting.cogSettingDisplay(0).Image = m_Image(m_nID)
    
            'Set g_CogImageBuffer(m_nID) = Nothing
            
            ImageGrabSetting = True
        ElseIf (GrabResult.status = EPylonGrabStatus.Failed) Then
        
            ' Grab error. Display the error code from the GrabResult.
            'MsgBox "Failed to grab image. Errorcode = " & Hex(GrabResult.ErrorCode), vbExclamation
            ImageGrabSetting = False
        End If
        
    Else
        ImageGrabSetting = False
        m_PixelMemory(m_nID) = 0
        'MsgBox "Timeout while acquiring the image.", vbExclamation
    End If

    
Exit Function
err:
    ImageGrabSetting = False
    m_PixelMemory(m_nID) = 0
End Function
Public Sub SetExposure(Value As Long)
On Error GoTo err
    Dim val As Long
    Dim min As Long
    Dim max As Long
    Dim inc As Long
    Dim newVal As Long

    val = PylonDeviceGetIntegerFeature(m_hDevice, "ExposureTimeRaw")
    min = PylonDeviceGetIntegerFeatureMin(m_hDevice, "ExposureTimeRaw")
    max = PylonDeviceGetIntegerFeatureMax(m_hDevice, "ExposureTimeRaw")
    inc = PylonDeviceGetIntegerFeatureInc(m_hDevice, "ExposureTimeRaw")
    
    newVal = min + CLng(Value) * inc
     
    If newVal >= min And newVal < max Then
        PylonDeviceSetIntegerFeature m_hDevice, "ExposureTimeRaw", newVal
    End If
     
     
Exit Sub
err:

End Sub
Public Sub SetGain(Value As Long)
On Error GoTo err
    Dim val As Long
    Dim min As Long
    Dim max As Long
    Dim inc As Long
    Dim newVal As Long

    val = PylonDeviceGetIntegerFeature(m_hDevice, "GainRaw")
    min = PylonDeviceGetIntegerFeatureMin(m_hDevice, "GainRaw")
    max = PylonDeviceGetIntegerFeatureMax(m_hDevice, "GainRaw")
    inc = PylonDeviceGetIntegerFeatureInc(m_hDevice, "GainRaw")
    
    newVal = min + CLng(Value) * inc
     
    If newVal >= min And newVal < max Then
        PylonDeviceSetIntegerFeature m_hDevice, "GainRaw", newVal
    End If
    
    Exit Sub
    
Exit Sub
err:

End Sub
Private Sub C8Grey_Acquire(ByVal intcam As Integer, ByVal longbuffer As Long)
    
    On Error GoTo dra
    
    Dim cog_Iamge8G(7) As New CogImage8Grey
    Dim tmpWidth As Long
    Dim tmpHeight As Long
    Dim tme As IDisposable
    Dim buffer As ICogImage8RootBuffer   '¿øº»:ICogImage8RootBuffer * <============================
    Set buffer = New CogImage8Root
    'Set g_CogImageBuffer(intcam) = Nothing
    ' Intialize the buffer, giving it the image dimensions and a reference
    ' back to this object so it can call Dispose when it's done with
    ' the pixel memory.
    tmpWidth = PylonDeviceGetIntegerFeature(m_hDevice, "Width")
    tmpHeight = PylonDeviceGetIntegerFeature(m_hDevice, "Height")
    
    buffer.Initialize tmpWidth, tmpHeight, longbuffer, tmpWidth, tme
    
    ' Return the buffer to the caller. Note that it's important NOT to store
    ' a reference to the Buffer in this class, because the Buffer already
    ' has a reference to this class's IDisposable interface. If a reference
    ' to the Buffer were stored in this class it would create a circular
    ' reference and the two objects would never get freed up.
    '//// cognex
    m_Image(intcam).SetRoot buffer
 
    Set buffer = Nothing

     
    
    Exit Sub
dra: Debug.Print err.Description
    
End Sub


'====================================================================
' Initialize a pylon device.
' Parameter:
'   hDevice    handle to the device
' Note:
'   Requires the "ExposureTimeRaw" and "GainRaw" features and support
'   for the "Mono8" pixel format.
'   Retrieves the size of the current AOI (Width,Height) and the payload size.
'====================================================================
Private Sub ConfigureDevice(hDevice As PYLON_DEVICE_HANDLE)
    
    Debug.Assert PylonDeviceIsOpen(hDevice)
    
    Dim bShutter As Boolean
    
    ' Disable acquisition start trigger if available.
    If PylonDeviceFeatureIsAvailable(hDevice, "EnumEntry_TriggerSelector_AcquisitionStart") Then
        PylonDeviceFeatureFromString hDevice, "TriggerSelector", "AcquisitionStart"
        PylonDeviceFeatureFromString hDevice, "TriggerMode", "Off"
    End If

    ' Disable frame start trigger if available.
    If PylonDeviceFeatureIsAvailable(hDevice, "EnumEntry_TriggerSelector_FrameStart") Then
        PylonDeviceFeatureFromString hDevice, "TriggerSelector", "FrameStart"
        PylonDeviceFeatureFromString hDevice, "TriggerMode", "Off"
    End If
    
    ' For GigE cameras, we recommend increasing the packet size for better
    ' performance. If the network adapter supports jumbo frames,
    ' set the packet size to a value > 1500, e.g., to 8192.
    ' In this sample, we only set the packet size to 1500 for highest compatibility.
    
    ' Check first to see if the GigE camera packet size parameter
    ' is supported and if it is writable.
    ' On non GigE devices this function will return false.
    If (PylonDeviceFeatureIsWritable(hDevice, "GevSCPSPacketSize") <> 0) Then
        ' The device supports the packet size feature. Set a value.
        PylonDeviceSetIntegerFeature hDevice, "GevSCPSPacketSize", 1500
    End If
    
    ' Make sure chunk mode is off.
    If PylonDeviceFeatureIsWritable(hDevice, "ChunkModeActive") Then
        PylonDeviceSetBooleanFeature hDevice, "ChunkModeActive", 0
    End If
    
    
    
    ' Set the Height and Width parameters.
    Dim Width As Long
    Dim Height As Long
    
    Width = PylonDeviceGetIntegerFeature(hDevice, "Width")
    Height = PylonDeviceGetIntegerFeature(hDevice, "Height")
    
    Dim w As Long
    w = Width Mod 4
    If (w > 0) Then
        ' Adjust image width to be divisible by 4.
        PylonDeviceSetIntegerFeature hDevice, "Width", Width - w
        Width = PylonDeviceGetIntegerFeature(hDevice, "Width")
    End If
    
End Sub


