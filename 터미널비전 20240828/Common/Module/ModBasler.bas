Attribute VB_Name = "ModBasler"
Option Explicit


'Public Declare Function GetAddrOf Lib "kernel32" Alias "MulDiv" (nNumber As Any, Optional ByVal nNumerator As Long = 1, Optional ByVal nDenominator As Long = 1) As Long '
'Public m_Camera(3) As BaslerCamera
Public g_CogImageBuffer(7) As CogImage8Grey
Public m_LiveCheck As Boolean
Public m_icamIndex As Integer

Public Function Init_camera() As Boolean
On Error GoTo err

Dim hDeviceInfo As PYLON_DEVICE_INFO_HANDLE
Dim numDevices As Long
Dim i As Long
Dim hDevice As PYLON_DEVICE_HANDLE


    'm_hDevice = cPylonInvalidHandle
    'm_hCallback = cPylonInvalidHandle
    
    
    
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
        name = Left(name, 1)
 '       name = CStr(CLng(name) + 1)
'        If CLng(name - 2) = m_nID Then
'
'            DeviceIndex = m_nID
'            Exit For
'
'        End If

        

        Select Case CLng(name)
        
            Case 1
                m_Camera(0).m_hDevice = PylonCreateDeviceByIndex(i)
                PylonDeviceOpen m_Camera(0).m_hDevice, (cPylonAccessModeExclusive Or cPylonAccessModeStream Or cPylonAccessModeControl)
                ConfigureDevice m_Camera(0).m_hDevice
            Case 2
                m_Camera(1).m_hDevice = PylonCreateDeviceByIndex(i)
                PylonDeviceOpen m_Camera(1).m_hDevice, (cPylonAccessModeExclusive Or cPylonAccessModeStream Or cPylonAccessModeControl)
                ConfigureDevice m_Camera(1).m_hDevice
            Case 3
                m_Camera(2).m_hDevice = PylonCreateDeviceByIndex(i)
                PylonDeviceOpen m_Camera(2).m_hDevice, (cPylonAccessModeExclusive Or cPylonAccessModeStream Or cPylonAccessModeControl)
                ConfigureDevice m_Camera(2).m_hDevice
            Case 4
                m_Camera(3).m_hDevice = PylonCreateDeviceByIndex(i)
                PylonDeviceOpen m_Camera(3).m_hDevice, (cPylonAccessModeExclusive Or cPylonAccessModeStream Or cPylonAccessModeControl)
                ConfigureDevice m_Camera(3).m_hDevice
        End Select
    
'    If m_hDevice = 0 Then
'        MsgBox "Cam Init failed!!!"
'        Exit Function
'    End If
    
    ' Open the device.
'    PylonDeviceOpen hDevice, (cPylonAccessModeExclusive Or cPylonAccessModeStream Or cPylonAccessModeControl)
    ' Extend the heartbeat timeout to make debuggig easier.
    'SetHeartbeatTimeoutFromCommandLine m_hDevice
    ' Set up the camera (will raise an error if the camera isn't suitable).
'    ConfigureDevice hDevice
    
     Next i

    Init_camera = True
Exit Function
err:
    Init_camera = False
    
End Function

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
