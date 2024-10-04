Attribute VB_Name = "ModIntelJPEGLibrary"
Option Explicit

' ==================================================================================
' Filename:    mIntelJPEGLibrary.bas
' Author:      Steve McMahon
' Date:        15 March 1999
' Requires:    cDIBSection.cls (vbAccelerator)
'              IJL.DLL (Intel)
'
' An interface to Intel's IJL (Intel JPG Library) for use in VB.
'
' Copyright © 1999 Steve McMahon for vbAccelerator
' http://vbaccelerator.com/
'
' IJL.DLL is a copyright © Intel, which is a registered trade mark of the Intel
' Corporation.
'
' Note.
' Intel are not responsible for any errors in this code and should not be
' mentioned in any Help, About or support in any product using the Intel library.
'
' ==================================================================================

Private Enum IJLERR
  '// The following "error" values indicate an "OK" condition.
  IJL_OK = 0
  IJL_INTERRUPT_OK = 1
  IJL_ROI_OK = 2

  '// The following "error" values indicate an error has occurred.
  IJL_EXCEPTION_DETECTED = -1
  IJL_INVALID_ENCODER = -2
  IJL_UNSUPPORTED_SUBSAMPLING = -3
  IJL_UNSUPPORTED_BYTES_PER_PIXEL = -4
  IJL_MEMORY_ERROR = -5
  IJL_BAD_HUFFMAN_TABLE = -6
  IJL_BAD_QUANT_TABLE = -7
  IJL_INVALID_JPEG_PROPERTIES = -8
  IJL_ERR_FILECLOSE = -9
  IJL_INVALID_FILENAME = -10
  IJL_ERROR_EOF = -11
  IJL_PROG_NOT_SUPPORTED = -12
  IJL_ERR_NOT_JPEG = -13
  IJL_ERR_COMP = -14
  IJL_ERR_SOF = -15
  IJL_ERR_DNL = -16
  IJL_ERR_NO_HUF = -17
  IJL_ERR_NO_QUAN = -18
  IJL_ERR_NO_FRAME = -19
  IJL_ERR_MULT_FRAME = -20
  IJL_ERR_DATA = -21
  IJL_ERR_NO_IMAGE = -22
  IJL_FILE_ERROR = -23
  IJL_INTERNAL_ERROR = -24
  IJL_BAD_RST_MARKER = -25
  IJL_THUMBNAIL_DIB_TOO_SMALL = -26
  IJL_THUMBNAIL_DIB_WRONG_COLOR = -27
  IJL_RESERVED = -99

End Enum

Private Enum IJLIOTYPE
  IJL_SETUP = -1&
  ''// Read JPEG parameters (i.e., height, width, channels,
  ''// sampling, etc.) from a JPEG bit stream.
  IJL_JFILE_READPARAMS = 0&
  IJL_JBUFF_READPARAMS = 1&
  ''// Read a JPEG Interchange Format image.
  IJL_JFILE_READWHOLEIMAGE = 2&
  IJL_JBUFF_READWHOLEIMAGE = 3&
  ''// Read JPEG tables from a JPEG Abbreviated Format bit stream.
  IJL_JFILE_READHEADER = 4&
  IJL_JBUFF_READHEADER = 5&
  ''// Read image info from a JPEG Abbreviated Format bit stream.
  IJL_JFILE_READENTROPY = 6&
  IJL_JBUFF_READENTROPY = 7&
  ''// Write an entire JFIF bit stream.
  IJL_JFILE_WRITEWHOLEIMAGE = 8&
  IJL_JBUFF_WRITEWHOLEIMAGE = 9&
  ''// Write a JPEG Abbreviated Format bit stream.
  IJL_JFILE_WRITEHEADER = 10&
  IJL_JBUFF_WRITEHEADER = 11&
  ''// Write image info to a JPEG Abbreviated Format bit stream.
  IJL_JFILE_WRITEENTROPY = 12&
  IJL_JBUFF_WRITEENTROPY = 13&
  ''// Scaled Decoding Options:
  ''// Reads a JPEG image scaled to 1/2 size.
  IJL_JFILE_READONEHALF = 14&
  IJL_JBUFF_READONEHALF = 15&
  ''// Reads a JPEG image scaled to 1/4 size.
  IJL_JFILE_READONEQUARTER = 16&
  IJL_JBUFF_READONEQUARTER = 17&
  ''// Reads a JPEG image scaled to 1/8 size.
  IJL_JFILE_READONEEIGHTH = 18&
  IJL_JBUFF_READONEEIGHTH = 19&
  ''// Reads an embedded thumbnail from a JFIF bit stream.
  IJL_JFILE_READTHUMBNAIL = 20&
  IJL_JBUFF_READTHUMBNAIL = 21&

End Enum

Private Type JPEG_CORE_PROPERTIES_VB
  UseJPEGPROPERTIES As Long                      '// default = 0

  '// DIB specific I/O data specifiers.
  DIBBytes As Long ';                  '// default = NULL 4
  DIBWidth As Long ';                  '// default = 0 8
  DIBHeight As Long ';                 '// default = 0 12
  DIBPadBytes As Long ';               '// default = 0 16
  DIBChannels As Long ';               '// default = 3 20
  DIBColor As Long ';                  '// default = IJL_BGR 24
  DIBSubsampling As Long  ';            '// default = IJL_NONE 28

  '// JPEG specific I/O data specifiers.
  JPGFile As Long 'LPTSTR              JPGFile;                32   '// default = NULL
  JPGBytes As Long ';                  '// default = NULL 36
  JPGSizeBytes As Long ';              '// default = 0 40
  JPGWidth As Long ';                  '// default = 0 44
  JPGHeight As Long ';                 '// default = 0 48
  JPGChannels As Long ';               '// default = 3
  JPGColor As Long           ';                  '// default = IJL_YCBCR
  JPGSubsampling As Long  ';            '// default = IJL_411
  JPGThumbWidth As Long ' ;             '// default = 0
  JPGThumbHeight As Long ';            '// default = 0

  '// JPEG conversion properties.
  cconversion_reqd As Long ';          '// default = TRUE
  upsampling_reqd As Long ';           '// default = TRUE
  jquality As Long ';                  '// default = 75.  100 is my preferred quality setting.

  '// Low-level properties - 20,000 bytes.  If the whole structure
  ' is written out then VB fails with an obscure error message
  ' "Too Many Local Variables" !
  ' These all default if they are not otherwise specified so there
  ' is no trouble.
  jprops(0 To 19999) As Byte

End Type


Private Declare Function ijlInit Lib "ijl15.dll" (jcprops As Any) As Long
Private Declare Function ijlFree Lib "ijl15.dll" (jcprops As Any) As Long
Private Declare Function ijlRead Lib "ijl15.dll" (jcprops As Any, ByVal ioType As Long) As Long
Private Declare Function ijlWrite Lib "ijl15.dll" (jcprops As Any, ByVal ioType As Long) As Long
Private Declare Function ijlGetLibVersion Lib "ijl15.dll" () As Long
Private Declare Function ijlGetErrorString Lib "ijl15.dll" (ByVal code As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Const GMEM_DDESHARE = &H2000
Public Const GMEM_DISCARDABLE = &H100
Public Const GMEM_DISCARDED = &H4000
Public Const GMEM_FIXED = &H0
Public Const GMEM_INVALID_HANDLE = &H8000
Public Const GMEM_LOCKCOUNT = &HFF
Public Const GMEM_MODIFY = &H80
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_NOCOMPACT = &H10
Public Const GMEM_NODISCARD = &H20
Public Const GMEM_NOT_BANKED = &H1000
Public Const GMEM_NOTIFY = &H4000
Public Const GMEM_SHARE = &H2000
Public Const GMEM_VALID_FLAGS = &H7F72
Public Const GMEM_ZEROINIT = &H40
Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Public Function LoadJPG( _
      ByRef cDib As cDIBSection, _
      ByVal sFile As String _
   ) As Boolean
Dim tJ As JPEG_CORE_PROPERTIES_VB
Dim bFile() As Byte
Dim lR As Long
Dim lPtr As Long
Dim lJPGWidth As Long, lJPGHeight As Long

   lR = ijlInit(tJ)
   If lR = IJL_OK Then
      
      ' Write the filename to the jcprops.JPGFile member:
      bFile = StrConv(sFile, vbFromUnicode)
      ReDim Preserve bFile(0 To UBound(bFile) + 1) As Byte
      bFile(UBound(bFile)) = 0
      lPtr = VarPtr(bFile(0))
      CopyMemory tJ.JPGFile, lPtr, 4
      
      ' Read the JPEG file parameters:
      lR = ijlRead(tJ, IJL_JFILE_READPARAMS)
      If lR <> IJL_OK Then
         ' Throw error
         MsgBox "Failed to read JPG", vbExclamation
      Else
      
         ' Get the JPGWidth ...
         lJPGWidth = tJ.JPGWidth
         ' .. & JPGHeight member values:
         lJPGHeight = tJ.JPGHeight
      
         ' Create a buffer of sufficient size to hold the image:
         If cDib.Create(lJPGWidth, lJPGHeight) Then
            ' Store DIBWidth:
            tJ.DIBWidth = lJPGWidth
            ' Very important: tell IJL how many bytes extra there
            ' are on each DIB scan line to pad to 32 bit boundaries:
            tJ.DIBPadBytes = cDib.BytesPerScanLine - lJPGWidth * 3
            ' Store DIBHeight:
            tJ.DIBHeight = -lJPGHeight
            ' Store Channels:
            tJ.DIBChannels = 3&
            ' Store DIBBytes (pointer to uncompressed JPG data):
            tJ.DIBBytes = cDib.DIBSectionBitsPtr
            
            ' Now decompress the JPG into the DIBSection:
            lR = ijlRead(tJ, IJL_JFILE_READWHOLEIMAGE)
            If lR = IJL_OK Then
               ' That's it!  cDib now contains the uncompressed JPG.
               LoadJPG = True
            Else
               ' Throw error:
               MsgBox "Cannot read Image Data from file.", vbExclamation
            End If
         Else
            ' failed to create the DIB...
         End If
      End If
                        
      ' Ensure we have freed memory:
      ijlFree tJ
   Else
      ' Throw error:
      MsgBox "Failed to initialise the IJL library: " & lR, vbExclamation
   End If
   
   
End Function
Public Function SaveJPG( _
      ByRef cDib As cDIBSection, _
      ByVal sFile As String, _
      Optional ByVal lQuality As Long = 90 _
   ) As Boolean
Dim tJ As JPEG_CORE_PROPERTIES_VB
Dim bFile() As Byte
Dim lPtr As Long
Dim lR As Long
   
   lR = ijlInit(tJ)
   If lR = IJL_OK Then
      ' Set up the DIB information:
      ' Store DIBWidth:
      tJ.DIBWidth = cDib.Width
      ' Store DIBHeight:
      tJ.DIBHeight = -cDib.Height
      ' Store DIBBytes (pointer to uncompressed JPG data):
      tJ.DIBBytes = cDib.DIBSectionBitsPtr
      ' Very important: tell IJL how many bytes extra there
      ' are on each DIB scan line to pad to 32 bit boundaries:
      tJ.DIBPadBytes = cDib.BytesPerScanLine - cDib.Width * 3
      
      ' Set up the JPEG information:
      
      ' Store JPGFile:
      bFile = StrConv(sFile, vbFromUnicode)
      ReDim Preserve bFile(0 To UBound(bFile) + 1) As Byte
      bFile(UBound(bFile)) = 0
      lPtr = VarPtr(bFile(0))
      CopyMemory tJ.JPGFile, lPtr, 4
      ' Store JPGWidth:
      tJ.JPGWidth = cDib.Width
      ' .. & JPGHeight member values:
      tJ.JPGHeight = cDib.Height
      ' Set the quality/compression to save:
      tJ.jquality = lQuality
            
      ' Write the image:
      lR = ijlWrite(tJ, IJL_JFILE_WRITEWHOLEIMAGE)
      If lR = IJL_OK Then
         SaveJPG = True
      Else
         ' Throw error
         'MsgBox "Failed to save to JPG", vbExclamation
      End If
      
      ' Ensure we have freed memory:
      ijlFree tJ
   Else
      ' Throw error:
      MsgBox "Failed to initialise the IJL library: " & lR, vbExclamation
   End If
   

End Function
