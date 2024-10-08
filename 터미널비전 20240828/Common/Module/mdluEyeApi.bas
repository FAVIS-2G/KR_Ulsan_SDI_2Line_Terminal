Attribute VB_Name = "mdluEyeApi"

'/===========================================================================//
'/                                                                           //
'/  Copyright (C) 2004 - 2007                                                //
'/  IDS Imaging GmbH                                                         //
'/  Dimbacherstr. 6                                                          //
'/  D-74182 Obersulm-Willsbach                                               //
'/                                                                           //
'/  The information in this document is subject to change without            //
'/  notice and should not be construed as a commitment by IDS Imaging GmbH.  //
'/  IDS Imaging GmbH does not assume any responsibility for any errors       //
'/  that may appear in this document.                                        //
'/                                                                           //
'/  This document, or source code, is provided solely as an example          //
'/  of how to utilize IDS software libraries in a sample application.        //
'/  IDS Imaging GmbH does not assume any responsibility for the use or       //
'/  reliability of any portion of this document or the described software.   //
'/                                                                           //
'/  General permission to copy or modify, but not for profit, is hereby      //
'/  granted,  provided that the above copyright notice is included and       //
'/  reference made to the fact that reproduction privileges were granted     //
'/  by IDS Imaging GmbH.                                                     //
'/                                                                           //
'/  IDS cannot assume any responsibility for the use or misuse of any        //
'/  portion of this software for other than its intended diagnostic purpose  //
'/  in calibrating and testing IDS manufactured cameras and software.        //
'/                                                                           //
'/===========================================================================//

Option Explicit


' ***********************************************
' uEye dll name
' ***********************************************
Public Const DRIVER_DLL_NAME = "ueye_api.dll"


' ***********************************************
' basic color modes
' ***********************************************
Public Const IS_COLORMODE_INVALID = 0
Public Const IS_COLORMODE_MONOCHROME = 1
Public Const IS_COLORMODE_BAYER = 2


' ***********************************************
' sensor types
' ***********************************************
Public Const IS_SENSOR_INVALID = &H0&

' CMOS Sensors
Public Const IS_SENSOR_UI141X_M = &H1&          ' VGA rolling shutter - VGA monochrome
Public Const IS_SENSOR_UI141X_C = &H2&          ' VGA rolling shutter - VGA color
Public Const IS_SENSOR_UI144X_M = &H3&          ' SXGA rolling shutter - SXGA monochrome
Public Const IS_SENSOR_UI144X_C = &H4&          ' SXGA rolling shutter - SXGA color

Public Const IS_SENSOR_UI145X_C = &H8&          ' UXGA rolling shutter - UXGA color
Public Const IS_SENSOR_UI146X_C = &HA&          ' QXGA rolling shutter - QXGA color
Public Const IS_SENSOR_UI148X_C = &HC&          '

Public Const IS_SENSOR_UI121X_M = &H10&         ' VGA global shutter - VGA monochrome
Public Const IS_SENSOR_UI121X_C = &H11&         ' VGA global shutter - VGA color
Public Const IS_SENSOR_UI122X_M = &H12&         ' VGA global shutter - VGA monochrome
Public Const IS_SENSOR_UI122X_C = &H13&         ' VGA global shutter - VGA color

Public Const IS_SENSOR_UI164X_C = &H15&         ' SXGA rolling shutter, color
Public Const IS_SENSOR_UI155X_C = &H17&         ' UXGA rolling shutter, color

Public Const IS_SENSOR_UI1225_M = &H22&         ' WVGA global shutter, monochrome, LE model
Public Const IS_SENSOR_UI1225_C = &H23&         ' WVGA global shutter, color, LE model

Public Const IS_SENSOR_UI1645_C = &H25&         ' SXGA rolling shutter, color, LE model
Public Const IS_SENSOR_UI1555_C = &H27&         ' UXGA rolling shutter, color, LE model


Public Const IS_SENSOR_UI1545_M = &H28&         ' SXGA rolling shutter, monochrome, LE model
Public Const IS_SENSOR_UI1545_C = &H29&         ' SXGA rolling shutter, color, LE model
Public Const IS_SENSOR_UI1455_C = &H2B&         ' UXGA rolling shutter, color, LE model
Public Const IS_SENSOR_UI1465_C = &H2D&         ' QXGA rolling shutter, color, LE model
Public Const IS_SENSOR_UI1485_C = &H2F&         ' 5MP rolling shutter, color, LE model

Public Const IS_SENSOR_UI154X_M = &H30&         ' QXGA rolling shutter - QXGA monochrome
Public Const IS_SENSOR_UI154X_C = &H31&         ' QXGA rolling shutter - QXGA color
Public Const IS_SENSOR_UI1543_M = &H32&         ' SXGA rolling shutter - SXGA monochrome
Public Const IS_SENSOR_UI1543_C = &H33&         ' SXGA rolling shutter - SXGA color

Public Const IS_SENSOR_UI1453_C = &H35&         ' UXGA rolling shutter - UXGA color
Public Const IS_SENSOR_UI1463_C = &H37&         ' QXGA rolling shutter - QXGA color
Public Const IS_SENSOR_UI1483_C = &H39&         ' 5MP rolling shutter, color, single board
Public Const IS_SENSOR_UI1544_C = &H3B&         ' SXGA rolling shutter, color, single board

' CCD Sensors
Public Const IS_SENSOR_UI223X_M = &H80&         ' Sony CCD sensor - XGA monochrome
Public Const IS_SENSOR_UI223X_C = &H81&         ' Sony CCD sensor - XGA color

Public Const IS_SENSOR_UI241X_M = &H82&         ' Sony CCD sensor - VGA monochrome
Public Const IS_SENSOR_UI241X_C = &H83&         ' Sony CCD sensor - VGA color

Public Const IS_SENSOR_UI234X_M = &H84&         '
Public Const IS_SENSOR_UI234X_C = &H85&         '

Public Const IS_SENSOR_UI233X_M = &H86&         ' Sony CCD sensor - XGA / SXGA monochrome
Public Const IS_SENSOR_UI233X_C = &H87&         ' Sony CCD sensor - XGA / SXGA color

Public Const IS_SENSOR_UI221X_M = &H88&         ' Sony CCD sensor - VGA monochrome
Public Const IS_SENSOR_UI221X_C = &H89&         ' Sony CCD sensor - VGA color

Public Const IS_SENSOR_UI231X_M = &H90&         ' Sony CCD sensor - VGA monochrome
Public Const IS_SENSOR_UI231X_C = &H91&         ' Sony CCD sensor - VGA color

Public Const IS_SENSOR_UI222X_M = &H92&         ' Sony CCD sensor - CCIR / PAL monochrome
Public Const IS_SENSOR_UI222X_C = &H93&         ' Sony CCD sensor - CCIR / PAL color

Public Const IS_SENSOR_UI224X_M = &H96&         ' Sony CCD sensor - SXGA monochrome
Public Const IS_SENSOR_UI224X_C = &H97&         ' Sony CCD sensor - SXGA color

Public Const IS_SENSOR_UI225X_M = &H98&         ' Sony CCD sensor - UXGA monochrome
Public Const IS_SENSOR_UI225X_C = &H99&         ' Sony CCD sensor - UXGA color


' ***********************************************
' return values/error codes
' ***********************************************
Public Const IS_SUCCESS = 0
Public Const IS_NO_SUCCESS = -1
Public Const IS_INVALID_CAMERA_HANDLE = 1
Public Const IS_INVALID_HANDLE = 1

Public Const IS_IO_REQUEST_FAILED = 2
Public Const IS_CANT_OPEN_DEVICE = 3
Public Const IS_CANT_CLOSE_DEVICE = 4
Public Const IS_CANT_SETUP_MEMORY = 5
Public Const IS_NO_HWND_FOR_ERROR_REPORT = 6
Public Const IS_ERROR_MESSAGE_NOT_CREATED = 7
Public Const IS_ERROR_STRING_NOT_FOUND = 8
Public Const IS_HOOK_NOT_CREATED = 9
Public Const IS_TIMER_NOT_CREATED = 10
Public Const IS_CANT_OPEN_REGISTRY = 11
Public Const IS_CANT_READ_REGISTRY = 12
Public Const IS_CANT_VALIDATE_BOARD = 13
Public Const IS_CANT_GIVE_BOARD_ACESS = 14
Public Const IS_NO_IMAGE_MEM_ALLOCATED = 15
Public Const IS_CANT_CLEANUP_MEMORY = 16
Public Const IS_CANT_COMMUNICATE_WITH_DRIVER = 17
Public Const IS_FUNCTION_NOT_SUPPORTED_YET = 18
Public Const IS_OPERATING_SYSTEM_NOT_SUPPORTED = 19

Public Const IS_INVALID_VIDEO_IN = 20
Public Const IS_INVALID_IMG_SIZE = 21
Public Const IS_INVALID_ADDRESS = 22
Public Const IS_INVALID_VIDEO_MODE = 23
Public Const IS_INVALID_AGC_MODE = 24
Public Const IS_INVALID_GAMMA_MODE = 25
Public Const IS_INVALID_SYNC_LEVEL = 26
Public Const IS_INVALID_CBARS_MODE = 27
Public Const IS_INVALID_COLOR_MODE = 28
Public Const IS_INVALID_SCALE_FACTOR = 29
Public Const IS_INVALID_IMAGE_SIZE = 30
Public Const IS_INVALID_IMAGE_POS = 31
Public Const IS_INVALID_CAPTURE_MODE = 32
Public Const IS_INVALID_RISC_PROGRAM = 33
Public Const IS_INVALID_BRIGHTNESS = 34
Public Const IS_INVALID_CONTRAST = 35
Public Const IS_INVALID_SATURATION_U = 36
Public Const IS_INVALID_SATURATION_V = 37
Public Const IS_INVALID_HUE = 38
Public Const IS_INVALID_HOR_FILTER_STEP = 39
Public Const IS_INVALID_VERT_FILTER_STEP = 40
Public Const IS_INVALID_EEPROM_READ_ADDRESS = 41
Public Const IS_INVALID_EEPROM_WRITE_ADDRESS = 42
Public Const IS_INVALID_EEPROM_READ_LENGTH = 43
Public Const IS_INVALID_EEPROM_WRITE_LENGTH = 44
Public Const IS_INVALID_BOARD_INFO_POINTER = 45
Public Const IS_INVALID_DISPLAY_MODE = 46
Public Const IS_INVALID_ERR_REP_MODE = 47
Public Const IS_INVALID_BITS_PIXEL = 48
Public Const IS_INVALID_MEMORY_POINTER = 49

Public Const IS_FILE_WRITE_OPEN_ERROR = 50
Public Const IS_FILE_READ_OPEN_ERROR = 51
Public Const IS_FILE_READ_INVALID_BMP_ID = 52
Public Const IS_FILE_READ_INVALID_BMP_SIZE = 53
Public Const IS_FILE_READ_INVALID_BIT_COUNT = 54
Public Const IS_WRONG_KERNEL_VERSION = 55

Public Const IS_RISC_INVALID_XLENGTH = 60
Public Const IS_RISC_INVALID_YLENGTH = 61
Public Const IS_RISC_EXCEED_IMG_SIZE = 62

Public Const IS_DD_MAIN_FAILED = 70
Public Const IS_DD_PRIMSURFACE_FAILED = 71
Public Const IS_DD_SCRN_SIZE_NOT_SUPPORTED = 72
Public Const IS_DD_CLIPPER_FAILED = 73
Public Const IS_DD_CLIPPER_HWND_FAILED = 74
Public Const IS_DD_CLIPPER_CONNECT_FAILED = 75
Public Const IS_DD_BACKSURFACE_FAILED = 76
Public Const IS_DD_BACKSURFACE_IN_SYSMEM = 77
Public Const IS_DD_MDL_MALLOC_ERR = 78
Public Const IS_DD_MDL_SIZE_ERR = 79
Public Const IS_DD_CLIP_NO_CHANGE = 80
Public Const IS_DD_PRIMMEM_NULL = 81
Public Const IS_DD_BACKMEM_NULL = 82
Public Const IS_DD_BACKOVLMEM_NULL = 83
Public Const IS_DD_OVERLAYSURFACE_FAILED = 84
Public Const IS_DD_OVERLAYSURFACE_IN_SYSMEM = 85
Public Const IS_DD_OVERLAY_NOT_ALLOWED = 86
Public Const IS_DD_OVERLAY_COLKEY_ERR = 87
Public Const IS_DD_OVERLAY_NOT_ENABLED = 88
Public Const IS_DD_GET_DC_ERROR = 89
Public Const IS_DD_DDRAW_DLL_NOT_LOADED = 90
Public Const IS_DD_THREAD_NOT_CREATED = 91
Public Const IS_DD_CANT_GET_CAPS = 92
Public Const IS_DD_NO_OVERLAYSURFACE = 93
Public Const IS_DD_NO_OVERLAYSTRETCH = 94
Public Const IS_DD_CANT_CREATE_OVERLAYSURFACE = 95
Public Const IS_DD_CANT_UPDATE_OVERLAYSURFACE = 96
Public Const IS_DD_INVALID_STRETCH = 97

Public Const IS_EV_INVALID_EVENT_NUMBER = 100
Public Const IS_INVALID_MODE = 101
Public Const IS_CANT_FIND_FALCHOOK = 102
Public Const IS_CANT_FIND_HOOK = 102
Public Const IS_CANT_GET_HOOK_PROC_ADDR = 103
Public Const IS_CANT_CHAIN_HOOK_PROC = 104
Public Const IS_CANT_SETUP_WND_PROC = 105
Public Const IS_HWND_NULL = 106
Public Const IS_INVALID_UPDATE_MODE = 107
Public Const IS_NO_ACTIVE_IMG_MEM = 108
Public Const IS_CANT_INIT_EVENT = 109
Public Const IS_FUNC_NOT_AVAIL_IN_OS = 110
Public Const IS_CAMERA_NOT_CONNECTED = 111
Public Const IS_SEQUENCE_LIST_EMPTY = 112
Public Const IS_CANT_ADD_TO_SEQUENCE = 113
Public Const IS_LOW_OF_SEQUENCE_RISC_MEM = 114
Public Const IS_IMGMEM2FREE_USED_IN_SEQ = 115
Public Const IS_IMGMEM_NOT_IN_SEQUENCE_LIST = 116
Public Const IS_SEQUENCE_BUF_ALREADY_LOCKED = 117
Public Const IS_INVALID_DEVICE_ID = 118
Public Const IS_INVALID_BOARD_ID = 119
Public Const IS_ALL_DEVICES_BUSY = 120
Public Const IS_HOOK_BUSY = 121
Public Const IS_TIMED_OUT = 122
Public Const IS_NULL_POINTER = 123
Public Const IS_WRONG_HOOK_VERSION = 124
Public Const IS_INVALID_PARAMETER = 125
Public Const IS_NOT_ALLOWED = 126
Public Const IS_OUT_OF_MEMORY = 127
Public Const IS_INVALID_WHILE_LIVE = 128
Public Const IS_ACCESS_VIOLATION = 129
Public Const IS_UNKNOWN_ROP_EFFECT = 130
Public Const IS_INVALID_RENDER_MODE = 131
Public Const IS_INVALID_THREAD_CONTEXT = 132
Public Const IS_NO_HARDWARE_INSTALLED = 133
Public Const IS_INVALID_WATCHDOG_TIME = 134
Public Const IS_INVALID_WATCHDOG_MODE = 135
Public Const IS_INVALID_PASSTHROUGH_IN = 136
Public Const IS_ERROR_SETTING_PASSTHROUGH_IN = 137
Public Const IS_FAILURE_ON_SETTING_WATCHDOG = 138
Public Const IS_ERROR_SETTING_DIGITAL_OUT = 139
Public Const IS_CAPTURE_RUNNING = 140

Public Const IS_MEMORY_BOARD_ACTIVATED = 141
Public Const IS_MEMORY_BOARD_DEACTIVATED = 142
Public Const IS_NO_MEMORY_BOARD_CONNECTED = 143
Public Const IS_TOO_LESS_MEMORY = 144
Public Const IS_IMAGE_NOT_PRESENT = 145
Public Const IS_MEMORY_MODE_RUNNING = 146
Public Const IS_MEMORYBOARD_DISABLED = 147

Public Const IS_TRIGGER_ACTIVATED = 148
Public Const IS_WRONG_KEY = 150
Public Const IS_CRC_ERROR = 151
Public Const IS_NOT_YET_RELEASED = 152
Public Const IS_NOT_CALIBRATED = 153
Public Const IS_WAITING_FOR_KERNEL = 154
Public Const IS_NOT_SUPPORTED = 155
Public Const IS_TRIGGER_NOT_ACTIVATED = 156
Public Const IS_OPERATION_ABORTED = 157
Public Const IS_BAD_STRUCTURE_SIZE = 158
Public Const IS_INVALID_BUFFER_SIZE = 159
Public Const IS_INVALID_PIXEL_CLOCK = 160
Public Const IS_INVALID_EXPOSURE_TIME = 161
Public Const IS_AUTO_EXPOSURE_RUNNING = 162
Public Const IS_CANNOT_CREATE_BB_SURF = 163         ' error creating backbuffer surface
Public Const IS_CANNOT_CREATE_BB_MIX = 164          ' backbuffer mixer surfaces can not be created
Public Const IS_BB_OVLMEM_NULL = 165                ' backbuffer overlay mem could not be locked
Public Const IS_CANNOT_CREATE_BB_OVL = 166          ' backbuffer overlay mem could not be created
Public Const IS_NOT_SUPP_IN_OVL_SURF_MODE = 167     ' function not supported in overlay surface mode
Public Const IS_INVALID_SURFACE = 168               ' surface invalid
Public Const IS_SURFACE_LOST = 169                  ' surface has been lost
Public Const IS_RELEASE_BB_OVL_DC = 170             ' error releasing backbuffer overlay DC
Public Const IS_BB_TIMER_NOT_CREATED = 171          ' backbuffer timer could not be created
Public Const IS_BB_OVL_NOT_EN = 172                 ' backbuffer overlay has not been enabled
Public Const IS_ONLY_IN_BB_MODE = 173               ' only possible in backbuffer mode
Public Const IS_INVALID_COLOR_FORMAT = 174          ' invalid color format
Public Const IS_INVALID_WB_BINNING_MODE = 175
Public Const IS_INVALID_I2C_DEVICE_ADDRESS = 176
Public Const IS_COULD_NOT_CONVERT = 177             ' current image couldn't be converted
Public Const IS_TRANSFER_ERROR = 178                ' transfer failed


' ***********************************************
' common definitions
' ***********************************************
Public Const IS_OFF = 0
Public Const IS_ON = 1
Public Const IS_IGNORE_PARAMETER = -1


' ***********************************************
' device enumeration
' ***********************************************
Public Const IS_USE_DEVICE_ID = &H8000&


' ***********************************************
' autoExit enable/disable
' ***********************************************
Public Const IS_GET_AUTO_EXIT_ENABLED = &H8000&
Public Const IS_DISABLE_AUTO_EXIT = 0
Public Const IS_ENABLE_AUTO_EXIT = 1


' ***********************************************
' live/freeze parameters
' ***********************************************
Public Const IS_GET_LIVE = &H8000&

Public Const IS_WAIT = 1
Public Const IS_DONT_WAIT = 0
Public Const IS_FORCE_VIDEO_STOP = &H4000&
Public Const IS_FORCE_VIDEO_START = &H4000&


' ***********************************************
' video finish constants
' ***********************************************
Public Const IS_VIDEO_NOT_FINISH = 0
Public Const IS_VIDEO_FINISH = 1


' ***********************************************
' bitmap render modes
' ***********************************************
Public Const IS_GET_RENDER_MODE = &H8000&

Public Const IS_RENDER_DISABLED = 0
Public Const IS_RENDER_NORMAL = 1
Public Const IS_RENDER_FIT_TO_WINDOW = 2
Public Const IS_RENDER_DOWNSCALE_1_2 = 4
Public Const IS_RENDER_MIRROR_UPDOWN = 16
Public Const IS_RENDER_DOUBLE_HEIGHT = 32
Public Const IS_RENDER_HALF_HEIGHT = 64


' ***********************************************
' external trigger mode constants
' ***********************************************
Public Const IS_GET_EXTERNALTRIGGER = &H8000&
Public Const IS_GET_TRIGGER_STATUS = &H8001&
Public Const IS_GET_TRIGGER_MASK = &H8002&
Public Const IS_GET_TRIGGER_INPUTS = &H8003&
Public Const IS_GET_SUPPORTED_TRIGGER_MODE = &H8004&
Public Const IS_GET_TRIGGER_COUNTER = &H8000&

Public Const IS_SET_TRIG_OFF = &H0&
Public Const IS_SET_TRIG_HI_LO = &H1&
Public Const IS_SET_TRIG_LO_HI = &H2&
Public Const IS_SET_TRIG_SOFTWARE = &H8&
Public Const IS_SET_TRIG_MASK = &H100&

Public Const IS_GET_TRIGGER_DELAY = &H8000&
Public Const IS_GET_MIN_TRIGGER_DELAY = &H8001&
Public Const IS_GET_MAX_TRIGGER_DELAY = &H8002&
Public Const IS_GET_TRIGGER_DELAY_GRANULARITY = &H8003&


' ***********************************************
'  timing
' ***********************************************
' pixelclock
Public Const IS_GET_PIXEL_CLOCK = &H8000&
Public Const IS_GET_DEFAULT_PIXEL_CLK = &H8001&
' framerate
Public Const IS_GET_FRAMERATE = &H8000&
Public Const IS_GET_DEFAULT_FRAMERATE = &H8001&
' exposure
Public Const IS_GET_EXPOSURE_TIME = &H8000&
Public Const IS_GET_DEFAULT_EXPOSURE = &H8001&


' ***********************************************
' gain definitions
' ***********************************************
Public Const IS_GET_MASTER_GAIN = &H8000&
Public Const IS_GET_RED_GAIN = &H8001&
Public Const IS_GET_GREEN_GAIN = &H8002&
Public Const IS_GET_BLUE_GAIN = &H8003&
Public Const IS_GET_DEFAULT_MASTER = &H8004&
Public Const IS_GET_DEFAULT_RED = &H8005&
Public Const IS_GET_DEFAULT_GREEN = &H8006&
Public Const IS_GET_DEFAULT_BLUE = &H8007&
Public Const IS_GET_GAINBOOST = &H8008&
Public Const IS_SET_GAINBOOST_ON = &H1&
Public Const IS_SET_GAINBOOST_OFF = &H0&
Public Const IS_GET_SUPPORTED_GAINBOOST = &H2&


' ***********************************************
' gain factor definitions
' ***********************************************
Public Const IS_GET_MASTER_GAIN_FACTOR = &H8000&
Public Const IS_GET_RED_GAIN_FACTOR = &H8001&
Public Const IS_GET_GREEN_GAIN_FACTOR = &H8002&
Public Const IS_GET_BLUE_GAIN_FACTOR = &H8003&
Public Const IS_SET_MASTER_GAIN_FACTOR = &H8004&
Public Const IS_SET_RED_GAIN_FACTOR = &H8005&
Public Const IS_SET_GREEN_GAIN_FACTOR = &H8006&
Public Const IS_SET_BLUE_GAIN_FACTOR = &H8007&
Public Const IS_GET_DEFAULT_MASTER_GAIN_FACTOR = &H8008&
Public Const IS_GET_DEFAULT_RED_GAIN_FACTOR = &H8009&
Public Const IS_GET_DEFAULT_GREEN_GAIN_FACTOR = &H800A&
Public Const IS_GET_DEFAULT_BLUE_GAIN_FACTOR = &H800B&
Public Const IS_INQUIRE_MASTER_GAIN_FACTOR = &H800C&
Public Const IS_INQUIRE_RED_GAIN_FACTOR = &H800D&
Public Const IS_INQUIRE_GREEN_GAIN_FACTOR = &H800E&
Public Const IS_INQUIRE_BLUE_GAIN_FACTOR = &H800F&


' ***********************************************
' blacklevel compensation
' ***********************************************
Public Const IS_GET_BL_COMPENSATION = &H8000&
Public Const IS_GET_BL_OFFSET = &H8001&
Public Const IS_GET_BL_DEFAULT_MODE = &H8002&
Public Const IS_GET_BL_DEFAULT_OFFSET = &H8003&
Public Const IS_GET_BL_SUPPORTED_MODE = &H8004&

Public Const IS_BL_COMPENSATION_DISABLE = 0
Public Const IS_BL_COMPENSATION_ENABLE = 1
Public Const IS_BL_COMPENSATION_OFFSET = 32


' ***********************************************
' hardware gamma definitions
' ***********************************************
Public Const IS_GET_HW_GAMMA = &H8000&
Public Const IS_GET_HW_SUPPORTED_GAMMA = &H8001&
Public Const IS_SET_HW_GAMMA_OFF = &H0&
Public Const IS_SET_HW_GAMMA_ON = &H1&


' ***********************************************
' Image parameters
' ***********************************************
' brightness
Public Const IS_GET_BRIGHTNESS = &H8000&
Public Const IS_MIN_BRIGHTNESS = 0
Public Const IS_MAX_BRIGHTNESS = 255
Public Const IS_DEFAULT_BRIGHTNESS = -1
' contrast
Public Const IS_GET_CONTRAST = &H8000&
Public Const IS_MIN_CONTRAST = 0
Public Const IS_MAX_CONTRAST = 511
Public Const IS_DEFAULT_CONTRAST = -1
' gamma
Public Const IS_GET_GAMMA = &H8000&
Public Const IS_MIN_GAMMA = 1
Public Const IS_MAX_GAMMA = 1000
Public Const IS_DEFAULT_GAMMA = -1
' saturation
Public Const IS_GET_SATURATION_U = &H8000&
Public Const IS_MIN_SATURATION_U = 0
Public Const IS_MAX_SATURATION_U = 511
Public Const IS_DEFAULT_SATURATION_U = 254
Public Const IS_GET_SATURATION_V = &H8001&
Public Const IS_MIN_SATURATION_V = 0
Public Const IS_MAX_SATURATION_V = 511
Public Const IS_DEFAULT_SATURATION_V = 180
' hue
Public Const IS_GET_HUE = &H8000&
Public Const IS_MIN_HUE = 0
Public Const IS_MAX_HUE = 255
Public Const IS_DEFAULT_HUE = 128


' ***********************************************
' image pos + size
' ***********************************************
Public Const IS_GET_IMAGE_SIZE_X = &H8000&
Public Const IS_GET_IMAGE_SIZE_Y = &H8001&
Public Const IS_GET_IMAGE_SIZE_X_INC = &H8002&
Public Const IS_GET_IMAGE_SIZE_Y_INC = &H8003&
Public Const IS_GET_IMAGE_SIZE_X_MIN = &H8004&
Public Const IS_GET_IMAGE_SIZE_Y_MIN = &H8005&
Public Const IS_GET_IMAGE_SIZE_X_MAX = &H8006&
Public Const IS_GET_IMAGE_SIZE_Y_MAX = &H8007&

Public Const IS_GET_IMAGE_POS_X = &H8001&
Public Const IS_GET_IMAGE_POS_Y = &H8002&
Public Const IS_GET_IMAGE_POS_X_ABS = &HC001&
Public Const IS_GET_IMAGE_POS_Y_ABS = &HC002&
Public Const IS_GET_IMAGE_POS_X_INC = &HC003&
Public Const IS_GET_IMAGE_POS_Y_INC = &HC004&
Public Const IS_GET_IMAGE_POS_X_MIN = &HC005&
Public Const IS_GET_IMAGE_POS_Y_MIN = &HC006&
Public Const IS_GET_IMAGE_POS_X_MAX = &HC007&
Public Const IS_GET_IMAGE_POS_Y_MAX = &HC008&

Public Const IS_SET_IMAGE_POS_X_ABS = &H10000
Public Const IS_SET_IMAGE_POS_Y_ABS = &H10000

' Compatibility
Public Const IS_SET_IMAGEPOS_X_ABS = &H8000&
Public Const IS_SET_IMAGEPOS_Y_ABS = &H8000&


' ***********************************************
' rop effect constants
' ***********************************************
Public Const IS_GET_ROP_EFFECT = &H8000&

Public Const IS_SET_ROP_NONE = 0
Public Const IS_SET_ROP_MIRROR_UPDOWN = 8
Public Const IS_SET_ROP_MIRROR_UPDOWN_ODD = 16
Public Const IS_SET_ROP_MIRROR_UPDOWN_EVEN = 32
Public Const IS_SET_ROP_MIRROR_LEFTRIGHT = 64


' ***********************************************
' subsampling
' ***********************************************
Public Const IS_GET_SUBSAMPLING = &H8000&
Public Const IS_GET_SUPPORTED_SUBSAMPLING = &H8001&

Public Const IS_SUBSAMPLING_DISABLE = &H0&

Public Const IS_SUBSAMPLING_2X_VERTICAL = &H1&
Public Const IS_SUBSAMPLING_2X_HORIZONTAL = &H2&
Public Const IS_SUBSAMPLING_4X_VERTICAL = &H4&
Public Const IS_SUBSAMPLING_4X_HORIZONTAL = &H8&

Public Const IS_SUBSAMPLING_MASK_VERTICAL = (IS_SUBSAMPLING_2X_VERTICAL Or IS_SUBSAMPLING_4X_VERTICAL)
Public Const IS_SUBSAMPLING_MASK_HORIZONTAL = (IS_SUBSAMPLING_2X_HORIZONTAL Or IS_SUBSAMPLING_4X_HORIZONTAL)

' Compatibility
Public Const IS_SUBSAMPLING_VERT = IS_SUBSAMPLING_2X_VERTICAL
Public Const IS_SUBSAMPLING_HOR = IS_SUBSAMPLING_2X_HORIZONTAL


' ***********************************************
' binning
' ***********************************************
Public Const IS_GET_BINNING = &H8000&
Public Const IS_GET_SUPPORTED_BINNING = &H8001&
Public Const IS_GET_BINNING_TYPE = &H8002&

Public Const IS_BINNING_DISABLE = &H0&
Public Const IS_BINNING_VERT = &H1&
Public Const IS_BINNING_HOR = &H2&
Public Const IS_BINNING_2X_VERTICAL = IS_BINNING_VERT
Public Const IS_BINNING_4X_VERTICAL = &H4&
Public Const IS_BINNING_2X_HORIZONTAL = IS_BINNING_HOR
Public Const IS_BINNING_4X_HORIRONTAL = &H8&
Public Const IS_BINNING_COLOR = &H1&
Public Const IS_BINNING_MONO = &H2&

' ***********************************************
' Auto Control Parameter
' ***********************************************
Public Const IS_SET_ENABLE_AUTO_GAIN = &H8800&
Public Const IS_GET_ENABLE_AUTO_GAIN = &H8801&
Public Const IS_SET_ENABLE_AUTO_SHUTTER = &H8802&
Public Const IS_GET_ENABLE_AUTO_SHUTTER = &H8803&
Public Const IS_SET_ENABLE_AUTO_WHITEBALANCE = &H8804&
Public Const IS_GET_ENABLE_AUTO_WHITEBALANCE = &H8805&
Public Const IS_SET_ENABLE_AUTO_FRAMERATE = &H8806&
Public Const IS_GET_ENABLE_AUTO_FRAMERATE = &H8807&

Public Const IS_SET_AUTO_REFERENCE = &H8000&
Public Const IS_GET_AUTO_REFERENCE = &H8001&
Public Const IS_SET_AUTO_GAIN_MAX = &H8002&
Public Const IS_GET_AUTO_GAIN_MAX = &H8003&
Public Const IS_SET_AUTO_SHUTTER_MAX = &H8004&
Public Const IS_GET_AUTO_SHUTTER_MAX = &H8005&
Public Const IS_SET_AUTO_SPEED = &H8006&
Public Const IS_GET_AUTO_SPEED = &H8007&
Public Const IS_SET_AUTO_WB_OFFSET = &H8008&
Public Const IS_GET_AUTO_WB_OFFSET = &H8009&
Public Const IS_SET_AUTO_WB_GAIN_RANGE = &H800A&
Public Const IS_GET_AUTO_WB_GAIN_RANGE = &H800B&
Public Const IS_SET_AUTO_WB_SPEED = &H800C&
Public Const IS_GET_AUTO_WB_SPEED = &H800D&
Public Const IS_SET_AUTO_WB_ONCE = &H800E&
Public Const IS_GET_AUTO_WB_ONCE = &H800F&
Public Const IS_SET_AUTO_BRIGHTNESS_ONCE = &H8010&
Public Const IS_GET_AUTO_BRIGHTNESS_ONCE = &H8011&


' ***********************************************
' Auto Control definitions
' ***********************************************
Public Const IS_MIN_AUTO_BRIGHT_REFERENCE = 0
Public Const IS_MAX_AUTO_BRIGHT_REFERENCE = 255
Public Const IS_DEFAULT_AUTO_BRIGHT_REFERENCE = 128
Public Const IS_MIN_AUTO_SPEED = 0
Public Const IS_MAX_AUTO_SPEED = 100
Public Const IS_DEFAULT_AUTO_SPEED = 50
Public Const IS_DEFAULT_WB_OFFSET = 0
Public Const IS_MIN_WB_OFFSET = -50
Public Const IS_MAX_WB_OFFSET = 50
Public Const IS_DEFAULT_AUTO_WB_SPEED = 50
Public Const IS_MIN_AUTO_WB_SPEED = 0
Public Const IS_MAX_AUTO_WB_SPEED = 100
Public Const IS_MIN_AUTO_WB_REFERENCE = 0
Public Const IS_MAX_AUTO_WB_REFERENCE = 255


' ***********************************************
' AOI types to set/get
' ***********************************************
Public Const IS_SET_AUTO_BRIGHT_AOI = &H8000&
Public Const IS_GET_AUTO_BRIGHT_AOI = &H8001&
Public Const IS_SET_IMAGE_AOI = &H8002&
Public Const IS_GET_IMAGE_AOI = &H8003&
Public Const IS_SET_AUTO_WB_AOI = &H8004&
Public Const IS_GET_AUTO_WB_AOI = &H8005&


' ***********************************************
' color modes
' ***********************************************
Public Const IS_GET_COLOR_MODE = &H8000&

Public Const IS_SET_CM_RGB32 = 0
Public Const IS_SET_CM_RGB24 = 1
Public Const IS_SET_CM_RGB16 = 2
Public Const IS_SET_CM_RGB15 = 3
Public Const IS_SET_CM_Y8 = 6
Public Const IS_SET_CM_RGB8 = 7
Public Const IS_SET_CM_BAYER = 11
Public Const IS_SET_CM_UYVY = 12
Public Const IS_SET_CM_UYVY_MONO = 13
Public Const IS_SET_CM_UYVY_BAYER = 14


' ***********************************************
' badpixel correction
' ***********************************************
Public Const IS_GET_BPC_MODE = &H8000&
Public Const IS_GET_BPC_THRESHOLD = &H8001&
Public Const IS_GET_BPC_SUPPORTED_MODE = &H8002&

Public Const IS_BPC_DISABLE = 0
Public Const IS_BPC_ENABLE_LEVEL_1 = 1
Public Const IS_BPC_ENABLE_LEVEL_2 = 2
Public Const IS_BPC_ENABLE_USER = 4
Public Const IS_BPC_ENABLE_SOFTWARE = IS_BPC_ENABLE_LEVEL_2
Public Const IS_BPC_ENABLE_HARDWARE = IS_BPC_ENABLE_LEVEL_1

Public Const IS_SET_BADPIXEL_LIST = &H1&
Public Const IS_GET_BADPIXEL_LIST = &H2&
Public Const IS_GET_LIST_SIZE = &H3&


' ***********************************************
' color correction definitions
' ***********************************************
Public Const IS_GET_CCOR_MODE = &H8000&
Public Const IS_CCOR_DISABLE = &H0&
Public Const IS_CCOR_ENABLE = &H1&


' ***********************************************
'  bayer algorithm modes
' ***********************************************
Public Const IS_GET_BAYER_CV_MODE = &H8000&

Public Const IS_SET_BAYER_CV_NORMAL = &H0&
Public Const IS_SET_BAYER_CV_BETTER = &H1&
Public Const IS_SET_BAYER_CV_BEST = &H2&


' ***********************************************
' edge enhancement
' ***********************************************
Public Const IS_GET_EDGE_ENHANCEMENT = &H8000&

Public Const IS_EDGE_EN_DISABLE = 0
Public Const IS_EDGE_EN_STRONG = 1
Public Const IS_EDGE_EN_WEAK = 2


' ***********************************************
'  white balance modes
' ***********************************************
Public Const IS_GET_WB_MODE = &H8000&

Public Const IS_SET_WB_DISABLE = &H0&
Public Const IS_SET_WB_USER = &H1&
Public Const IS_SET_WB_AUTO_ENABLE = &H2&
Public Const IS_SET_WB_AUTO_ENABLE_ONCE = &H4&

Public Const IS_SET_WB_DAYLIGHT_65 = &H101&
Public Const IS_SET_WB_COOL_WHITE = &H102&
Public Const IS_SET_WB_U30 = &H103&
Public Const IS_SET_WB_ILLUMINANT_A = &H104&
Public Const IS_SET_WB_HORIZON = &H105&


' ***********************************************
' flash strobe constants
' ***********************************************
Public Const IS_GET_FLASHSTROBE_MODE = &H8000&
Public Const IS_GET_FLASHSTROBE_LINE = &H8001&
Public Const IS_GET_SUPPORTED_FLASH_IO_PORTS = &H8002&

Public Const IS_SET_FLASH_OFF = 0
Public Const IS_SET_FLASH_ON = 1
Public Const IS_SET_FLASH_LO_ACTIVE = IS_SET_FLASH_ON
Public Const IS_SET_FLASH_HI_ACTIVE = 2
Public Const IS_SET_FLASH_HIGH = 3
Public Const IS_SET_FLASH_LOW = 4
Public Const IS_SET_FLASH_LO_ACTIVE_FREERUN = 5
Public Const IS_SET_FLASH_HI_ACTIVE_FREERUN = 6
Public Const IS_SET_FLASH_IO_1 = &H10&
Public Const IS_SET_FLASH_IO_2 = &H20&
Public Const IS_SET_FLASH_IO_3 = &H40&
Public Const IS_SET_FLASH_IO_4 = &H80&
Public Const IS_FLASH_IO_PORT_MASK = (IS_SET_FLASH_IO_1 Or IS_SET_FLASH_IO_2 Or IS_SET_FLASH_IO_3 Or IS_SET_FLASH_IO_4)

Public Const IS_GET_FLASH_DELAY = -1
Public Const IS_GET_FLASH_DURATION = -2
Public Const IS_GET_MAX_FLASH_DELAY = -3
Public Const IS_GET_MAX_FLASH_DURATION = -4
Public Const IS_GET_MIN_FLASH_DELAY = -5
Public Const IS_GET_MIN_FLASH_DURATION = -6
Public Const IS_GET_FLASH_DELAY_GRANULARITY = -7
Public Const IS_GET_FLASH_DURATION_GRANULARITY = -8

' ***********************************************
' Digital IO constants
' ***********************************************
Public Const IS_GET_IO = &H8000&
Public Const IS_GET_IO_MASK = &H8000&


' ***********************************************
' EEPROM defines
' ***********************************************
Public Const IS_EEPROM_MIN_USER_ADDRESS = 0
Public Const IS_EEPROM_MAX_USER_ADDRESS = 63
Public Const IS_EEPROM_MAX_USER_SPACE = 64


' ***********************************************
' error report modes
' ***********************************************
Public Const IS_GET_ERR_REP_MODE = &H8000&
Public Const IS_DISABLE_ERR_REP = 0
Public Const IS_ENABLE_ERR_REP = 1


' ***********************************************
' display mode slectors
' ***********************************************
Public Const IS_GET_DISPLAY_MODE = &H8000&
Public Const IS_GET_DISPLAY_SIZE_X = &H8000&
Public Const IS_GET_DISPLAY_SIZE_Y = &H8001&
Public Const IS_GET_DISPLAY_POS_X = &H8000&
Public Const IS_GET_DISPLAY_POS_Y = &H8001&

Public Const IS_SET_DM_DIB = &H1&
Public Const IS_SET_DM_DIRECTDRAW = &H2&
Public Const IS_SET_DM_ALLOW_SYSMEM = &H40&
Public Const IS_SET_DM_ALLOW_PRIMARY = &H80&
' -- overlay display mode ---
Public Const IS_GET_DD_OVERLAY_SCALE = &H8000&

Public Const IS_SET_DM_ALLOW_OVERLAY = &H100&
Public Const IS_SET_DM_ALLOW_SCALING = &H200&
Public Const IS_SET_DM_ALLOW_FIELDSKIP = &H400&
Public Const IS_SET_DM_MONO = &H800&
Public Const IS_SET_DM_BAYER = &H1000&

' -- backbuffer display mode ---
Public Const IS_SET_DM_BACKBUFFER = &H2000&

' ***********************************************
' DirectDraw keying color constants
' ***********************************************
Public Const IS_GET_KC_RED = &H8000&
Public Const IS_GET_KC_GREEN = &H8001&
Public Const IS_GET_KC_BLUE = &H8002&
Public Const IS_GET_KC_RGB = &H8003&
Public Const IS_GET_KC_INDEX = &H8004&
Public Const IS_GET_KEYOFFSET_X = &H8000&
Public Const IS_GET_KEYOFFSET_Y = &H8001&

' RGB-triple for default key-color in 15,16,24,32 bit mode
Public Const IS_SET_KC_DEFAULT = &HFF00FF
' colorindex for default key-color in 8bit palette mode
Public Const IS_SET_KC_DEFAULT_8 = 253


' ***********************************************
' memoryboard
' ***********************************************
Public Const IS_MEMORY_GET_COUNT = &H8000&
Public Const IS_MEMORY_GET_DELAY = &H8001&
Public Const IS_MEMORY_MODE_DISABLE = &H0&
Public Const IS_MEMORY_USE_TRIGGER = &HFFFF&


' ***********************************************
'  test image modes
' ***********************************************
Public Const IS_GET_TEST_IMAGE = &H8000&

Public Const IS_SET_TEST_IMAGE_DISABLED = &H0&
Public Const IS_SET_TEST_IMAGE_MEMORY_1 = &H1&
Public Const IS_SET_TEST_IMAGE_MEMORY_2 = &H2&
Public Const IS_SET_TEST_IMAGE_MEMORY_3 = &H3&

' ***********************************************
' Led settings
' ***********************************************
Public Const IS_SET_LED_OFF = &H0&
Public Const IS_SET_LED_ON = &H1&
Public Const IS_SET_LED_TOGGLE = &H2&
Public Const IS_GET_LED = &H8000&

' ***********************************************
' save options
' ***********************************************
Public Const IS_SAVE_USE_ACTUAL_IMAGE_SIZE = &H10000


' ***********************************************
' event constants
' ***********************************************
Public Const IS_SET_EVENT_ODD = 0
Public Const IS_SET_EVENT_EVEN = 1
Public Const IS_SET_EVENT_FRAME = 2
Public Const IS_SET_EVENT_EXTTRIG = 3
Public Const IS_SET_EVENT_VSYNC = 4
Public Const IS_SET_EVENT_SEQ = 5
Public Const IS_SET_EVENT_STEAL = 6
Public Const IS_SET_EVENT_VPRES = 7
Public Const IS_SET_EVENT_TRANSFER_FAILED = 8
Public Const IS_SET_EVENT_DEVICE_RECONNECTED = 9
Public Const IS_SET_EVENT_MEMORY_MODE_FINISH = 10
Public Const IS_SET_EVENT_FRAME_RECEIVED = 11
Public Const IS_SET_EVENT_WB_FINISHED = 12
Public Const IS_SET_EVENT_AUTOBRIGHTNESS_FINISHED = 13

Public Const IS_SET_EVENT_REMOVE = 128
Public Const IS_SET_EVENT_REMOVAL = 129
Public Const IS_SET_EVENT_NEW_DEVICE = 130


' ***********************************************
' Window message defines
' ***********************************************
Public Const WM_USER = &H400&
Public Const IS_UEYE_MESSAGE = WM_USER + &H100& '0x0400 + 0x0100
  Public Const IS_FRAME = &H0&
  Public Const IS_SEQUENCE = &H1&
  Public Const IS_TRIGGER = &H2&
  Public Const IS_TRANSFER_FAILED = &H3&
  Public Const IS_DEVICE_RECONNECTED = &H4&
  Public Const IS_MEMORY_MODE_FINISH = &H5&
  Public Const IS_FRAME_RECEIVED = &H6&
  Public Const IS_GENERIC_ERROR = &H7&
  Public Const IS_STEAL_VIDEO = &H8&
  Public Const IS_WB_FINISHED = &H9&
  Public Const IS_AUTOBRIGHTNESS_FINISHED = &HA&

  Public Const IS_DEVICE_REMOVED = &H1000&
  Public Const IS_DEVICE_REMOVAL = &H1001&
  Public Const IS_NEW_DEVICE = &H1002&
  
  
' ***********************************************
' camera id constants
' ***********************************************
Public Const IS_GET_CAMERA_ID = &H8000&


' ***********************************************
' board info constants
' ***********************************************
Public Const IS_GET_STATUS = &H8000&

Public Const IS_EXT_TRIGGER_EVENT_CNT = 0
Public Const IS_FIFO_OVR_CNT = 1
Public Const IS_SEQUENCE_CNT = 2
Public Const IS_LAST_FRAME_FIFO_OVR = 3
Public Const IS_SEQUENCE_SIZE = 4
Public Const IS_VIDEO_PRESENT = 5
Public Const IS_STEAL_FINISHED = 6
Public Const IS_STORE_FILE_PATH = 7
Public Const IS_LUMA_BANDWIDTH_FILTER = 8
Public Const IS_BOARD_REVISION = 9
Public Const IS_MIRROR_BITMAP_UPDOWN = 10
Public Const IS_BUS_OVR_CNT = 11
Public Const IS_STEAL_ERROR_CNT = 12
Public Const IS_LOW_COLOR_REMOVAL = 13
Public Const IS_CHROMA_COMB_FILTER = 14
Public Const IS_CHROMA_AGC = 15
Public Const IS_WATCHDOG_ON_BOARD = 16
Public Const IS_PASSTHROUGH_ON_BOARD = 17
Public Const IS_EXTERNAL_VREF_MODE = 18
Public Const IS_WAIT_TIMEOUT = 19
Public Const IS_TRIGGER_MISSED = 20
Public Const IS_LAST_CAPTURE_ERROR = 21

' ***********************************************
' board type defines
' ***********************************************
Public Const IS_BOARD_TYPE_FALCON = 1
Public Const IS_BOARD_TYPE_EAGLE = 2
Public Const IS_BOARD_TYPE_FALCON2 = 3
Public Const IS_BOARD_TYPE_FALCON_PLUS = 7
Public Const IS_BOARD_TYPE_FALCON_QUATTRO = 9
Public Const IS_BOARD_TYPE_FALCON_DUO = 10
Public Const IS_BOARD_TYPE_EAGLE_QUATTRO = 11
Public Const IS_BOARD_TYPE_EAGLE_DUO = 12
Public Const IS_BOARD_TYPE_UEYE_USB = &H40&


' ***********************************************
' readable operation system defines
' ***********************************************
Public Const IS_OS_UNDETERMINED = 0
Public Const IS_OS_WIN95 = 1
Public Const IS_OS_WINNT40 = 2
Public Const IS_OS_WIN98 = 3
Public Const IS_OS_WIN2000 = 4
Public Const IS_OS_WINXP = 5
Public Const IS_OS_WINME = 6
Public Const IS_OS_WINNET = 7
Public Const IS_OS_WINSERVER2003 = 8
Public Const IS_OS_WINVISTA = 9
Public Const IS_OS_LINUX24 = 10
Public Const IS_OS_LINUX26 = 11

' ***********************************************
' usb bus speed
' ***********************************************
Public Const IS_USB_10 = 1
Public Const IS_USB_20 = 4


' ***********************************************
' sequence flags
' ***********************************************
Public Const IS_LOCK_LAST_BUFFER = &H8002&


' ***********************************************
' steal video constants
' ***********************************************
Public Const IS_INIT_STEAL_VIDEO = 1
Public Const IS_EXIT_STEAL_VIDEO = 2
Public Const IS_INIT_STEAL_VIDEO_MANUAL = 3
Public Const IS_INIT_STEAL_VIDEO_AUTO = 4
Public Const IS_SET_STEAL_RATIO = 64
Public Const IS_USE_MEM_IMAGE_SIZE = 128
Public Const IS_STEAL_MODES_MASK = 7
Public Const IS_SET_STEAL_COPY = 4096
Public Const IS_SET_STEAL_NORMAL = 8192


' ***********************************************
' AGC modes
' ***********************************************
Public Const IS_GET_AGC_MODE = &H8000&
Public Const IS_SET_AGC_OFF = 0
Public Const IS_SET_AGC_ON = 1


' ***********************************************
' gamma modes
' ***********************************************
Public Const IS_GET_GAMMA_MODE = &H8000&
Public Const IS_SET_GAMMA_OFF = 0
Public Const IS_SET_GAMMA_ON = 1


' ***********************************************
' sync levels
' ***********************************************
Public Const IS_GET_SYNC_LEVEL = &H8000&
Public Const IS_SET_SYNC_75 = 0
Public Const IS_SET_SYNC_125 = 1


' ***********************************************
' color bar modes
' ***********************************************
Public Const IS_GET_CBARS_MODE = &H8000&
Public Const IS_SET_CBARS_OFF = 0
Public Const IS_SET_CBARS_ON = 1


' ***********************************************
' horizontal filter definitions
' ***********************************************
Public Const IS_GET_HOR_FILTER_MODE = &H8000&
Public Const IS_GET_HOR_FILTER_STEP = &H8001&

Public Const IS_DISABLE_HOR_FILTER = 0
Public Const IS_ENABLE_HOR_FILTER = 1
Public Const IS_HOR_FILTER_STEP1 = 2
Public Const IS_HOR_FILTER_STEP2 = 4
Public Const IS_HOR_FILTER_STEP3 = 6


' ***********************************************
' vertical filter definitions
' ***********************************************
Public Const IS_GET_VERT_FILTER_MODE = &H8000&
Public Const IS_GET_VERT_FILTER_STEP = &H8001&

Public Const IS_DISABLE_VERT_FILTER = 0
Public Const IS_ENABLE_VERT_FILTER = 1
Public Const IS_VERT_FILTER_STEP1 = 2
Public Const IS_VERT_FILTER_STEP2 = 4
Public Const IS_VERT_FILTER_STEP3 = 6


' ***********************************************
' scaler modes
' ***********************************************
Public Const IS_GET_SCALER_MODE As Single = 1000#
Public Const IS_SET_SCALER_OFF As Single = 0#
Public Const IS_SET_SCALER_ON As Single = 1#

Public Const IS_MIN_SCALE_X As Single = 6.25
Public Const IS_MAX_SCALE_X As Single = 100#
Public Const IS_MIN_SCALE_Y As Single = 6.25
Public Const IS_MAX_SCALE_Y As Single = 100#


' ***********************************************
' video source selectors
' ***********************************************
Public Const IS_GET_VIDEO_IN = &H8000&
Public Const IS_GET_VIDEO_PASSTHROUGH = &H8000&
Public Const IS_GET_VIDEO_IN_TOGGLE = &H8001&
Public Const IS_GET_TOGGLE_INPUT_1 = &H8000&
Public Const IS_GET_TOGGLE_INPUT_2 = &H8001&
Public Const IS_GET_TOGGLE_INPUT_3 = &H8002&
Public Const IS_GET_TOGGLE_INPUT_4 = &H8003&

Public Const IS_SET_VIDEO_IN_1 = &H0&
Public Const IS_SET_VIDEO_IN_2 = &H1&
Public Const IS_SET_VIDEO_IN_S = &H2&
Public Const IS_SET_VIDEO_IN_3 = &H3&
Public Const IS_SET_VIDEO_IN_4 = &H4&
Public Const IS_SET_VIDEO_IN_1S = &H10&
Public Const IS_SET_VIDEO_IN_2S = &H11&
Public Const IS_SET_VIDEO_IN_3S = &H13&
Public Const IS_SET_VIDEO_IN_4S = &H14&
Public Const IS_SET_TOGGLE_OFF = &HFF&
Public Const IS_SET_VIDEO_IN_SYNC = &H4000&
Public Const IS_VIDEO_IN_MASK = &H7&
Public Const IS_VIDEO_IN_S_MASK = &H17&


' ***********************************************
' video crossbar selectors
' ***********************************************
Public Const IS_GET_CROSSBAR = &H8000&

Public Const IS_CROSSBAR_1 = 0
Public Const IS_CROSSBAR_2 = 1
Public Const IS_CROSSBAR_3 = 2
Public Const IS_CROSSBAR_4 = 3
Public Const IS_CROSSBAR_5 = 4
Public Const IS_CROSSBAR_6 = 5
Public Const IS_CROSSBAR_7 = 6
Public Const IS_CROSSBAR_8 = 7
Public Const IS_CROSSBAR_9 = 8
Public Const IS_CROSSBAR_10 = 9
Public Const IS_CROSSBAR_11 = 10
Public Const IS_CROSSBAR_12 = 11
Public Const IS_CROSSBAR_13 = 12
Public Const IS_CROSSBAR_14 = 13
Public Const IS_CROSSBAR_15 = 14
Public Const IS_CROSSBAR_16 = 15
Public Const IS_SELECT_AS_INPUT = 128


' ***********************************************
' video format selectors
' ***********************************************
Public Const IS_GET_VIDEO_MODE = &H8000&

Public Const IS_SET_VM_PAL = 0
Public Const IS_SET_VM_NTSC = 1
Public Const IS_SET_VM_SECAM = 2
Public Const IS_SET_VM_AUTO = 3


' ***********************************************
' capture Modes
' ***********************************************
Public Const IS_GET_CAPTURE_MODE = &H8000&

Public Const IS_SET_CM_ODD = &H1&
Public Const IS_SET_CM_EVEN = &H2&
Public Const IS_SET_CM_FRAME = &H4&
Public Const IS_SET_CM_NONINTERLACED = &H8&
Public Const IS_SET_CM_NEXT_FRAME = &H10&
Public Const IS_SET_CM_NEXT_FIELD = &H20&
Public Const IS_SET_CM_BOTHFIELDS = &HB&
Public Const IS_SET_CM_FRAME_STEREO = &H2004&


' ***********************************************
' display update mode constants
' ***********************************************
Public Const IS_GET_UPDATE_MODE = &H8000&
Public Const IS_SET_UPDATE_TIMER = 1
Public Const IS_SET_UPDATE_EVENT = 2


' ***********************************************
' sync generator mode constants
' ***********************************************
Public Const IS_GET_SYNC_GEN = &H8000&
Public Const IS_SET_SYNC_GEN_OFF = 0
Public Const IS_SET_SYNC_GEN_ON = 1


' ***********************************************
' decimation modes
' ***********************************************
Public Const IS_GET_DECIMATION_MODE = &H8000&
Public Const IS_GET_DECIMATION_NUMBER = &H8001&

Public Const IS_DECIMATION_OFF = 0
Public Const IS_DECIMATION_CONSECUTIVE = 1
Public Const IS_DECIMATION_DISTRIBUTED = 2


' ***********************************************
' hardware watchdog defines
' ***********************************************
Public Const IS_GET_WATCHDOG_TIME = &H2000&
Public Const IS_GET_WATCHDOG_RESOLUTION = &H4000&
Public Const IS_GET_WATCHDOG_ENABLE = &H8000&

Public Const IS_WATCHDOG_MINUTES = 0
Public Const IS_WATCHDOG_SECONDS = &H8000&
Public Const IS_DISABLE_WATCHDOG = 0
Public Const IS_ENABLE_WATCHDOG = 1
Public Const IS_RETRIGGER_WATCHDOG = 2
Public Const IS_ENABLE_AUTO_DEACTIVATION = 4
Public Const IS_DISABLE_AUTO_DEACTIVATION = 8
Public Const IS_WATCHDOG_RESERVED = &H1000&


' ***********************************************
' Global Shutter definitions
' ***********************************************
Public Const IS_SET_GLOBAL_SHUTTER_ON = &H1&
Public Const IS_SET_GLOBAL_SHUTTER_OFF = &H0&
Public Const IS_GET_GLOBAL_SHUTTER = &H10&
Public Const IS_GET_SUPPORTED_GLOBAL_SHUTTER = &H20&


' ***********************************************
' Image files types
' ***********************************************
Public Const IS_IMG_BMP = 0
Public Const IS_IMG_JPG = 1
Public Const IS_IMG_PNG = 2
Public Const IS_IMG_RAW = 4


' ***********************************************
' typedefs
' ***********************************************
Public HIDS As Long
Public hCam As Long
Public HFALC As Long


' ***********************************************
' BOARDINFO structure
' ***********************************************
Public Type BoardInfo
    SerNo As String * 12            ' e.g. "1234512345"  (11 char)
    ID As String * 20               ' e.g. "IDS GmbH"
    Version As String * 10          ' e.g. "V2.10"  (9 char)
    Datum As String * 12            ' e.g. "24.01.2006" (11 char)
    Select As Byte                  ' contains board select number for multi board support
    Type As Byte                    ' e.g. IS_BOARD_TYPE_UEYE_USB
    Reserverd(8) As Byte
End Type


' ***********************************************
' CAMERAINFO structure
' ***********************************************
Public Type CameraInfo
    SerNo As String * 12            ' e.g. "1234512345"  (11 char)
    ID As String * 20               ' e.g. "IDS GmbH"
    Version As String * 10          ' e.g. "V2.10"  (9 char)
    Datum As String * 12            ' e.g. "24.01.2006" (11 char)
    Select As Byte                  ' contains board select number for multi board support
    Type As Byte                    ' e.g. IS_BOARD_TYPE_UEYE_USB
    Reserverd(8) As Byte
End Type



' ***********************************************
' SENSORINFO structure
' ***********************************************
Public Type SensorInfo
    SensorID As Integer             ' e.g. IS_SENSOR_UI224X_C
    strSensorName As String * 32    ' e.g. "UI-224X-C"
    nColorMode As Byte              ' e.g. IS_COLORMODE_BAYER
    nMaxWidth As Long               ' e.g. 1280
    nMaxHeight As Long              ' e.g. 1024
    bMasterGain As Boolean          ' e.g. True
    bRGain As Boolean               ' e.g. True
    bGGain As Boolean               ' e.g. True
    bBGain As Boolean               ' e.g. True
    bGlobShutter As Boolean         ' e.g. True
    reserved(16) As Byte
End Type


' ***********************************************
' REVISIONINFO structure
' ***********************************************
Public Type RevisionInfo
    size As Integer                 ' 2
    Sensor As Integer               ' 2
    Cypress As Integer              ' 2
    Blackfin As Integer             ' 4
    DspFirmware As Integer          ' 2
                                    ' --12
    USB_Board As Integer            ' 2
    Sensor_Board As Integer         ' 2
    Processing_Board As Integer     ' 2
    Memory_Board As Integer         ' 2
    Housing As Integer              ' 2
    Filter As Integer               ' 2
    Timing_Board As Integer         ' 2
    Product As Integer              ' 2
                                    ' --24
    reserved(100) As Byte           ' --128
End Type


' ***********************************************
' UEYE_CAMERA_INFO + UEYE_CAMERA_LIST structure
' ***********************************************
Public Type UEYE_CAMERA_INFO
    dwCameraID As Long              ' this is the user defineable camera ID
    dwDeviceID As Long              ' this is the systems enumeration ID
    dwSensorID As Long              ' this is the sensor ID e.g. IS_SENSOR_UI141X_M
    dwInUse As Long                 ' flag, whether the camera is in use or not
    SerNo As String * 16            ' serial numer of the camera
    Model As String * 16            ' model name of the camera
    dwReserved(16) As Long
End Type


' usage of the list:
' 1. call the DLL with .dwCount = 0
' 2. DLL returns .dwCount = N  (N = number of available cameras)
' 3. call DLL with .dwCount = N and a pointer to UEYE_CAMERA_LIST with
'    and array of UEYE_CAMERA_INFO[N]
' 4. DLL will fill in the array with the camera infos and
'    will update the .dwCount member with the actual number of cameras
'    because there may be a change in number of cameras between step 2 and 3
' 5. check if there's a difference in actual .dwCount and formerly
'    reported value of N and call DLL again with an updated array size

Public Type UEYE_CAMERA_LIST
    dwCount As Long
    uci(1) As UEYE_CAMERA_INFO
End Type


' ***********************************************
' auto feature structs and definitions
' ***********************************************
Public Const AC_SHUTTER = &H1&
Public Const AC_GAIN = &H2&
Public Const AC_WHITEBAL = &H4&
Public Const AC_WB_RED_CHANNEL = &H8&
Public Const AC_WB_GREEN_CHANNEL = &H10&
Public Const AC_WB_BLUE_CHANNEL = &H20&

Public Const ACS_ADJUSTING = &H1&
Public Const ACS_FINISHED = &H2&
Public Const ACS_DISABLED = &H4&

Public Type AUTO_BRIGHT_STATUS
    curValue As Long                    ' current average greylevel
    curError As Long                    ' current auto brightness error
    curController As Long               ' current active brightness controller -> AC_x
    curCtrlStatus As Long               ' current control status -> ACS_x
End Type

Public Type AUTO_WB_CHANNEL_STATUS
    curValue As Long                    ' current average greylevel
    curError As Long                    ' current auto wb error
    curCtrlStatus As Long               ' current control status -> ACS_x
End Type

Public Type AUTO_WB_STATUS
    RedChannel As AUTO_WB_CHANNEL_STATUS
    GreenChannel As AUTO_WB_CHANNEL_STATUS
    BlueChannel As AUTO_WB_CHANNEL_STATUS
    curController As Long               ' current active wb controller -> AC_x
End Type

Public Type UEYE_AUTO_INFO
    AutoAbility As Long                 ' autocontrol ability
    sBrightCtrlStatus As AUTO_BRIGHT_STATUS ' brightness autocontrol status
    sWBCtrlStatus As AUTO_WB_STATUS     ' white balance autocontrol status
    reserved(12) As Long
End Type


' ***********************************************
' exports from uEye_api.dll
' ***********************************************


' ***********************************************
' functions only effective on Falcon boards
' ***********************************************
Public Declare Function iss_SetVideoInput Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal Source As Long) As Long

Public Declare Function iss_SetSaturation Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal ChromU As Long, ByVal ChromV As Long) As Long

Public Declare Function iss_SetHue Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal Hue As Long) As Long

Public Declare Function iss_SetVideoMode Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal mode As Long) As Long

Public Declare Function iss_SetAGC Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal mode As Long) As Long

Public Declare Function iss_SetSyncLevel Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal Level As Long) As Long

Public Declare Function iss_ShowColorBars Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal mode As Long) As Long

Public Declare Function iss_SetScaler Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal ScaleX As Single, ByVal ScaleY As Single) As Long

Public Declare Function iss_SetCaptureMode Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal mode As Long) As Long

Public Declare Function iss_SetHorFilter Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal mode As Long) As Long

Public Declare Function iss_SetVertFilter Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal mode As Long) As Long

Public Declare Function iss_ScaleDDOverlay Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal boScale As Long) As Long

Public Declare Function iss_GetCurrentField Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef pField As Long) As Long

Public Declare Function iss_SetVideoSize Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal xPos As Long, ByVal yPos As Long, ByVal xsize As Long, ByVal ysize As Long) As Long

Public Declare Function iss_SetKeyOffset Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nOffsetX As Long, ByVal nOffsetY As Long) As Long

Public Declare Function iss_PrepareStealVideo Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal mode As Integer, ByVal StealColorMode As Long) As Long

Public Declare Function iss_SetParentHwnd Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal hWnd As Long) As Long

Public Declare Function iss_SetUpdateMode Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal mode As Long) As Long

Public Declare Function iss_OvlSurfaceOffWhileMove Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal boMode As Long) As Long

Public Declare Function iss_GetPciSlot Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef pnSlot As Long) As Long

Public Declare Function iss_GetIRQ Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef pnIRQ As Long) As Long

Public Declare Function iss_SetToggleMode Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal Input1 As Long, ByVal Input2 As Long, ByVal Input3 As Long, ByVal Input4 As Long) As Long

Public Declare Function iss_SetDecimationMode Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal mode As Long, ByVal Decimate As Long) As Long

Public Declare Function iss_SetSync Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nSync As Long) As Long

' only FALCON duo/quattro:
Public Declare Function iss_SetVideoCrossbar Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal VidIn As Long, ByVal VidOut As Long) As Long

Public Declare Function iss_WatchdogTime Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal lTime As Long) As Long

Public Declare Function iss_Watchdog Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal lMode As Long) As Long

Public Declare Function iss_SetPassthrough Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nSource As Long) As Long


' ***********************************************
' alias functions for compatibility
' ***********************************************
Public Declare Function iss_InitBoard Lib "uEye_api.dll" _
(ByRef hCam As Long, ByVal hWnd As Long) As Long

Public Declare Function iss_ExitBoard Lib "uEye_api.dll" _
(ByVal hCam As Long) As Long

Public Declare Function iss_InitFalcon Lib "uEye_api.dll" _
(ByRef hCam As Long, ByVal hWnd As Long) As Long

Public Declare Function iss_ExitFalcon Lib "uEye_api.dll" _
(ByVal hCam As Long) As Long


Public Declare Function iss_GetBoardType Lib "uEye_api.dll" _
(ByVal hCam As Long) As Long

Public Declare Function iss_GetBoardInfo Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef pBoardinfo As BoardInfo) As Long

Public Declare Function iss_BoardStatus Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nInfo As Long, ByVal nValue As Long) As Long

Public Declare Function iss_GetNumberOfDevices Lib "uEye_api.dll" () As Long

Public Declare Function iss_GetNumberOfBoards Lib "uEye_api.dll" _
(ByRef pnNumBoards As Long) As Long


' ***********************************************
' common functions
' ***********************************************
Public Declare Function iss_StopLiveVideo Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal mode As Long) As Long

Public Declare Function iss_FreezeVideo Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal Wait As Long) As Long

Public Declare Function iss_CaptureVideo Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal captureMode As Long) As Long

Public Declare Function iss_IsVideoFinish Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef pBool As Long) As Long

Public Declare Function iss_HasVideoStarted Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef pBool As Long) As Long


Public Declare Function iss_SetBrightness Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal bright As Long) As Long

Public Declare Function iss_SetContrast Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal Cont As Long) As Long

Public Declare Function iss_SetGamma Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal mode As Long) As Long


Public Declare Function iss_AllocImageMem Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal width As Long, ByVal height As Long, ByVal bpp As Long, _
ByRef pImgMem As Long, ByRef ID As Long) As Long

Public Declare Function iss_SetImageMem Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal pImgMem As Long, ByVal ID As Long) As Long

Public Declare Function iss_FreeImageMem Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal pImgMem As Long, ByVal ID As Long) As Long

Public Declare Function iss_GetImageMem Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef pImgMem As Long) As Long

Public Declare Function iss_GetActiveImageMem Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef ppcMem As Long, ByRef pnID As Long) As Long

Public Declare Function iss_InquireImageMem Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal pcMem As Long, ByVal nID As Long, ByRef pnX As Long, ByRef pnY As Long, ByRef pnBits As Long, ByRef pnPitch As Long) As Long

Public Declare Function iss_GetImageMemPitch Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef pPitch As Long) As Long


Public Declare Function iss_SetAllocatedImageMem Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal width As Long, ByVal height As Long, ByVal bpp As Long, _
ByVal pImgMem As Long, ByRef ID As Long) As Long

Public Declare Function iss_SaveImageMem Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal pcFileString As Any, ByVal pcMem As Long, ByVal nID As Long) As Long

Public Declare Function iss_CopyImageMem Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal pcSource As Long, ByVal nID As Long, ByVal pcDest As Long) As Long

Public Declare Function iss_CopyImageMemLines Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal pcSource As Long, ByVal nID As Long, ByVal nLines As Long, ByVal pcDest As Long) As Long


Public Declare Function iss_AddToSequence Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef pcMem As Long, ByVal nID As Long) As Long

Public Declare Function iss_ClearSequence Lib "uEye_api.dll" _
(ByVal hCam As Long) As Long

Public Declare Function iss_GetActSeqBuf Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef pnNum As Long, ByRef ppcMem As Long, ByRef ppcMemLast As Long) As Long

Public Declare Function iss_LockSeqBuf Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nNum As Long, ByVal pcMem As Long) As Long

Public Declare Function iss_UnlockSeqBuf Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nNum As Long, ByVal pcMem As Long) As Long


Public Declare Function iss_SetImageSize Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal X As Long, ByVal Y As Long) As Long

Public Declare Function iss_SetImagePos Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal X As Long, ByVal Y As Long) As Long


Public Declare Function iss_GetError Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef pErrNo As Long, ByRef pErrString As Long) As Long

Public Declare Function iss_SetErrorReport Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal mode As Long) As Long


Public Declare Function iss_ReadEEPROM Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal Adr As Long, ByRef pcByte As Byte, ByVal Count As Long) As Long

Public Declare Function iss_WriteEEPROM Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal Adr As Long, ByRef pcByte As Byte, ByVal Count As Long) As Long

Public Declare Function iss_SaveImage Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal pcFileString As Any) As Long


Public Declare Function iss_SetColorMode Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal ColorMode As Long) As Long

Public Declare Function iss_GetColorDepth Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef pnCol As Long, ByRef pnColMode As Long) As Long

Public Declare Function iss_RenderBitmap Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal MemID As Long, ByVal hWnd As Long, ByVal mode As Long) As Long


Public Declare Function iss_SetDisplayMode Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal ColorMode As Long) As Long

Public Declare Function iss_GetDC Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef phDC As Long) As Long

Public Declare Function iss_ReleaseDC Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal hDC As Long) As Long

Public Declare Function iss_UpdateDisplay Lib "uEye_api.dll" _
(ByVal hCam As Long) As Long

Public Declare Function iss_SetDisplaySize Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal X As Long, ByVal Y As Long) As Long

Public Declare Function iss_SetDisplayPos Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal X As Long, ByVal Y As Long) As Long


Public Declare Function iss_LockDDOverlayMem Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef ppMem As Long, ByRef pPitch As Long) As Long

Public Declare Function iss_UnlockDDOverlayMem Lib "uEye_api.dll" _
(ByVal hCam As Long) As Long

Public Declare Function iss_LockDDMem Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef ppMem As Long, ByRef pPitch As Long) As Long

Public Declare Function iss_UnlockDDMem Lib "uEye_api.dll" _
(ByVal hCam As Long) As Long

Public Declare Function iss_GetDDOvlSurface Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef ppDDSurf As Any) As Long

Public Declare Function iss_SetKeyColor Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal red As Long, ByVal green As Long, ByVal blue As Long) As Long

Public Declare Function iss_StealVideo Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal Wait As Integer) As Long

Public Declare Function iss_SetHwnd Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal hWnd As Long) As Long


Public Declare Function iss_SetDDUpdateTime Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal ms As Long) As Long

Public Declare Function iss_EnableDDOverlay Lib "uEye_api.dll" _
(ByVal hCam As Long) As Long

Public Declare Function iss_DisableDDOverlay Lib "uEye_api.dll" _
(ByVal hCam As Long) As Long

Public Declare Function iss_ShowDDOverlay Lib "uEye_api.dll" _
(ByVal hCam As Long) As Long

Public Declare Function iss_HideDDOverlay Lib "uEye_api.dll" _
(ByVal hCam As Long) As Long



Public Declare Function iss_GetVsyncCount Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef pIntr As Long, ByRef pActIntr As Long) As Long

Public Declare Function iss_GetOsVersion Lib "uEye_api.dll" () As Long

Public Declare Function iss_GetDLLVersion Lib "uEye_api.dll" () As Long


Public Declare Function iss_InitEvent Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal hEv As Long, ByVal which As Long) As Long

Public Declare Function iss_ExitEvent Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal which As Long) As Long

Public Declare Function iss_EnableEvent Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal which As Long) As Long

Public Declare Function iss_DisableEvent Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal which As Long) As Long


Public Declare Function iss_SetIO Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nIO As Long) As Long

Public Declare Function iss_SetFlashStrobe Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nField As Long, ByVal nField As Long) As Long

Public Declare Function iss_SetExternalTrigger Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nTriggerMode As Long) As Long

Public Declare Function iss_SetRopEffect Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal effect As Long, ByVal Param As Long, ByVal reserved As Long) As Long


' ***********************************************
' new functions only valid for uEye
' ***********************************************
' camera functions
Public Declare Function iss_InitCamera Lib "uEye_api.dll" _
(ByRef hCam As Long, ByVal hWnd As Long) As Long

Public Declare Function iss_ExitCamera Lib "uEye_api.dll" _
(ByVal hCam As Long) As Long

Public Declare Function iss_GetCameraInfo Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef pInfo As CameraInfo) As Long

Public Declare Function iss_CameraStatus Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nInfo As Long, ByVal ulValue As Long) As Long

Public Declare Function iss_GetCameraType Lib "uEye_api.dll" _
(ByVal hCam As Long) As Long

Public Declare Function iss_GetNumberOfCameras Lib "uEye_api.dll" _
(ByRef pnNumCams As Long) As Long


' PixelClock
Public Declare Function iss_GetPixelClockRange Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef pnMin As Long, ByRef pnMax As Long) As Long

Public Declare Function iss_SetPixelClock Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal Clock As Long) As Long

Public Declare Function iss_GetUsedBandwidth Lib "uEye_api.dll" _
(ByVal hCam As Long) As Long

' framerate
Public Declare Function iss_GetFrameTimeRange Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef min As Double, ByRef max As Double, ByRef intervall As Double) As Long
 
Public Declare Function iss_SetFrameRate Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal FPS As Double, ByRef newFPS As Double) As Long

' set/get exposure
Public Declare Function iss_GetExposureRange Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef min As Double, ByRef max As Double, ByRef intervall As Double) As Long

Public Declare Function iss_SetExposureTime Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal EXP As Double, ByRef newEXP As Double) As Long

' get frames per second
Public Declare Function iss_GetFramesPerSecond Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef dblFPS As Double) As Long
             
        
' setIO mask
Public Declare Function iss_SetIOMask Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nMask As Long) As Long


' Get Sensorinfo
Public Declare Function iss_GetSensorInfo Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef pSensorInfo As SensorInfo) As Long

' Get RevisionInfo
Public Declare Function iss_GetRevisionInfo Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef prevInfo As RevisionInfo) As Long
  
  
' enable/disable auto exit after device remove
Public Declare Function iss_EnableAutoExit Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef nMode As Long) As Long

' message
Public Declare Function iss_EnableMessage Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal which As Long, ByVal hWnd As Long) As Long
  
    
' hardware gain settings
Public Declare Function iss_SetHardwareGain Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nMaster As Long, ByVal nRed As Long, ByVal nGreen As Long, ByVal nBlue As Long) As Long


' set render mode
Public Declare Function iss_SetRenderMode Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal mode As Long) As Long


' enable/disable white balance
Public Declare Function iss_SetWhiteBalance Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nMode As Long) As Long

Public Declare Function iss_SetWhiteBalanceMultipliers Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal dblRed As Double, ByVal dblGreen As Double, ByVal dblBlue As Double) As Long

Public Declare Function iss_GetWhiteBalanceMultipliers Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef pdblRed As Double, ByRef pdblGreen As Double, ByRef pdblBlue As Double) As Long

' edge enhancement
Public Declare Function iss_SetEdgeEnhancement Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nEnable As Long) As Long


' sensor features
Public Declare Function iss_SetColorCorrection Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nEnable As Long, ByRef factors As Double) As Long
 
Public Declare Function iss_SetBlCompensation Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nEnable As Long, ByVal offset As Long, ByVal reserved As Long) As Long

 
' hot pixel correction
Public Declare Function iss_SetBadPixelCorrection Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nEnable As Long, ByVal threshold As Long) As Long

Public Declare Function iss_LoadBadPixelCorrectionTable Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal pcFileString As Any) As Long

Public Declare Function iss_SaveBadPixelCorrectionTable Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal pcFileString As Any) As Long

Public Declare Function iss_SetBadPixelCorrectionTable Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nMode As Long, ByRef pList As Integer) As Long


' memoryboard
Public Declare Function iss_SetMemoryMode Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nCount As Long, ByVal nDelay As Long) As Long

Public Declare Function iss_TransferImage Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nMemID As Long, ByVal seqID As Long, ByVal imageNr As Long, ByVal reserved As Long) As Long

Public Declare Function iss_TransferMemorySequence Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal seqID As Long, ByVal StartNr As Long, ByVal nCount As Long, ByVal nSeqPos As Long) As Long

Public Declare Function iss_MemoryFreezeVideo Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nMemID As Long, ByVal Wait As Long) As Long

Public Declare Function iss_GetLastMemorySequence Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef pid As Long) As Long

Public Declare Function iss_GetNumberOfMemoryImages Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nID As Long, ByRef pnCount As Long) As Long

Public Declare Function iss_GetMemorySequenceWindow Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nID As Long, ByRef left As Long, ByRef top As Long, ByRef right As Long, ByRef bottom As Long) As Long

Public Declare Function iss_IsMemoryBoardConnected Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef pConnected As Long) As Long

Public Declare Function iss_ResetMemory Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nReserved As Long) As Long


Public Declare Function iss_SetSubSampling Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal mode As Long) As Long

Public Declare Function iss_ForceTrigger Lib "uEye_api.dll" _
(ByVal hCam As Long) As Long


' new with driver version 1.12.0006
Public Declare Function iss_GetBusSpeed Lib "uEye_api.dll" _
(ByVal hCam As Long) As Long

' new with driver version 1.12.0015
Public Declare Function iss_SetBinning Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal mode As Long) As Long


' new with driver version 1.12.0017
Public Declare Function iss_ResetToDefault Lib "uEye_api.dll" _
(ByVal hCam As Long) As Long

Public Declare Function iss_LoadParameters Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal pcFileString As Any) As Long

Public Declare Function iss_SaveParameters Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal pcFileString As Any) As Long


' new with driver version 1.14.0001
Public Declare Function iss_GetGlobalFlashDelays Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef Delay As Long, ByRef Duration As Long) As Long

Public Declare Function iss_SetFlashDelay Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal Delay As Long, ByVal Duration As Long) As Long

' new with driver version 1.14.0002
Public Declare Function iss_LoadImage Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal pcFileString As Any) As Long


' new with driver version 1.14.0008
Public Declare Function iss_SetImageAOI Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal xPos As Long, ByVal yPos As Long, _
ByVal width As Long, ByVal height As Long) As Long

Public Declare Function iss_SetCameraID Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal ID As Long) As Long

Public Declare Function iss_SetBayerConversion Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal mode As Long) As Long

Public Declare Function iss_SetTestImage Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal mode As Long) As Long

' new with driver version 1.14.0009
Public Declare Function iss_SetHardwareGamma Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nMode As Long) As Long


' new with driver version 2.00.0001
Public Declare Function iss_GetCameraList Lib "uEye_api.dll" _
(ByRef pucl As UEYE_CAMERA_LIST) As Long


' new with driver version 2.00.0011
Public Declare Function iss_SetAOI Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nType As Long, ByRef pXPos As Long, ByRef pYPos As Long, ByRef pWidth As Long, ByRef pHeight As Long) As Long

Public Declare Function iss_SetAutoParameter Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal Param As Long, ByRef pval1 As Double, ByRef pval2 As Double) As Long
'(ByVal hCam As Long, ByVal Param As Long, ByVal pval1 As Long, ByVal pval2 As Long) As Long

Public Declare Function iss_GetAutoInfo Lib "uEye_api.dll" _
(ByVal hCam As Long, ByRef pInfo As UEYE_AUTO_INFO) As Long

Public Declare Function iss_ConvertImage Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal pcSource As Long, ByVal nIDSource As Long, _
 ByRef pcDest As Long, ByRef nIDSource As Long, ByRef reserve As Long) As Long

Public Declare Function iss_SetConvertParam Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal ColorCorrection As Long, ByVal BayerConversionMode As Long, _
 ByVal ColorMode As Long, ByVal Gamma As Long, ByRef WhiteBalanceMultipliers As Double) As Long
  
Public Declare Function iss_SaveImageEx Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal pcFileString As Any, ByVal fileFormat As Long, ByVal Param As Long) As Long
  
Public Declare Function iss_SaveImageMemEx Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal pcFileString As Any, ByVal pcMem As Long, ByVal nID As Long, ByVal fileFormat As Long, ByVal Param As Long) As Long

Public Declare Function iss_LoadImageMem Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal pcFileString As Any, ByRef ppcImgMem As Long, ByRef pid As Long) As Long
  
Public Declare Function iss_GetImageHistogram Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nID As Long, ByVal ColorMode As Long, ByRef pHistoMem As Long) As Long
 
Public Declare Function iss_SetGainBoost Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nMode As Long) As Long

Public Declare Function iss_SetLED Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nMode As Long) As Long

Public Declare Function iss_SetGlobalShutter Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nMode As Long) As Long

Public Declare Function iss_SetExtendedRegister Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal Index As Long, value As Integer) As Long

' new with driver version 2.22.0002
Public Declare Function iss_SetHWGainFactor Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nMode As Long, nIndex As Long) As Long

Public Declare Function iss_SetTriggerDelay Lib "uEye_api.dll" _
(ByVal hCam As Long, ByVal nTriggerDelay As Long) As Long
  

' ***********************************************
' Win32API deklarationen
' ***********************************************
Private Const BITSPIXEL = 12         '  number of bits per pixel

Private Declare Function GetDC Lib "user32" _
(ByVal hWnd As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" _
(ByVal hDC As Long, ByVal nIndex As Long) As Long

Public hG As Long
Public hC As Long

' ***********************************************************************
' Start the DirectDraw live video in the corresponding windows color mode
' ***********************************************************************
Public Function StartLiveVideo(ByVal hCam As Long) As Boolean
    Dim bResult As Boolean
    bResult = False
    Do
        If IS_SUCCESS <> iss_SetColorMode(hCam, GetColorModeForDisplayBpp()) Then
            Exit Do
        End If

        If IS_SUCCESS <> iss_SetDisplayMode(hCam, IS_SET_DM_DIRECTDRAW Or IS_SET_DM_ALLOW_PRIMARY) Then
            Exit Do
        End If

        If IS_SUCCESS <> iss_CaptureVideo(hCam, IS_WAIT) Then
            Exit Do
        End If
        bResult = True
    Loop While False
    
    StartLiveVideo = bResult
End Function

' ***********************************************************************
' Start the DirectDraw live video Extented
'
' The StartLiveVideoEx allows to select the color mode and the display mode
' The StartLiveVideo   does not have this option
'
' Note:     In case an other color mode than the selected Windows colors
'           (see your grapgic adaptor) is chosen the IDS display mode
'           must be set to IS_SET_DM_DIB.
' ***********************************************************************
Public Function StartLiveVideoEx(ByVal hCam As Long, ByVal ColorMode As Integer, ByVal DisplayMode As Integer) As Boolean
    Dim bResult As Boolean
    bResult = False
    Do
        If IS_SUCCESS <> iss_SetColorMode(hCam, ColorMode) Then
            Exit Do
        End If

        If IS_SUCCESS <> iss_SetDisplayMode(hCam, DisplayMode) Then
            Exit Do
        End If

        If IS_SUCCESS <> iss_CaptureVideo(hCam, IS_WAIT) Then
            Exit Do
        End If
        bResult = True
    Loop While False
    
    StartLiveVideoEx = bResult
End Function

' ***********************************************************************
' Calculate the Windows color bit depth to the corredponding
' uEye color mode
' ***********************************************************************
Private Function BppToColorMode(ByVal nBpp As Integer) As Long
    Select Case nBpp
    Case 8:
        BppToColorMode = IS_SET_CM_RGB8
    Case 15:
        BppToColorMode = IS_SET_CM_RGB15
    Case 16:
        BppToColorMode = IS_SET_CM_RGB16
    Case 24:
        BppToColorMode = IS_SET_CM_RGB24
    Case 32:
        BppToColorMode = IS_SET_CM_RGB32
    Case Else
        BppToColorMode = 0
    End Select
End Function

' ***********************************************************************
' Query the color depth of the Windows screen => USE API FUNCTION INSTEAD
' ***********************************************************************
Private Function GetPixelDepth() As Long
    Dim hDC As Long
    hDC = GetDC(0)
    GetPixelDepth = GetDeviceCaps(hDC, BITSPIXEL)
End Function

' ***********************************************************************
' Query the Windows screen color depth of the corresponding
' DFG/LC1 color mode => USE API FUNCTION INSTEAD
' ***********************************************************************
Private Function GetColorModeForDisplayBpp() As Long
    GetColorModeForDisplayBpp = BppToColorMode(GetPixelDepth())
End Function


