VERSION 5.00
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0ACBBD11-6E21-11D4-9751-0060089571FC}#1.0#0"; "CogDisplay.ocx"
Begin VB.Form frmSetting 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  '단일 고정
   Caption         =   "SETTING"
   ClientHeight    =   13935
   ClientLeft      =   75
   ClientTop       =   1305
   ClientWidth     =   19185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13935
   ScaleMode       =   0  '사용자
   ScaleWidth      =   53633.67
   Begin CognexDisplay.CogDisplay CogDisplay 
      Height          =   10755
      Left            =   240
      TabIndex        =   45
      Top             =   810
      Width           =   14385
      _cx             =   25374
      _cy             =   18971
      BackColor       =   6556170
      HorizontalScrollBar=   0   'False
      VerticalScrollBar=   0   'False
      PopupMenu       =   -1  'True
      MouseMode       =   0
      ScalingMethod   =   1
      PanX            =   0
      PanY            =   0
      Zoom            =   1
      GridColor       =   12582912
      GridEnabled     =   -1  'True
      SubPixelGridEnabled=   -1  'True
      SubPixelGridColor=   32768
      TouchDistance   =   3
      InteractiveGraphicTipsEnabled=   -1  'True
      MultiSelectionEnabled=   0   'False
      _ipam           =   1
      _spam           =   1
      Enabled         =   -1  'True
      AutoFit         =   -1  'True
   End
   Begin VB.CheckBox chkNsdRegionSelection 
      BackColor       =   &H8000000E&
      Caption         =   "NSD ROI 선택(체크해제시 좌측, 체크시 우측)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   9720
      TabIndex        =   109
      Top             =   12600
      Value           =   1  '확인
      Width           =   4845
   End
   Begin VB.Frame ROI 
      Caption         =   "ROI선택"
      Height          =   765
      Left            =   0
      TabIndex        =   100
      Top             =   0
      Width           =   6315
      Begin VB.OptionButton optROI 
         Caption         =   "기본영역"
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   107
         Top             =   300
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optROI 
         Caption         =   "ROI1"
         Height          =   285
         Index           =   1
         Left            =   1365
         TabIndex        =   106
         Top             =   300
         Width           =   825
      End
      Begin VB.OptionButton optROI 
         Caption         =   "ROI2"
         Height          =   285
         Index           =   2
         Left            =   2145
         TabIndex        =   105
         Top             =   300
         Width           =   825
      End
      Begin VB.OptionButton optROI 
         Caption         =   "ROI3"
         Height          =   285
         Index           =   3
         Left            =   2910
         TabIndex        =   104
         Top             =   300
         Width           =   825
      End
      Begin VB.OptionButton optROI 
         Caption         =   "ROI4"
         Height          =   285
         Index           =   4
         Left            =   3675
         TabIndex        =   103
         Top             =   300
         Width           =   825
      End
      Begin VB.OptionButton optROI 
         Caption         =   "ROI5"
         Height          =   285
         Index           =   5
         Left            =   4455
         TabIndex        =   102
         Top             =   300
         Width           =   825
      End
      Begin VB.OptionButton optROI 
         Caption         =   "Dummy"
         Height          =   285
         Index           =   6
         Left            =   5220
         TabIndex        =   101
         Top             =   300
         Width           =   1005
      End
   End
   Begin VB.Frame fraCalibration2 
      BackColor       =   &H80000004&
      Caption         =   "카메라 간 간격"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      Left            =   14730
      TabIndex        =   61
      Top             =   6030
      Width           =   4215
      Begin VB.TextBox txtGapWidth1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   60
         TabIndex        =   75
         Text            =   "255"
         Top             =   1800
         Width           =   990
      End
      Begin VB.TextBox txtGapWidth2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   74
         Text            =   "0"
         Top             =   1800
         Width           =   990
      End
      Begin VB.TextBox txtGapHeight1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2100
         TabIndex        =   73
         Text            =   "0"
         Top             =   1800
         Width           =   990
      End
      Begin VB.TextBox txtGapHeight2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   72
         Text            =   "0"
         Top             =   1800
         Width           =   990
      End
      Begin VB.TextBox txtCalmmWidth1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   60
         TabIndex        =   71
         Text            =   "255"
         Top             =   960
         Width           =   990
      End
      Begin VB.TextBox txtCalmmWidth2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   70
         Text            =   "0"
         Top             =   960
         Width           =   990
      End
      Begin VB.TextBox txtCalmmHeight1 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2100
         TabIndex        =   69
         Text            =   "0"
         Top             =   960
         Width           =   990
      End
      Begin VB.TextBox txtCalmmHeight2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   68
         Text            =   "0"
         Top             =   960
         Width           =   990
      End
      Begin VB.TextBox txtCalibThreshold2 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   62
         Text            =   "0"
         Top             =   270
         Width           =   1530
      End
      Begin BHButton.BHImageButton btnCameraGapRegion 
         Height          =   615
         Left            =   240
         TabIndex        =   63
         Top             =   2430
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         Caption         =   "영역"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton btnMeasureGap 
         Height          =   615
         Left            =   2190
         TabIndex        =   64
         Top             =   2430
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         Caption         =   "측정"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Contrast Threshold"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   9
         Left            =   180
         TabIndex        =   67
         Top             =   360
         Width           =   2130
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "실측값 (mm)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   8
         Left            =   180
         TabIndex        =   66
         Top             =   690
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "측정값 (mm)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   180
         TabIndex        =   65
         Top             =   1515
         Width           =   1380
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000E&
      Caption         =   "조명 밝기 (0~255)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   14730
      TabIndex        =   46
      Top             =   1200
      Width           =   4185
      Begin VB.TextBox txtLightBrightness 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   90
         TabIndex        =   56
         Text            =   "0"
         Top             =   990
         Width           =   1320
      End
      Begin VB.TextBox txtLightBrightness 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   2790
         TabIndex        =   55
         Text            =   "0"
         Top             =   990
         Width           =   1320
      End
      Begin VB.TextBox txtLightBrightness 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1440
         TabIndex        =   54
         Text            =   "0"
         Top             =   990
         Width           =   1320
      End
      Begin VB.Label lblLightSide 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   345
         TabIndex        =   53
         Top             =   495
         Width           =   795
      End
      Begin VB.Label lblLightNSD 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "NSD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2430
         TabIndex        =   52
         Top             =   495
         Width           =   705
      End
      Begin VB.Shape shpLightSide 
         BackColor       =   &H00000000&
         BackStyle       =   1  '투명하지 않음
         Height          =   525
         Left            =   90
         Top             =   420
         Width           =   1305
      End
      Begin VB.Shape shpLightNSD 
         BackColor       =   &H00000000&
         BackStyle       =   1  '투명하지 않음
         Height          =   525
         Left            =   1440
         Top             =   420
         Width           =   2655
      End
   End
   Begin VB.Frame Frame22 
      BackColor       =   &H80000004&
      Caption         =   "켈리브레이션 (mm/1Pixel)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3045
      Left            =   14730
      TabIndex        =   36
      Top             =   2910
      Width           =   4185
      Begin VB.TextBox txtCalmmP 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   43
         Text            =   "0"
         Top             =   1530
         Width           =   1530
      End
      Begin VB.TextBox txtCalmm 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   38
         Text            =   "0"
         Top             =   968
         Width           =   1530
      End
      Begin VB.TextBox txtCalibThreshold 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   37
         Text            =   "0"
         Top             =   420
         Width           =   1530
      End
      Begin BHButton.BHImageButton btnCalibrationRegion 
         Height          =   615
         Left            =   240
         TabIndex        =   39
         Top             =   2220
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         Caption         =   "영역"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton btnDoCalibration 
         Height          =   615
         Left            =   2160
         TabIndex        =   40
         Top             =   2220
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   1085
         Caption         =   "측정"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton btnDoCalibrationY 
         Height          =   615
         Left            =   3150
         TabIndex        =   88
         Top             =   2220
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   1085
         Caption         =   "Y축"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin VB.Label lblCalib 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "측정값 (mm/Pixel)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   180
         TabIndex        =   44
         Top             =   1635
         Width           =   2085
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "실측값 (mm)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   180
         TabIndex        =   42
         Top             =   1080
         Width           =   1380
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Contrast Threshold"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   180
         TabIndex        =   41
         Top             =   510
         Width           =   2130
      End
   End
   Begin BHButton.BHImageButton BHBImageLoad 
      Height          =   420
      Left            =   12960
      TabIndex        =   33
      Top             =   300
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   741
      Caption         =   "이미지 불러오기"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H8000000E&
      Caption         =   "검사결과"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   210
      TabIndex        =   12
      Top             =   13680
      Visible         =   0   'False
      Width           =   14475
      Begin VB.Label lblResultName 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "항목1"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   360
         Index           =   0
         Left            =   90
         TabIndex        =   32
         Top             =   315
         Width           =   1800
      End
      Begin VB.Label lblResultData 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H8000000E&
         BorderStyle     =   1  '단일 고정
         Caption         =   "00.00"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   90
         TabIndex        =   31
         Top             =   660
         Width           =   1800
      End
      Begin VB.Label lblResultName 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "항목1"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   360
         Index           =   1
         Left            =   1875
         TabIndex        =   30
         Top             =   315
         Width           =   1800
      End
      Begin VB.Label lblResultData 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H8000000E&
         BorderStyle     =   1  '단일 고정
         Caption         =   "00.00"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1875
         TabIndex        =   29
         Top             =   660
         Width           =   1800
      End
      Begin VB.Label lblResultName 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "항목1"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   360
         Index           =   2
         Left            =   3660
         TabIndex        =   28
         Top             =   315
         Width           =   1800
      End
      Begin VB.Label lblResultData 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H8000000E&
         BorderStyle     =   1  '단일 고정
         Caption         =   "00.00"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3660
         TabIndex        =   27
         Top             =   660
         Width           =   1800
      End
      Begin VB.Label lblResultName 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "항목1"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   360
         Index           =   3
         Left            =   5445
         TabIndex        =   26
         Top             =   315
         Width           =   1800
      End
      Begin VB.Label lblResultData 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H8000000E&
         BorderStyle     =   1  '단일 고정
         Caption         =   "00.00"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   5445
         TabIndex        =   25
         Top             =   660
         Width           =   1800
      End
      Begin VB.Label lblResultName 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "항목1"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   360
         Index           =   4
         Left            =   7230
         TabIndex        =   24
         Top             =   315
         Width           =   1800
      End
      Begin VB.Label lblResultData 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H8000000E&
         BorderStyle     =   1  '단일 고정
         Caption         =   "00.00"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   7230
         TabIndex        =   23
         Top             =   660
         Width           =   1800
      End
      Begin VB.Label lblResultName 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "항목1"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   360
         Index           =   5
         Left            =   9015
         TabIndex        =   22
         Top             =   315
         Width           =   1800
      End
      Begin VB.Label lblResultData 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H8000000E&
         BorderStyle     =   1  '단일 고정
         Caption         =   "00.00"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   9015
         TabIndex        =   21
         Top             =   660
         Width           =   1800
      End
      Begin VB.Label lblResultName 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "항목1"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   360
         Index           =   6
         Left            =   10800
         TabIndex        =   20
         Top             =   315
         Width           =   1800
      End
      Begin VB.Label lblResultData 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H8000000E&
         BorderStyle     =   1  '단일 고정
         Caption         =   "00.00"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   10800
         TabIndex        =   19
         Top             =   660
         Width           =   1800
      End
      Begin VB.Label lblResultName 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  '단일 고정
         Caption         =   "항목1"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   360
         Index           =   7
         Left            =   12585
         TabIndex        =   18
         Top             =   315
         Width           =   1800
      End
      Begin VB.Label lblResultData 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H8000000E&
         BorderStyle     =   1  '단일 고정
         Caption         =   "00.00"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   12585
         TabIndex        =   17
         Top             =   660
         Width           =   1800
      End
      Begin VB.Label lblResultName 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         BorderStyle     =   1  '단일 고정
         Caption         =   "항목1"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   360
         Index           =   8
         Left            =   14460
         TabIndex        =   16
         Top             =   315
         Width           =   1800
      End
      Begin VB.Label lblResultData 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H8000000E&
         BorderStyle     =   1  '단일 고정
         Caption         =   "00.00"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   14460
         TabIndex        =   15
         Top             =   660
         Width           =   1800
      End
      Begin VB.Label lblResultName 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         BorderStyle     =   1  '단일 고정
         Caption         =   "항목1"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   360
         Index           =   9
         Left            =   16245
         TabIndex        =   14
         Top             =   315
         Width           =   1800
      End
      Begin VB.Label lblResultData 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H8000000E&
         BorderStyle     =   1  '단일 고정
         Caption         =   "00.00"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   16245
         TabIndex        =   13
         Top             =   660
         Width           =   1800
      End
   End
   Begin BHButton.BHImageButton BHBToolSave 
      Height          =   690
      Left            =   11055
      TabIndex        =   10
      Top             =   11625
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   1217
      Caption         =   "저 장"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton BHBAcquireS 
      Height          =   690
      Left            =   3855
      TabIndex        =   4
      Top             =   11625
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   1217
      Caption         =   "1회 촬영"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AlphaColor      =   -2147483630
      ImgOutLineSize  =   3
   End
   Begin MSComDlg.CommonDialog ComDialogS 
      Left            =   12450
      Top             =   270
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraCamera 
      BackColor       =   &H8000000E&
      Caption         =   "카메라 선택"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   14700
      TabIndex        =   1
      Top             =   240
      Width           =   4260
      Begin VB.OptionButton optSelectCam 
         BackColor       =   &H00FFFFFF&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   3
         Left            =   3390
         TabIndex        =   35
         Top             =   330
         Width           =   555
      End
      Begin VB.OptionButton optSelectCam 
         BackColor       =   &H00FFFFFF&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   2
         Left            =   2430
         TabIndex        =   34
         Top             =   330
         Width           =   555
      End
      Begin VB.OptionButton optSelectCam 
         BackColor       =   &H00FFFFFF&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   1
         Left            =   1290
         TabIndex        =   3
         Top             =   330
         Width           =   555
      End
      Begin VB.OptionButton optSelectCam 
         BackColor       =   &H00FFFFFF&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   0
         Left            =   210
         TabIndex        =   2
         Top             =   330
         Value           =   -1  'True
         Width           =   555
      End
   End
   Begin BHButton.BHImageButton BHBMasterSave 
      Height          =   690
      Left            =   5655
      TabIndex        =   5
      Top             =   11625
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   1217
      Caption         =   "마스터 저장"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AlphaColor      =   -2147483630
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton BHBMasterLoad 
      Height          =   690
      Left            =   7455
      TabIndex        =   6
      Top             =   11625
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   1217
      Caption         =   "마스터 로드"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AlphaColor      =   -2147483630
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton BHBManualRun 
      Height          =   690
      Left            =   9255
      TabIndex        =   7
      Top             =   11625
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   1217
      Caption         =   "수동 검사"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AlphaColor      =   -2147483630
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton BHBLiveS 
      Height          =   690
      Left            =   255
      TabIndex        =   8
      Top             =   11625
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   1217
      Caption         =   "연속 촬영"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AlphaColor      =   -2147483630
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton BHBLiveStop 
      Height          =   690
      Left            =   2055
      TabIndex        =   9
      Top             =   11625
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   1217
      Caption         =   "정지"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AlphaColor      =   -2147483630
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton BHBSetupEnd 
      Height          =   690
      Left            =   12855
      TabIndex        =   11
      Top             =   11625
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   1217
      Caption         =   "닫 기"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin CognexDisplay.CogDisplay CogDisplayFull 
      Height          =   5355
      Index           =   0
      Left            =   240
      TabIndex        =   57
      Top             =   810
      Width           =   7185
      _cx             =   12674
      _cy             =   9446
      BackColor       =   6556170
      HorizontalScrollBar=   0   'False
      VerticalScrollBar=   0   'False
      PopupMenu       =   -1  'True
      MouseMode       =   0
      ScalingMethod   =   1
      PanX            =   0
      PanY            =   0
      Zoom            =   1
      GridColor       =   12582912
      GridEnabled     =   -1  'True
      SubPixelGridEnabled=   -1  'True
      SubPixelGridColor=   32768
      TouchDistance   =   3
      InteractiveGraphicTipsEnabled=   -1  'True
      MultiSelectionEnabled=   0   'False
      _ipam           =   1
      _spam           =   1
      Enabled         =   -1  'True
      AutoFit         =   -1  'True
   End
   Begin CognexDisplay.CogDisplay CogDisplayFull 
      Height          =   5355
      Index           =   1
      Left            =   7440
      TabIndex        =   58
      Top             =   810
      Width           =   7185
      _cx             =   12674
      _cy             =   9446
      BackColor       =   6556170
      HorizontalScrollBar=   0   'False
      VerticalScrollBar=   0   'False
      PopupMenu       =   -1  'True
      MouseMode       =   0
      ScalingMethod   =   1
      PanX            =   0
      PanY            =   0
      Zoom            =   1
      GridColor       =   12582912
      GridEnabled     =   -1  'True
      SubPixelGridEnabled=   -1  'True
      SubPixelGridColor=   32768
      TouchDistance   =   3
      InteractiveGraphicTipsEnabled=   -1  'True
      MultiSelectionEnabled=   0   'False
      _ipam           =   1
      _spam           =   1
      Enabled         =   -1  'True
      AutoFit         =   -1  'True
   End
   Begin CognexDisplay.CogDisplay CogDisplayFull 
      Height          =   5355
      Index           =   2
      Left            =   240
      TabIndex        =   59
      Top             =   6180
      Width           =   7185
      _cx             =   12674
      _cy             =   9446
      BackColor       =   6556170
      HorizontalScrollBar=   0   'False
      VerticalScrollBar=   0   'False
      PopupMenu       =   -1  'True
      MouseMode       =   0
      ScalingMethod   =   1
      PanX            =   0
      PanY            =   0
      Zoom            =   1
      GridColor       =   12582912
      GridEnabled     =   -1  'True
      SubPixelGridEnabled=   -1  'True
      SubPixelGridColor=   32768
      TouchDistance   =   3
      InteractiveGraphicTipsEnabled=   -1  'True
      MultiSelectionEnabled=   0   'False
      _ipam           =   1
      _spam           =   1
      Enabled         =   -1  'True
      AutoFit         =   -1  'True
   End
   Begin CognexDisplay.CogDisplay CogDisplayFull 
      Height          =   5355
      Index           =   3
      Left            =   7440
      TabIndex        =   60
      Top             =   6180
      Width           =   7185
      _cx             =   12674
      _cy             =   9446
      BackColor       =   6556170
      HorizontalScrollBar=   0   'False
      VerticalScrollBar=   0   'False
      PopupMenu       =   -1  'True
      MouseMode       =   0
      ScalingMethod   =   1
      PanX            =   0
      PanY            =   0
      Zoom            =   1
      GridColor       =   12582912
      GridEnabled     =   -1  'True
      SubPixelGridEnabled=   -1  'True
      SubPixelGridColor=   32768
      TouchDistance   =   3
      InteractiveGraphicTipsEnabled=   -1  'True
      MultiSelectionEnabled=   0   'False
      _ipam           =   1
      _spam           =   1
      Enabled         =   -1  'True
      AutoFit         =   -1  'True
   End
   Begin BHButton.BHImageButton BHBCopy 
      Height          =   720
      Left            =   6360
      TabIndex        =   108
      Top             =   30
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   1270
      Caption         =   "ROI복사"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin VB.Frame fraBlob 
      BackColor       =   &H80000009&
      Caption         =   "NSD 유무 검사"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   14730
      TabIndex        =   89
      Top             =   9300
      Visible         =   0   'False
      Width           =   4185
      Begin VB.TextBox txtBlobIndex 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2370
         TabIndex        =   92
         Text            =   "0"
         Top             =   360
         Width           =   1530
      End
      Begin VB.TextBox txtBlobThreshold 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2370
         TabIndex        =   91
         Text            =   "0"
         Top             =   990
         Width           =   1530
      End
      Begin VB.TextBox txtBlobMinArea 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2370
         TabIndex        =   90
         Text            =   "0"
         Top             =   1620
         Width           =   1530
      End
      Begin BHButton.BHImageButton btnBlobRegion 
         Height          =   615
         Left            =   240
         TabIndex        =   93
         Top             =   2370
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1085
         Caption         =   "영역"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton btnRunBlob 
         Height          =   615
         Left            =   1500
         TabIndex        =   94
         Top             =   2370
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1085
         Caption         =   "검사"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton btnCloseBlob 
         Height          =   615
         Left            =   2760
         TabIndex        =   95
         Top             =   2370
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1085
         Caption         =   "닫기"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "인덱스"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   12
         Left            =   300
         TabIndex        =   98
         Top             =   480
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "경계값(0~255)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   13
         Left            =   300
         TabIndex        =   97
         Top             =   1110
         Width           =   1710
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "최소 영역"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   14
         Left            =   300
         TabIndex        =   96
         Top             =   1740
         Width           =   1020
      End
   End
   Begin BHButton.BHImageButton btnCaliperMode 
      Height          =   405
      Left            =   18150
      TabIndex        =   87
      Top             =   9300
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   714
      Caption         =   "NSD"
      CaptionChecked  =   "Caliper"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckOption     =   1
      ImgOutLineSize  =   3
   End
   Begin VB.Frame fraNSD 
      BackColor       =   &H8000000E&
      Caption         =   "NSD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   14730
      TabIndex        =   76
      Top             =   9300
      Width           =   4185
      Begin VB.ComboBox cboCaliperNSD 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         TabIndex        =   83
         Text            =   "cboCaliper"
         Top             =   270
         Width           =   2370
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H8000000E&
         Caption         =   "Tool1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   80
         Top             =   660
         Width           =   1995
         Begin VB.CheckBox chkNsdPolarity 
            Caption         =   "DarkToLight"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   114
            Top             =   1350
            Width           =   1815
         End
         Begin VB.TextBox txtNsdFilterWidth 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   1080
            TabIndex        =   110
            Text            =   "3"
            Top             =   750
            Width           =   840
         End
         Begin VB.TextBox txtNsdThreshold 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   1080
            TabIndex        =   81
            Text            =   "10"
            ToolTipText     =   "경계값 (0~50)"
            Top             =   210
            Width           =   840
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "필터너비"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   10
            Left            =   60
            TabIndex        =   111
            Top             =   870
            Width           =   960
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "경계값"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   3
            Left            =   60
            TabIndex        =   82
            ToolTipText     =   "경계값 (0~50)"
            Top             =   330
            Width           =   720
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000E&
         Caption         =   "Tool2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   2130
         TabIndex        =   77
         Top             =   660
         Width           =   1995
         Begin VB.CheckBox chkNsdPolarity 
            Caption         =   "DarkToLight"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   90
            TabIndex        =   115
            Top             =   1350
            Width           =   1815
         End
         Begin VB.TextBox txtNsdFilterWidth 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   1080
            TabIndex        =   112
            Text            =   "3"
            Top             =   750
            Width           =   840
         End
         Begin VB.TextBox txtNsdThreshold 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   1080
            TabIndex        =   78
            Text            =   "10"
            ToolTipText     =   "경계값 (0~50)"
            Top             =   210
            Width           =   840
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "필터너비"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   18
            Left            =   60
            TabIndex        =   113
            Top             =   870
            Width           =   960
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "경계값"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   60
            TabIndex        =   79
            ToolTipText     =   "경계값 (0~50)"
            Top             =   330
            Width           =   720
         End
      End
      Begin BHButton.BHImageButton btnNsdRegion 
         Height          =   615
         Left            =   1440
         TabIndex        =   84
         Top             =   2490
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   1085
         Caption         =   "영역"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton btnFindNsd 
         Height          =   615
         Left            =   2790
         TabIndex        =   85
         Top             =   2490
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   1085
         Caption         =   "측정"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton btnBlobExist 
         Height          =   615
         Left            =   90
         TabIndex        =   99
         Top             =   2490
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   1085
         Caption         =   "유무"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사 항목"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   4
         Left            =   240
         TabIndex        =   86
         Top             =   330
         Width           =   1020
      End
   End
   Begin VB.Frame fraCaliper 
      BackColor       =   &H8000000E&
      Caption         =   "Caliper Setting"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   14730
      TabIndex        =   47
      Top             =   9300
      Width           =   4185
      Begin VB.Frame Frame8 
         BackColor       =   &H80000005&
         Caption         =   "Tool1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Left            =   90
         TabIndex        =   122
         Top             =   780
         Width           =   1995
         Begin VB.TextBox txtCaliperThreshold 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   1050
            TabIndex        =   125
            Text            =   "10"
            Top             =   240
            Width           =   840
         End
         Begin VB.TextBox txtCaliperFilterWidth 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   1050
            TabIndex        =   124
            Text            =   "3"
            Top             =   750
            Width           =   840
         End
         Begin VB.CheckBox chkCaliperPolarity 
            Caption         =   "DarkToLight"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   123
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "경계값"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   6
            Left            =   90
            TabIndex        =   127
            Top             =   330
            Width           =   720
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "필터너비"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   15
            Left            =   90
            TabIndex        =   126
            Top             =   840
            Width           =   960
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H80000005&
         Caption         =   "Tool2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Left            =   2130
         TabIndex        =   116
         Top             =   780
         Width           =   1995
         Begin VB.TextBox txtCaliperThreshold 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   1050
            TabIndex        =   119
            Text            =   "10"
            Top             =   240
            Width           =   840
         End
         Begin VB.TextBox txtCaliperFilterWidth 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   1050
            TabIndex        =   118
            Text            =   "3"
            Top             =   750
            Width           =   840
         End
         Begin VB.CheckBox chkCaliperPolarity 
            Caption         =   "DarkToLight"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   117
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "경계값"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   7
            Left            =   90
            TabIndex        =   121
            Top             =   450
            Width           =   720
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "필터너비"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   16
            Left            =   90
            TabIndex        =   120
            Top             =   840
            Width           =   960
         End
      End
      Begin VB.ComboBox cboCaliper 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         TabIndex        =   48
         Text            =   "cboCaliper"
         Top             =   330
         Width           =   2370
      End
      Begin BHButton.BHImageButton btnCaliperRegion 
         Height          =   615
         Left            =   240
         TabIndex        =   50
         Top             =   2490
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         Caption         =   "영역"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton btnFindCaliper 
         Height          =   615
         Left            =   2160
         TabIndex        =   51
         Top             =   2490
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         Caption         =   "측정"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "검사 항목"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   5
         Left            =   240
         TabIndex        =   49
         Top             =   390
         Width           =   1020
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   20
      Height          =   13860
      Left            =   45
      Top             =   60
      Width           =   19110
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   -83.868
      X2              =   40592.18
      Y1              =   11520
      Y2              =   11505
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   40885.71
      X2              =   40927.65
      Y1              =   -30
      Y2              =   12285
   End
   Begin VB.Label lblCameraNumber 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "CAMERA 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   435
      Left            =   9435
      TabIndex        =   0
      Top             =   285
      Width           =   1875
   End
   Begin VB.Shape ShpCamBs 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      Height          =   510
      Left            =   225
      Top             =   255
      Width           =   14475
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mROINo As Integer
Dim LightIdx As Integer
Dim bLightBack As Boolean
Dim bLightNSD As Boolean
Dim bLightSide As Boolean

Private Sub BHBAcquireS_Click()
 bLiveOn = False
    '조명켜기
    Call PWM_LightAll(True, 100)
    'Call PWM_Light(0, True, 100)
    
    If (CamIndex < 4) Then
        Dim i As Integer
        
        For i = 0 To 3
            CogDisplayClear CogDisplayFull(i)
            Set g_CogImage(i) = IDS_AcquireCognex(frmMain.uEyeCam1(i), CogDisplayFull(i))
        Next i
    Else
        CogDisplayClear CogDisplay
        Set g_CogImage(CamIdx) = IDS_AcquireCognex(frmMain.uEyeCam1(CamIdx), CogDisplay)
    End If
    
    '조명끄기
    Call PWM_LightAll(False)
    
End Sub

Private Sub BHBCopy_Click()
Dim ROINo As Long


    frmROI.ROI = mROINo
    frmROI.Show 1, Me
    
    ROINo = frmROI.ROI
    
    If ROINo >= 0 Then
        SaveMultiROI sModelName, ROINo
    End If
    
    Unload frmROI

End Sub

Private Sub BHBImageLoad_Click()
On Error GoTo ErrorHandle

Dim i As Integer
Dim TempImageName As String
Dim imgName As String
Dim tempIndex As Integer

    With frmSetting.ComDialogS
    
        .CancelError = False
        .FileName = "00000000"
        .DefaultExt = ".bmp"
        .ShowOpen
        .DialogTitle = "Image파일 로드"
        
    End With
    
    TempImageName = Me.ComDialogS.FileName
    imgName = TempImageName

    If Dir(imgName, vbNormal) = "" Then
        Exit Sub
    End If
    
    Dim title As String
    Dim ext As String
    
    title = Left(imgName, Len(imgName) - 5)
    ext = Right(imgName, 4)
    
    CogDisplayClear CogDisplay
    For i = 0 To 3
        CogDisplayClear CogDisplayFull(i)
        Set g_CogImage(i) = LoadCogImage(title & CStr(i + 1) & ext)
        Set CogDisplayFull(i).Image = g_CogImage(i)
    Next i
    
    Exit Sub
ErrorHandle:
    
    CogDisplayClear CogDisplay
    Set g_CogImage(CamIdx) = Nothing
    Set CogDisplay.Image = g_CogImage(CamIdx)
    
End Sub

Private Sub BHBLiveS_Click()

    bLiveOn = True
    
    Call PWM_LightAll(True)
    
    If (CamIdx < 4) Then
        Dim i As Integer
        
        Do
            DoEvents
            
            For i = 0 To 3
                CogDisplayClear CogDisplayFull(i)
                Set g_CogImage(i) = IDS_AcquireCognex(frmMain.uEyeCam1(i), CogDisplayFull(i))
            Next i
            Sleep 1
        Loop Until bLiveOn = False
    Else
        Do
            DoEvents
            
            CogDisplayClear CogDisplay
            Set g_CogImage(CamIdx) = IDS_AcquireCognex(frmMain.uEyeCam1(CamIdx), CogDisplay)
            Sleep 1
        Loop Until bLiveOn = False
    End If
        
    Call PWM_LightAll(False)
    
    
End Sub

Private Sub BHBLiveStop_Click()

    bLiveOn = False

End Sub

Private Sub BHBManualRun_Click()
bSettingManualRun = True

    Call PreWelding_RunWidthHeight(CogDisplayFull(0), CogDisplayFull(1), CogDisplayFull(2), CogDisplayFull(3))
    Call PreWelding_RunNSD(CogDisplayFull(0), CogDisplayFull(1), CogDisplayFull(2), CogDisplayFull(3), CogDisplayFull(g_CogBlobIndex))
    
bSettingManualRun = False
End Sub

Private Sub BHBMasterLoad_Click()

    If (CamIdx < 4) Then
        Dim i As Integer
        
        For i = 0 To 3
            Set g_CogImage(i) = LoadCogImage(App.Path & "\Model\" & sModelName & "\" & "Master" & CStr(i) & ".bmp")
    
            CogDisplayClear CogDisplayFull(i)
            Set CogDisplayFull(i).Image = g_CogImage(i)
            
            CogDisplayLabel CogDisplayFull(i), 200, 200, "Cam" & CStr(i + 1), cogColorGreen, "Tahoma", 16
        Next i
    Else
        Set g_CogImage(CamIdx) = LoadCogImage(App.Path & "\Model\" & sModelName & "\" & "Master" & CamIdx & ".bmp")
    
        CogDisplayClear CogDisplay
        Set CogDisplay.Image = g_CogImage(CamIdx)
    End If
    
End Sub

Private Sub BHBMasterSave_Click()

    If MsgBox("마스터 이미지로 저장 하시겠습니까?", vbOKCancel, "마스터 저장") = vbOK Then
        Dim i As Integer
        
        For i = 0 To 3
            SaveCogImage App.Path & "\Model\" & sModelName & "\" & "Master" & i & ".bmp", g_CogImage(i)
        Next i
        
    End If
    
End Sub


Private Sub BHBSetupEnd_Click()

    Unload Me
    
End Sub

Private Sub BHBToolSave_Click()
Dim tempROI As Integer

    If MsgBox("설정값을 저장 하시겠습니까?", vbOKCancel, "저장") = vbOK Then
        g_CameraGrap(0) = CDbl(txtGapWidth1.Text)
        g_CameraGrap(1) = CDbl(txtGapWidth2.Text)
        g_CameraGrap(2) = CDbl(txtGapHeight1.Text)
        g_CameraGrap(3) = CDbl(txtGapHeight2.Text)
        
        If mROINo = 0 Then
        
            Call SaveCogTool(sModelName)
            Call SaveSystemData
            Call Calibration_Save(sModelName)
            Call SaveCameraPosition(sModelName)
            Call FunctionValue_Save(sModelName)
        Else
            tempROI = mROINo
            Call SaveMultiROI(sModelName, mROINo)
            Call LoadMultiROI(sModelName, 0)
            Call SaveCogTool(sModelName)
            Call LoadMultiROI(sModelName, tempROI)
        End If
    End If
    
End Sub



Private Sub btnBlobExist_Click()

    fraBlob.Visible = True
    
End Sub

Private Sub btnBlobRegion_Click()

    Dim i As Integer
    
    For i = 0 To 3
        CogDisplayClear CogDisplayFull(i)
    Next i
    CogDisplayFull(g_CogBlobIndex).InteractiveGraphics.Add g_CogBlobRegion
    
End Sub

Private Sub btnCalibrationRegion_Click()

    Dim i As Integer
    
    CogDisplayClear CogDisplayFull(CamIdx)
    
    For i = 0 To 3
        CogDisplayFull(CamIdx).InteractiveGraphics.Add g_CogCalibrationRegion(i)
    Next i
    
End Sub

Private Sub btnCaliperMode_Click()
    
    If btnCaliperMode.Value = True Then
        fraNSD.Visible = True
        fraCaliper.Visible = False
    Else
        fraNSD.Visible = False
        fraCaliper.Visible = True
    End If

End Sub

Private Sub btnCaliperRegion_Click()

    Dim ToolIdx As Integer
    Dim i As Integer
    
    For i = 0 To 3
        CogDisplayClear CogDisplayFull(i)
    Next i
    
    Select Case cboCaliper.listIndex
    Case 0  '너비1
        CogDisplayFull(0).InteractiveGraphics.Add g_CogCaliperRegion(0)
        CogDisplayFull(1).InteractiveGraphics.Add g_CogCaliperRegion(1)
    Case 1  '너비2
        CogDisplayFull(2).InteractiveGraphics.Add g_CogCaliperRegion(2)
        CogDisplayFull(3).InteractiveGraphics.Add g_CogCaliperRegion(3)
    Case 2  '높이1
        CogDisplayFull(0).InteractiveGraphics.Add g_CogCaliperRegion(4)
        CogDisplayFull(2).InteractiveGraphics.Add g_CogCaliperRegion(5)
    Case 3  '높이2
        CogDisplayFull(1).InteractiveGraphics.Add g_CogCaliperRegion(6)
        CogDisplayFull(3).InteractiveGraphics.Add g_CogCaliperRegion(7)
    End Select
    
End Sub

Private Sub btnCameraGapRegion_Click()
    
    Dim i As Integer
    
    For i = 0 To 3
        CogDisplayClear CogDisplayFull(i)
    Next i
    
    
    For i = 0 To 7
        g_CogGapRegion(i).CenterX = g_CogCaliperRegion(i).CenterX
        g_CogGapRegion(i).CenterY = g_CogCaliperRegion(i).CenterY
    Next i
    
    CogDisplayFull(0).InteractiveGraphics.Add g_CogGapRegion(0)
    CogDisplayFull(0).InteractiveGraphics.Add g_CogGapRegion(4)
    
    CogDisplayFull(1).InteractiveGraphics.Add g_CogGapRegion(1)
    CogDisplayFull(1).InteractiveGraphics.Add g_CogGapRegion(6)
    
    CogDisplayFull(2).InteractiveGraphics.Add g_CogGapRegion(2)
    CogDisplayFull(2).InteractiveGraphics.Add g_CogGapRegion(5)
    
    CogDisplayFull(3).InteractiveGraphics.Add g_CogGapRegion(3)
    CogDisplayFull(3).InteractiveGraphics.Add g_CogGapRegion(7)
    
End Sub

Private Sub btnCloseBlob_Click()

    fraBlob.Visible = False
    
End Sub

Private Sub btnDoCalibration_Click()

    Dim Distance As Double
    Dim i As Integer
    
    For i = 0 To 1
        g_CogCalibrationTool(i).RunParams.ContrastThreshold = CDbl(txtCalibThreshold.Text)
        Set g_CogCalibrationTool(i).InputImage = g_CogImage(CamIdx)
    Next i
    
    CogDisplayClear CogDisplay
    Distance = CogFindCaliperX(g_CogCalibrationTool(0), g_CogCalibrationTool(1), CogDisplayFull(CamIdx))
    
    If Distance < 0 Then
        CogDisplayLabel CogDisplay, 200, 100, "Not found.", cogColorRed, "Tahoma", 16
    End If
    
    dCaliMM(CamIdx) = CDbl(txtCalmm.Text)
    dCaliPX(CamIdx) = dCaliMM(CamIdx) / Distance
    txtCalmmP.Text = Format(dCaliPX(CamIdx), "#0.0000")

End Sub

Private Sub btnDoCalibrationY_Click()
On Error Resume Next
    Dim Distance As Double
    Dim i As Integer
    
    For i = 2 To 3
        g_CogCalibrationTool(i).RunParams.ContrastThreshold = CDbl(txtCalibThreshold.Text)
        Set g_CogCalibrationTool(i).InputImage = g_CogImage(CamIdx)
    Next i
    
    CogDisplayClear CogDisplay
    Distance = CogFindCaliperY(g_CogCalibrationTool(2), g_CogCalibrationTool(3), CogDisplayFull(CamIdx))
    
    If Distance < 0 Then
        CogDisplayLabel CogDisplay, 200, 100, "Not found.", cogColorRed, "Tahoma", 16
    End If
    
    dCaliPXY(CamIdx) = CDbl(InputBox("Y축 실측값", "입력")) / Distance
    MsgBox Format(dCaliPXY(CamIdx), "#0.0000")

End Sub

Private Sub btnFindCaliper_Click()

    Dim ResultPoint1 As Double
    Dim ResultPoint2 As Double
    Dim Distance As Double
    
    Dim Distance1 As Double
    Dim Distance2 As Double
    
    Dim ToolIdx As Integer
    Dim i As Integer
    
    For i = 0 To 3
        CogDisplayClear CogDisplayFull(i)
    Next i
    
    For i = 0 To 1
        ToolIdx = cboCaliper.listIndex * 2 + i
        g_CogCaliperTool(ToolIdx).RunParams.ContrastThreshold = CDbl(txtCaliperThreshold(i).Text)
        g_CogCaliperTool(ToolIdx).RunParams.FilterHalfSizeInPixels = CLng(txtCaliperFilterWidth(i).Text)
        
        If chkCaliperPolarity(i).Value = 1 Then
            g_CogCaliperTool(ToolIdx).RunParams.Edge0Polarity = cogCaliperPolarityDarkToLight
        Else
            g_CogCaliperTool(ToolIdx).RunParams.Edge0Polarity = cogCaliperPolarityLightToDark
        End If
    Next i
    
    Select Case cboCaliper.listIndex
    Case 0  '너비1
        Dim Index As Integer
        
        Set g_CogCaliperTool(0).InputImage = g_CogImage(0)
        Set g_CogCaliperTool(1).InputImage = g_CogImage(1)
        
        Index = 0
        g_CogCaliperTool(Index).Run
        If g_CogCaliperTool(Index).Results.Count <= 0 Then
            Exit Sub
        End If
        CogDisplayFull(0).StaticGraphics.Add g_CogCaliperTool(Index).Results.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
        ResultPoint1 = g_CogCaliperTool(Index).Results.Item(0).PositionX
        Distance1 = ((XRES / 2) - ResultPoint1) * dCaliPX(0)
        
        Index = 1
        g_CogCaliperTool(Index).Run
        If g_CogCaliperTool(Index).Results.Count <= 0 Then
            Exit Sub
        End If
        CogDisplayFull(1).StaticGraphics.Add g_CogCaliperTool(Index).Results.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
        ResultPoint2 = g_CogCaliperTool(Index).Results.Item(0).PositionX
        Distance2 = (ResultPoint2 - (XRES / 2)) * dCaliPX(0)
        
        Distance = g_CameraGrap(0) + (Distance1 + Distance2)
        
        CogDisplayLabel CogDisplayFull(0), 200, 200, "너비1 = " & Format(Distance, "#0.00"), cogColorGreen, "Tahoma", 16, True
        
    Case 1  '너비2
        Set g_CogCaliperTool(2).InputImage = g_CogImage(2)
        Set g_CogCaliperTool(3).InputImage = g_CogImage(3)
        
        Index = 2
        g_CogCaliperTool(Index).Run
        If g_CogCaliperTool(Index).Results.Count <= 0 Then
            Exit Sub
        End If
        CogDisplayFull(2).StaticGraphics.Add g_CogCaliperTool(Index).Results.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
        ResultPoint1 = g_CogCaliperTool(Index).Results.Item(0).PositionX
        Distance1 = ((XRES / 2) - ResultPoint1) * dCaliPX(2)
        
        Index = 3
        g_CogCaliperTool(Index).Run
        If g_CogCaliperTool(Index).Results.Count <= 0 Then
            Exit Sub
        End If
        CogDisplayFull(3).StaticGraphics.Add g_CogCaliperTool(Index).Results.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
        ResultPoint2 = g_CogCaliperTool(Index).Results.Item(0).PositionX
        Distance2 = (ResultPoint2 - (XRES / 2)) * dCaliPX(3)
        
        Distance = g_CameraGrap(1) + (Distance1 + Distance2)
        
        CogDisplayLabel CogDisplayFull(2), 200, 200, "너비2 = " & Format(Distance, "#0.00"), cogColorGreen, "Tahoma", 16, True
    Case 2  '높이1
        Set g_CogCaliperTool(4).InputImage = g_CogImage(0)
        Set g_CogCaliperTool(5).InputImage = g_CogImage(2)
        
        Index = 4
        g_CogCaliperTool(Index).Run
        If g_CogCaliperTool(Index).Results.Count <= 0 Then
            Exit Sub
        End If
        CogDisplayFull(0).StaticGraphics.Add g_CogCaliperTool(Index).Results.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
        ResultPoint1 = g_CogCaliperTool(Index).Results.Item(0).PositionY
        Distance1 = ((YRES / 2) - ResultPoint1) * dCaliPXY(0)
        
        Index = 5
        g_CogCaliperTool(Index).Run
        If g_CogCaliperTool(Index).Results.Count <= 0 Then
            Exit Sub
        End If
        CogDisplayFull(2).StaticGraphics.Add g_CogCaliperTool(Index).Results.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
        ResultPoint2 = g_CogCaliperTool(Index).Results.Item(0).PositionY
        Distance2 = (ResultPoint2 - (YRES / 2)) * dCaliPXY(2)
        
        Distance = g_CameraGrap(2) + (Distance1 + Distance2)
        
        CogDisplayLabel CogDisplayFull(0), 200, 200, "높이1 = " & Format(Distance, "#0.00"), cogColorGreen, "Tahoma", 16, True
    Case 3  '높이2
        Set g_CogCaliperTool(6).InputImage = g_CogImage(1)
        Set g_CogCaliperTool(7).InputImage = g_CogImage(3)
        
        Index = 6
        g_CogCaliperTool(Index).Run
        If g_CogCaliperTool(Index).Results.Count <= 0 Then
            Exit Sub
        End If
        CogDisplayFull(1).StaticGraphics.Add g_CogCaliperTool(Index).Results.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
        ResultPoint1 = g_CogCaliperTool(Index).Results.Item(0).PositionY
        Distance1 = ((YRES / 2) - ResultPoint1) * dCaliPXY(1)
        
        Index = 7
        g_CogCaliperTool(Index).Run
        If g_CogCaliperTool(Index).Results.Count <= 0 Then
            Exit Sub
        End If
        CogDisplayFull(3).StaticGraphics.Add g_CogCaliperTool(Index).Results.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
        ResultPoint2 = g_CogCaliperTool(Index).Results.Item(0).PositionY
        Distance2 = (ResultPoint2 - (YRES / 2)) * dCaliPXY(3)
        
        Distance = g_CameraGrap(3) + (Distance1 + Distance2)
        
        CogDisplayLabel CogDisplayFull(1), 200, 200, "높이2 = " & Format(Distance, "#0.00"), cogColorGreen, "Tahoma", 16, True
    End Select
    
End Sub

Private Sub btnFindNsd_Click()

    Dim ResultPoint1 As Double
    Dim ResultPoint2 As Double
    Dim Distance As Double
    
    Dim Distance1 As Double
    Dim Distance2 As Double
    
    Dim Index As Integer
    Dim listIndex As Integer
    Dim ToolIdx As Integer
    Dim i As Integer
    
    For i = 0 To 3
        CogDisplayClear CogDisplayFull(i)
    Next i
    
    Select Case cboCaliperNSD.listIndex
    Case 0
        If g_NsdRegionSelection = 1 Then
            Index = 1
        Else
            Index = 0
        End If
        listIndex = 0
    Case 1
        If g_NsdRegionSelection = 1 Then
            Index = 3
        Else
            Index = 2
        End If
        listIndex = 2
    Case 2
        Index = 1
        listIndex = 4
    Case 3
        Index = 3
        listIndex = 5
    End Select
    
    ToolIdx = listIndex * 2
    For i = 0 To 1
        g_CogNsdTool(ToolIdx + i).RunParams.ContrastThreshold = CDbl(txtNsdThreshold(i).Text)
        g_CogNsdTool(ToolIdx + i).RunParams.FilterHalfSizeInPixels = CLng(txtNsdFilterWidth(i).Text)
        
        If chkNsdPolarity(i).Value = 1 Then
            g_CogNsdTool(ToolIdx + i).RunParams.Edge0Polarity = cogCaliperPolarityDarkToLight
        Else
            g_CogNsdTool(ToolIdx + i).RunParams.Edge0Polarity = cogCaliperPolarityLightToDark
        End If
        
        Set g_CogNsdTool(ToolIdx + i).InputImage = g_CogImage(Index)
    Next i
    
    If listIndex < 4 Then
        Distance = CogFindCaliperY(g_CogNsdTool(ToolIdx + 0), g_CogNsdTool(ToolIdx + 1), CogDisplayFull(Index), dCaliPXY(Index))
    Else
        Distance = CogFindCaliperX(g_CogNsdTool(ToolIdx + 0), g_CogNsdTool(ToolIdx + 1), CogDisplayFull(Index), dCaliPX(Index))
    End If

End Sub

Private Sub btnMeasureGap_Click()

    Dim CaliperTool As New CogCaliper
    Dim CaliperResults As CogCaliperResults
    Dim dResult(0 To 7) As Double
    
    Dim i As Integer
    
    '캘리퍼 툴 초기화
    With CaliperTool
        .EdgeMode = cogCaliperEdgeModeSingle
        .Edge0Polarity = cogCaliperPolarityLightToDark
        .ContrastThreshold = CDbl(txtCalibThreshold2.Text)
        .FilterHalfSizeInPixels = 3
        .MaxResults = 1
'        .SingleEdgeScorers.Clear
'        .SingleEdgeScorers.Add New CogCaliperScorerPosition
    End With
    
    For i = 0 To 3
        CogDisplayClear CogDisplayFull(i)
    Next i
    '측정
    Set CaliperResults = CaliperTool.Execute(g_CogImage(0), g_CogGapRegion(0))
    If CaliperResults.Count <= 0 Then
        Exit Sub
    End If
    CogDisplayFull(0).StaticGraphics.Add CaliperResults.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
    dResult(0) = CaliperResults.Item(0).PositionX
    
    Set CaliperResults = CaliperTool.Execute(g_CogImage(1), g_CogGapRegion(1))
    If CaliperResults.Count <= 0 Then
        Exit Sub
    End If
    CogDisplayFull(1).StaticGraphics.Add CaliperResults.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
    dResult(1) = CaliperResults.Item(0).PositionX
    
    Set CaliperResults = CaliperTool.Execute(g_CogImage(2), g_CogGapRegion(2))
    If CaliperResults.Count <= 0 Then
        Exit Sub
    End If
    CogDisplayFull(2).StaticGraphics.Add CaliperResults.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
    dResult(2) = CaliperResults.Item(0).PositionX
    
    Set CaliperResults = CaliperTool.Execute(g_CogImage(3), g_CogGapRegion(3))
    If CaliperResults.Count <= 0 Then
        Exit Sub
    End If
    CogDisplayFull(3).StaticGraphics.Add CaliperResults.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
    dResult(3) = CaliperResults.Item(0).PositionX
    
    Set CaliperResults = CaliperTool.Execute(g_CogImage(0), g_CogGapRegion(4))
    If CaliperResults.Count <= 0 Then
        Exit Sub
    End If
    CogDisplayFull(0).StaticGraphics.Add CaliperResults.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
    dResult(4) = CaliperResults.Item(0).PositionY
    
    Set CaliperResults = CaliperTool.Execute(g_CogImage(2), g_CogGapRegion(5))
    If CaliperResults.Count <= 0 Then
        Exit Sub
    End If
    CogDisplayFull(2).StaticGraphics.Add CaliperResults.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
    dResult(5) = CaliperResults.Item(0).PositionY
    
    Set CaliperResults = CaliperTool.Execute(g_CogImage(1), g_CogGapRegion(6))
    If CaliperResults.Count <= 0 Then
        Exit Sub
    End If
    CogDisplayFull(1).StaticGraphics.Add CaliperResults.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
    dResult(6) = CaliperResults.Item(0).PositionY
    
    Set CaliperResults = CaliperTool.Execute(g_CogImage(3), g_CogGapRegion(7))
    If CaliperResults.Count <= 0 Then
        Exit Sub
    End If
    CogDisplayFull(3).StaticGraphics.Add CaliperResults.Item(0).CreateResultGraphics(cogCaliperResultGraphicEdges)
    dResult(7) = CaliperResults.Item(0).PositionY
    
    Dim Width1 As Double
    Dim Width2 As Double
    Dim Height1 As Double
    Dim Height2 As Double
    
    Width1 = CDbl(txtCalmmWidth1.Text)
    Width2 = CDbl(txtCalmmWidth2.Text)
    Height1 = CDbl(txtCalmmHeight1.Text)
    Height2 = CDbl(txtCalmmHeight2.Text)
    
    Dim dWidth1 As Double
    Dim dWidth2 As Double
    
    '너비1
    dWidth1 = ((XRES / 2) - dResult(0)) * dCaliPX(0)
    dWidth2 = (dResult(1) - (XRES / 2)) * dCaliPX(1)
    Width1 = Width1 - (dWidth1 + dWidth2)
    g_CameraGrap(0) = Width1
    g_ProductPt(0) = CDbl(txtCalmmWidth1.Text)
    txtGapWidth1.Text = Format(Width1, "#0.00")
    
    '너비2
    dWidth1 = ((XRES / 2) - dResult(2)) * dCaliPX(2)
    dWidth2 = (dResult(3) - (XRES / 2)) * dCaliPX(3)
    Width2 = Width2 - (dWidth1 + dWidth2)
    g_CameraGrap(1) = Width2
    g_ProductPt(1) = CDbl(txtCalmmWidth2.Text)
    txtGapWidth2.Text = Format(Width2, "#0.00")
    
    '높이1
    dHeight1 = ((YRES / 2) - dResult(4)) * dCaliPXY(0)
    dHeight2 = (dResult(5) - (YRES / 2)) * dCaliPXY(2)
    Height1 = Height1 - (dHeight1 + dHeight2)
    g_CameraGrap(2) = Height1
    g_ProductPt(2) = CDbl(txtCalmmHeight1.Text)
    txtGapHeight1.Text = Format(Height1, "#0.00")
    
    '높이2
    dHeight1 = ((YRES / 2) - dResult(6)) * dCaliPXY(1)
    dHeight2 = (dResult(7) - (YRES / 2)) * dCaliPXY(3)
    Height2 = Height2 - (dHeight1 + dHeight2)
    g_CameraGrap(3) = Height2
    g_ProductPt(3) = CDbl(txtCalmmHeight2.Text)
    txtGapHeight2.Text = Format(Height2, "#0.00")
    
End Sub

Private Sub btnNsdRegion_Click()

    Dim ToolIdx As Integer
    Dim i As Integer
    
    Dim Index As Integer
    Dim listIndex As Integer
        
    For i = 0 To 3
        CogDisplayClear CogDisplayFull(i)
    Next i
    
    Select Case cboCaliperNSD.listIndex
    Case 0
        If g_NsdRegionSelection = 1 Then
            Index = 1
        Else
            Index = 0
        End If
        listIndex = 0
    Case 1
        If g_NsdRegionSelection = 1 Then
            Index = 3
        Else
            Index = 2
        End If
        listIndex = 2
    Case 2
        Index = 1
        listIndex = 4
    Case 3
        Index = 3
        listIndex = 5
    End Select
    
    For i = 0 To 1
        CogDisplayFull(Index).InteractiveGraphics.Add g_CogNsdRegion(listIndex * 2 + i)
        CogDisplayLabel CogDisplayFull(Index), g_CogNsdRegion(listIndex * 2 + i).CenterX, g_CogNsdRegion(listIndex * 2 + i).CenterY, "[" & CStr(i + 1) & "]", cogColorGreen, "Tahoma", 16
    Next i
    
End Sub

Private Sub btnRunBlob_Click()

    CogDisplayClear CogDisplayFull(g_CogBlobIndex)
    Call PreWelding_RunBlob(CogDisplayFull(g_CogBlobIndex))

End Sub

Private Sub cboCaliper_Click()
    Dim i As Integer
    
    For i = 0 To 3
        CogDisplayClear CogDisplayFull(i)
    Next i
    
    txtCaliperThreshold(0).Text = CStr(g_CogCaliperTool(cboCaliper.listIndex * 2 + 0).RunParams.ContrastThreshold)
    txtCaliperThreshold(1).Text = CStr(g_CogCaliperTool(cboCaliper.listIndex * 2 + 1).RunParams.ContrastThreshold)
    
    txtCaliperFilterWidth(0).Text = CStr(g_CogCaliperTool(cboCaliper.listIndex * 2 + 0).RunParams.FilterHalfSizeInPixels)
    txtCaliperFilterWidth(1).Text = CStr(g_CogCaliperTool(cboCaliper.listIndex * 2 + 1).RunParams.FilterHalfSizeInPixels)
    
    If g_CogCaliperTool(cboCaliper.listIndex * 2 + 0).RunParams.Edge0Polarity = cogCaliperPolarityDarkToLight Then
        chkCaliperPolarity(0).Value = 1
    Else
        chkCaliperPolarity(0).Value = 0
    End If
    
    If g_CogCaliperTool(cboCaliper.listIndex * 2 + 1).RunParams.Edge0Polarity = cogCaliperPolarityDarkToLight Then
        chkCaliperPolarity(1).Value = 1
    Else
        chkCaliperPolarity(1).Value = 0
    End If
    
    CogDisplayClear CogDisplay
    
End Sub

Private Sub cboCaliper_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        Call cboCaliper_Click
    End If
End Sub

Private Sub chkNSD_Click()

    If chkNSD.Value = 1 Then
        fraCaliper.Visible = False
        fraNSD.Visible = True
    Else
        fraCaliper.Visible = True
        fraNSD.Visible = False
    End If
    
End Sub

Private Sub cboCaliperNSD_Click()
    
    Dim listIndex As Integer
    Dim i As Integer
    
    For i = 0 To 3
        CogDisplayClear CogDisplayFull(i)
    Next i
    
    Select Case cboCaliperNSD.listIndex
    Case 0
        listIndex = 0
    Case 1
        listIndex = 2
    Case 2
        listIndex = 4
    Case 3
        listIndex = 5
    End Select
    
    txtNsdThreshold(0).Text = CStr(g_CogNsdTool(listIndex * 2 + 0).RunParams.ContrastThreshold)
    txtNsdThreshold(1).Text = CStr(g_CogNsdTool(listIndex * 2 + 1).RunParams.ContrastThreshold)

    txtNsdFilterWidth(0).Text = CStr(g_CogNsdTool(listIndex * 2 + 0).RunParams.FilterHalfSizeInPixels)
    txtNsdFilterWidth(1).Text = CStr(g_CogNsdTool(listIndex * 2 + 1).RunParams.FilterHalfSizeInPixels)
    
    If g_CogNsdTool(listIndex * 2 + 0).RunParams.Edge0Polarity = cogCaliperPolarityDarkToLight Then
        chkNsdPolarity(0).Value = 1
    Else
        chkNsdPolarity(0).Value = 0
    End If
    
    If g_CogNsdTool(listIndex * 2 + 1).RunParams.Edge0Polarity = cogCaliperPolarityDarkToLight Then
        chkNsdPolarity(1).Value = 1
    Else
        chkNsdPolarity(1).Value = 0
    End If
    
End Sub

Private Sub chkCaliperPolarity_Click(Index As Integer)


    Dim ToolIdx As Integer
    
    ToolIdx = cboCaliper.listIndex * 2 + Index
    
    If chkCaliperPolarity(Index).Value = 1 Then
        g_CogCaliperTool(ToolIdx).RunParams.Edge0Polarity = cogCaliperPolarityDarkToLight
    Else
        g_CogCaliperTool(ToolIdx).RunParams.Edge0Polarity = cogCaliperPolarityLightToDark
    End If
    
    
End Sub

Private Sub chkNsdPolarity_Click(Index As Integer)

    
    Dim ToolIdx As Integer
    Dim listIndex As Integer
    
    Select Case cboCaliperNSD.listIndex
    Case 0
        listIndex = 0
    Case 1
        listIndex = 2
    Case 2
        listIndex = 4
    Case 3
        listIndex = 5
    End Select
    
    ToolIdx = listIndex * 2 + Index
    
    If chkNsdPolarity(Index).Value = 1 Then
        g_CogNsdTool(ToolIdx).RunParams.Edge0Polarity = cogCaliperPolarityDarkToLight
    Else
        g_CogNsdTool(ToolIdx).RunParams.Edge0Polarity = cogCaliperPolarityLightToDark
    End If
    
End Sub

Private Sub chkNsdRegionSelection_Click()

    g_NsdRegionSelection = chkNsdRegionSelection.Value
    
End Sub

Private Sub Form_Load()
    
    Dim i As Integer
    
    CamIdx = 0
    LightIdx = 0
    mROINo = 0
    
    '검사결과창 각 항목라벨에 SpecName 쓰기
    For i = 0 To 9
        lblResultName(i).Caption = sSpecName(i)
    Next i
    
    '카메라 간 간격
    txtCalmmWidth1.Text = Format(g_ProductPt(0), "#0.0")
    txtCalmmWidth2.Text = Format(g_ProductPt(1), "#0.0")
    txtCalmmHeight1.Text = Format(g_ProductPt(2), "#0.0")
    txtCalmmHeight2.Text = Format(g_ProductPt(3), "#0.0")
    
    txtGapWidth1.Text = Format(g_CameraGrap(0), "#0.0")
    txtGapWidth2.Text = Format(g_CameraGrap(1), "#0.0")
    txtGapHeight1.Text = Format(g_CameraGrap(2), "#0.0")
    txtGapHeight2.Text = Format(g_CameraGrap(3), "#0.0")
    
    g_CogCaliperTool(4).RunParams.SingleEdgeScorers.Clear
    g_CogCaliperTool(4).RunParams.SingleEdgeScorers.Add g_CogCaliperScorerPosition
    
    g_CogCaliperTool(6).RunParams.SingleEdgeScorers.Clear
    g_CogCaliperTool(6).RunParams.SingleEdgeScorers.Add g_CogCaliperScorerPosition
    
    '조명
    bLightBack = False
    bLightNSD = False
    bLightSide = False
    
    For i = 0 To kMaxLight - 1
        txtLightBrightness(i).Text = CStr(g_btLightBrightness(i))
    Next i
    
    'Caliper 검사 항목 콤보박스 추가
    For i = 0 To 3
        cboCaliper.AddItem sSpecName(i)
    Next i
    cboCaliper.listIndex = 0
    
    For i = 0 To 3
        cboCaliperNSD.AddItem "NSD" & CStr(i + 1)
    Next i
    cboCaliperNSD.listIndex = 0
    fraNSD.Visible = False
    
    'NSD 유무 검사
    txtBlobIndex.Text = CStr(g_CogBlobIndex + 1)
    txtBlobThreshold.Text = CStr(g_CogBlobTool.RunParams.SegmentationParams.HardFixedThreshold)
    txtBlobMinArea.Text = CStr(g_CogBlobTool.RunParams.ConnectivityMinPixels)
    
    chkNsdRegionSelection.Value = g_NsdRegionSelection
    
    Call optSelectCam_Click(0)
    
End Sub


Private Sub lblCalib_DblClick()
    
    If txtCalmmP.Enabled = True Then
        dCaliPX(CamIdx) = CDbl(txtCalmmP.Text)
        txtCalmmP.Enabled = False
    Else
        txtCalmmP.Enabled = True
        txtCalmmP.SetFocus
        txtCalmmP.SelStart = 0
        txtCalmmP.SelLength = Len(txtCalmmP.Text)
    End If
    
End Sub


Private Sub lblLightBack_DblClick()

    Dim i As Integer
    
    bLightBack = Not bLightBack
    lblLightBack.ForeColor = IIf(bLightBack = True, vbGreen, vbWhite)
    
    For i = 0 To 3
        txtLightBrightness(i).BackColor = IIf(bLightBack = True, vbGreen, vbWhite)
        Call PWM_Light(CLng(i), bLightBack)
    Next i
    
End Sub

Private Sub lblLightNSD_DblClick()

    Dim i As Integer
    
    bLightNSD = Not bLightNSD
    lblLightNSD.ForeColor = IIf(bLightNSD = True, vbGreen, vbWhite)
    
    For i = 1 To 2
        txtLightBrightness(i).BackColor = IIf(bLightNSD = True, vbGreen, vbWhite)
        Call PWM_Light(CLng(i), bLightNSD)
    Next i
    
End Sub

Private Sub lblLightSide_DblClick()

    Dim i As Integer
    
    bLightSide = Not bLightSide
    lblLightSide.ForeColor = IIf(bLightSide = True, vbGreen, vbWhite)
    
    For i = 0 To 0
        txtLightBrightness(i).BackColor = IIf(bLightSide = True, vbGreen, vbWhite)
        Call PWM_Light(CLng(i), bLightSide)
    Next i
    
End Sub

Private Sub optROI_Click(Index As Integer)

    mROINo = Index
        
    LoadMultiROI sModelName, mROINo
    
End Sub

Private Sub optSelectCam_Click(Index As Integer)
    
    CamIdx = Index
    
    '카메라넘버 표시
    lblCameraNumber.Caption = "CAMERA" & CStr(CamIdx + 1)
    
    '디스플레이
    For i = 0 To 3
        CogDisplayFull(i).Visible = (Index < 4)
    Next i
    CogDisplay.Visible = (Index >= 4)
    
    '프레임
    fraCaliper.Visible = (Index < 4)
    fraNSD.Visible = False
    
    '마스터 이미지 로드
    Call CogDisplayClear(CogDisplay)
    'Call BHBMasterLoad_Click
    If Index < 4 Then
        For i = 0 To 3
            Set CogDisplayFull(i).Image = g_CogImage(i)
        Next i
    Else
        Set CogDisplay.Image = g_CogImage(Index)
    End If
    
    '캘리브레이션
    txtCalibThreshold.Text = CStr(g_CogCalibrationTool(0).RunParams.ContrastThreshold)
    txtCalmm.Text = CStr(dCaliMM(CamIdx))
    txtCalmmP.Text = Format(dCaliPX(CamIdx), "#0.0000")
    
    'Caliper
    Call cboCaliper_Click
    Call cboCaliperNSD_Click
    
End Sub

Private Sub txtBlobIndex_Change()

    If CheckTextBox(txtBlobIndex, 1, 4) = True Then
        g_CogBlobIndex = CInt(txtBlobIndex.Text) - 1
    End If
    
End Sub

Private Sub txtBlobMinArea_Change()

    If CheckTextBox(txtBlobMinArea, 0, 100000000) = True Then
        g_CogBlobTool.RunParams.ConnectivityMinPixels = CLng(txtBlobMinArea.Text)
    End If
    
End Sub

Private Sub txtBlobThreshold_Change()
    
    If CheckTextBox(txtBlobThreshold, 0, 255) = True Then
        g_CogBlobTool.RunParams.SegmentationParams.HardFixedThreshold = CLng(txtBlobThreshold.Text)
    End If
    
End Sub

Private Sub txtCaliperFilterWidth_Change(Index As Integer)

    
    Dim ToolIdx As Integer
    
    ToolIdx = cboCaliper.listIndex * 2 + Index
    
    If CheckTextBox(txtCaliperFilterWidth(Index), 2, 50) = True Then
        g_CogCaliperTool(ToolIdx).RunParams.FilterHalfSizeInPixels = CLng(txtCaliperFilterWidth(Index).Text)
    End If
    
End Sub

Private Sub txtCaliperThreshold_Change(Index As Integer)
    Dim ToolIdx As Integer
    
    ToolIdx = cboCaliper.listIndex * 2 + Index
    
    If CheckTextBox(txtCaliperThreshold(Index), 0, 50) = True Then
        g_CogCaliperTool(ToolIdx).RunParams.ContrastThreshold = CDbl(txtCaliperThreshold(Index).Text)
    End If
End Sub

Private Sub txtLightBrightness_Change(Index As Integer)

    bLightBack = False
    bLightNSD = False
    bLightSide = False
    
    If CheckTextBox(txtLightBrightness(Index), 0, 255) = True Then
        g_btLightBrightness(Index) = CByte(txtLightBrightness(Index).Text)
        Call PWM_SetLight(CLng(Index), CLng(g_btLightBrightness(Index)))
    End If
    
End Sub

Private Sub txtLightBrightness_DblClick(Index As Integer)

    txtLightBrightness(Index).BackColor = IIf(txtLightBrightness(Index).BackColor = vbWhite, vbGreen, vbWhite)
    If (txtLightBrightness(Index).BackColor = vbGreen) Then
        Call PWM_LightOn(CLng(Index))
    Else
        Call PWM_LightOff(CLng(Index))
    End If
    
End Sub

Private Sub txtLightBrightness_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDown Then
        txtLightBrightness(Index).Text = CStr(g_btLightBrightness(Index) - 5)
    End If
    
    If KeyCode = vbKeyUp Then
        txtLightBrightness(Index).Text = CStr(g_btLightBrightness(Index) + 5)
    End If
End Sub

Private Sub txtNsdFilterWidth_Change(Index As Integer)


    Dim ToolIdx As Integer
    Dim listIndex As Integer
    
    Select Case cboCaliperNSD.listIndex
    Case 0
        listIndex = 0
    Case 1
        listIndex = 2
    Case 2
        listIndex = 4
    Case 3
        listIndex = 5
    End Select
    ToolIdx = listIndex * 2 + Index
    
    If CheckTextBox(txtNsdFilterWidth(Index), 2, 50) = True Then
        g_CogNsdTool(ToolIdx).RunParams.FilterHalfSizeInPixels = CLng(txtNsdFilterWidth(Index).Text)
    End If



End Sub

Private Sub txtNsdThreshold_Change(Index As Integer)


    Dim ToolIdx As Integer
    Dim listIndex As Integer
    
    Select Case cboCaliperNSD.listIndex
    Case 0
        listIndex = 0
    Case 1
        listIndex = 2
    Case 2
        listIndex = 4
    Case 3
        listIndex = 5
    End Select
    
    ToolIdx = listIndex * 2 + Index
    
    If CheckTextBox(txtNsdThreshold(Index), 0, 255) = True Then
        g_CogNsdTool(ToolIdx).RunParams.ContrastThreshold = CDbl(txtNsdThreshold(Index).Text)
    End If
    
End Sub
