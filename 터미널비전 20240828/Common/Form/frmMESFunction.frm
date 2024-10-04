VERSION 5.00
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmMESFunction 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  '없음
   Caption         =   "환경설정"
   ClientHeight    =   6675
   ClientLeft      =   390
   ClientTop       =   1755
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   10740
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Timeout"
      Height          =   615
      Left            =   7410
      TabIndex        =   102
      Top             =   5790
      Width           =   2865
      Begin VB.TextBox txtTimeoutRetry 
         Height          =   345
         Left            =   2130
         TabIndex        =   106
         Text            =   "3"
         Top             =   210
         Width           =   615
      End
      Begin VB.TextBox txtTimeoutInterval 
         Height          =   345
         Left            =   780
         TabIndex        =   104
         Text            =   "1000"
         Top             =   210
         Width           =   795
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Retry"
         Height          =   180
         Left            =   1650
         TabIndex        =   105
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Interval"
         Height          =   180
         Left            =   90
         TabIndex        =   103
         Top             =   300
         Width           =   600
      End
   End
   Begin VB.Frame fraSVCode 
      BackColor       =   &H8000000E&
      Caption         =   "항목코드(SV)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5445
      Left            =   480
      TabIndex        =   24
      Top             =   270
      Width           =   4680
      Begin VB.TextBox txtSpecCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   10
         Left            =   1710
         TabIndex        =   97
         Top             =   4920
         Width           =   2730
      End
      Begin VB.TextBox txtSpecCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   9
         Left            =   1710
         TabIndex        =   44
         Top             =   4461
         Width           =   2730
      End
      Begin VB.TextBox txtSpecCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   8
         Left            =   1710
         TabIndex        =   42
         Top             =   4002
         Width           =   2730
      End
      Begin VB.TextBox txtSpecCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   7
         Left            =   1710
         TabIndex        =   40
         Top             =   3543
         Width           =   2730
      End
      Begin VB.TextBox txtSpecCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   6
         Left            =   1710
         TabIndex        =   38
         Top             =   3084
         Width           =   2730
      End
      Begin VB.TextBox txtSpecCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   5
         Left            =   1710
         TabIndex        =   36
         Top             =   2625
         Width           =   2730
      End
      Begin VB.TextBox txtSpecCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   4
         Left            =   1710
         TabIndex        =   34
         Top             =   2166
         Width           =   2730
      End
      Begin VB.TextBox txtSpecCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   3
         Left            =   1710
         TabIndex        =   32
         Top             =   1707
         Width           =   2730
      End
      Begin VB.TextBox txtSpecCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   2
         Left            =   1710
         TabIndex        =   30
         Top             =   1248
         Width           =   2730
      End
      Begin VB.TextBox txtSpecCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   1710
         TabIndex        =   28
         Top             =   789
         Width           =   2730
      End
      Begin VB.TextBox txtSpecCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   1710
         TabIndex        =   26
         Top             =   330
         Width           =   2730
      End
      Begin VB.TextBox txtSpecCodeSJ 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1725
         TabIndex        =   95
         Top             =   4980
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.Label lbSpecCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   10
         Left            =   285
         TabIndex        =   100
         Top             =   4973
         Width           =   1230
      End
      Begin VB.Label lbSpecCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   9
         Left            =   285
         TabIndex        =   43
         Top             =   4532
         Width           =   1230
      End
      Begin VB.Label lbSpecCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   8
         Left            =   285
         TabIndex        =   41
         Top             =   4081
         Width           =   1230
      End
      Begin VB.Label lbSpecCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   7
         Left            =   285
         TabIndex        =   39
         Top             =   3615
         Width           =   1230
      End
      Begin VB.Label lbSpecCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   6
         Left            =   285
         TabIndex        =   37
         Top             =   3149
         Width           =   1230
      End
      Begin VB.Label lbSpecCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   5
         Left            =   285
         TabIndex        =   35
         Top             =   2683
         Width           =   1230
      End
      Begin VB.Label lbSpecCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   4
         Left            =   285
         TabIndex        =   33
         Top             =   2217
         Width           =   1230
      End
      Begin VB.Label lbSpecCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   3
         Left            =   285
         TabIndex        =   31
         Top             =   1751
         Width           =   1230
      End
      Begin VB.Label lbSpecCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   2
         Left            =   285
         TabIndex        =   29
         Top             =   1285
         Width           =   1230
      End
      Begin VB.Label lbSpecCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   1
         Left            =   285
         TabIndex        =   27
         Top             =   819
         Width           =   1230
      End
      Begin VB.Label lbSpecCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   0
         Left            =   285
         TabIndex        =   25
         Top             =   368
         Width           =   1230
      End
      Begin VB.Shape Shape10 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   390
         Left            =   150
         Shape           =   4  '둥근 사각형
         Top             =   315
         Width           =   1500
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   150
         Shape           =   4  '둥근 사각형
         Top             =   1683
         Width           =   1500
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   150
         Shape           =   4  '둥근 사각형
         Top             =   2615
         Width           =   1500
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   150
         Shape           =   4  '둥근 사각형
         Top             =   2149
         Width           =   1500
      End
      Begin VB.Shape Shape15 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   150
         Shape           =   4  '둥근 사각형
         Top             =   1217
         Width           =   1500
      End
      Begin VB.Shape Shape14 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   390
         Left            =   150
         Shape           =   4  '둥근 사각형
         Top             =   4479
         Width           =   1500
      End
      Begin VB.Shape Shape13 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   150
         Shape           =   4  '둥근 사각형
         Top             =   4013
         Width           =   1500
      End
      Begin VB.Shape Shape12 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   150
         Shape           =   4  '둥근 사각형
         Top             =   3547
         Width           =   1500
      End
      Begin VB.Shape Shape11 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   150
         Shape           =   4  '둥근 사각형
         Top             =   3081
         Width           =   1500
      End
      Begin VB.Shape Shape27 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   150
         Shape           =   4  '둥근 사각형
         Top             =   751
         Width           =   1500
      End
      Begin VB.Shape Shape28 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   390
         Left            =   150
         Shape           =   4  '둥근 사각형
         Top             =   4920
         Width           =   1500
      End
   End
   Begin BHButton.BHImageButton BHBFunctionSave 
      Height          =   585
      Left            =   3525
      TabIndex        =   1
      Top             =   5790
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1032
      Caption         =   "저 장"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton BHBFunctionEnd 
      Height          =   585
      Left            =   5460
      TabIndex        =   2
      Top             =   5790
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1032
      Caption         =   "닫 기"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton BHBSvPvCodeShow 
      Height          =   585
      Left            =   450
      TabIndex        =   73
      Top             =   5790
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1032
      Caption         =   "항목코드 보기"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin VB.Frame fraPVcode 
      BackColor       =   &H8000000E&
      Caption         =   "항목코드(PV)"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5445
      Left            =   5730
      TabIndex        =   52
      Top             =   1095
      Width           =   4680
      Begin VB.TextBox txtPvCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   10
         Left            =   1725
         TabIndex        =   96
         Top             =   4890
         Width           =   2730
      End
      Begin VB.TextBox txtPvCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   9
         Left            =   1725
         TabIndex        =   62
         Top             =   4422
         Width           =   2730
      End
      Begin VB.TextBox txtPvCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   8
         Left            =   1725
         TabIndex        =   61
         Top             =   3959
         Width           =   2730
      End
      Begin VB.TextBox txtPvCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   7
         Left            =   1725
         TabIndex        =   60
         Top             =   3496
         Width           =   2730
      End
      Begin VB.TextBox txtPvCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   6
         Left            =   1725
         TabIndex        =   59
         Top             =   3033
         Width           =   2730
      End
      Begin VB.TextBox txtPvCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   5
         Left            =   1725
         TabIndex        =   58
         Top             =   2570
         Width           =   2730
      End
      Begin VB.TextBox txtPvCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   4
         Left            =   1725
         TabIndex        =   57
         Top             =   2107
         Width           =   2730
      End
      Begin VB.TextBox txtPvCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   3
         Left            =   1725
         TabIndex        =   56
         Top             =   1644
         Width           =   2730
      End
      Begin VB.TextBox txtPvCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   2
         Left            =   1725
         TabIndex        =   55
         Top             =   1181
         Width           =   2730
      End
      Begin VB.TextBox txtPvCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   1725
         TabIndex        =   54
         Top             =   718
         Width           =   2730
      End
      Begin VB.TextBox txtPvCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1725
         TabIndex        =   53
         Top             =   315
         Width           =   2730
      End
      Begin VB.Label lbPvCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   10
         Left            =   285
         TabIndex        =   99
         Top             =   4943
         Width           =   1230
      End
      Begin VB.Label lbPvCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   9
         Left            =   300
         TabIndex        =   72
         Top             =   4493
         Width           =   1230
      End
      Begin VB.Label lbPvCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   8
         Left            =   300
         TabIndex        =   71
         Top             =   4028
         Width           =   1230
      End
      Begin VB.Label lbPvCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   7
         Left            =   300
         TabIndex        =   70
         Top             =   3570
         Width           =   1230
      End
      Begin VB.Label lbPvCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   6
         Left            =   300
         TabIndex        =   69
         Top             =   3098
         Width           =   1230
      End
      Begin VB.Label lbPvCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   5
         Left            =   300
         TabIndex        =   68
         Top             =   2633
         Width           =   1230
      End
      Begin VB.Label lbPvCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   4
         Left            =   300
         TabIndex        =   67
         Top             =   2168
         Width           =   1230
      End
      Begin VB.Label lbPvCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   3
         Left            =   300
         TabIndex        =   66
         Top             =   1703
         Width           =   1230
      End
      Begin VB.Label lbPvCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   2
         Left            =   300
         TabIndex        =   65
         Top             =   1238
         Width           =   1230
      End
      Begin VB.Label lbPvCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   1
         Left            =   300
         TabIndex        =   64
         Top             =   773
         Width           =   1230
      End
      Begin VB.Label lbPvCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   0
         Left            =   300
         TabIndex        =   63
         Top             =   323
         Width           =   1230
      End
      Begin VB.Shape Shape46 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   165
         Shape           =   4  '둥근 사각형
         Top             =   2100
         Width           =   1500
      End
      Begin VB.Shape Shape45 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   165
         Shape           =   4  '둥근 사각형
         Top             =   1635
         Width           =   1500
      End
      Begin VB.Shape Shape44 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   165
         Shape           =   4  '둥근 사각형
         Top             =   1170
         Width           =   1500
      End
      Begin VB.Shape Shape43 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   165
         Shape           =   4  '둥근 사각형
         Top             =   705
         Width           =   1500
      End
      Begin VB.Shape Shape42 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   390
         Left            =   165
         Shape           =   4  '둥근 사각형
         Top             =   270
         Width           =   1500
      End
      Begin VB.Shape Shape41 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   150
         Shape           =   4  '둥근 사각형
         Top             =   4425
         Width           =   1500
      End
      Begin VB.Shape Shape40 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   150
         Shape           =   4  '둥근 사각형
         Top             =   3960
         Width           =   1500
      End
      Begin VB.Shape Shape39 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   150
         Shape           =   4  '둥근 사각형
         Top             =   3495
         Width           =   1500
      End
      Begin VB.Shape Shape38 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   150
         Shape           =   4  '둥근 사각형
         Top             =   3030
         Width           =   1500
      End
      Begin VB.Shape Shape37 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   150
         Shape           =   4  '둥근 사각형
         Top             =   2565
         Width           =   1500
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   390
         Left            =   150
         Shape           =   4  '둥근 사각형
         Top             =   4890
         Width           =   1500
      End
   End
   Begin VB.Frame fraNGcode 
      BackColor       =   &H8000000E&
      Caption         =   "불량코드"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5445
      Left            =   5610
      TabIndex        =   74
      Top             =   300
      Width           =   4680
      Begin VB.TextBox txtNgCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   10
         Left            =   1710
         TabIndex        =   98
         Top             =   4830
         Width           =   2730
      End
      Begin VB.TextBox txtNgCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   1725
         TabIndex        =   84
         Top             =   315
         Width           =   2730
      End
      Begin VB.TextBox txtNgCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   1725
         TabIndex        =   83
         Top             =   766
         Width           =   2730
      End
      Begin VB.TextBox txtNgCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   2
         Left            =   1725
         TabIndex        =   82
         Top             =   1217
         Width           =   2730
      End
      Begin VB.TextBox txtNgCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   3
         Left            =   1725
         TabIndex        =   81
         Top             =   1668
         Width           =   2730
      End
      Begin VB.TextBox txtNgCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   4
         Left            =   1725
         TabIndex        =   80
         Top             =   2119
         Width           =   2730
      End
      Begin VB.TextBox txtNgCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   5
         Left            =   1725
         TabIndex        =   79
         Top             =   2570
         Width           =   2730
      End
      Begin VB.TextBox txtNgCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   6
         Left            =   1725
         TabIndex        =   78
         Top             =   3021
         Width           =   2730
      End
      Begin VB.TextBox txtNgCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   7
         Left            =   1725
         TabIndex        =   77
         Top             =   3472
         Width           =   2730
      End
      Begin VB.TextBox txtNgCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   8
         Left            =   1725
         TabIndex        =   76
         Top             =   3923
         Width           =   2730
      End
      Begin VB.TextBox txtNgCode 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   9
         Left            =   1725
         TabIndex        =   75
         Top             =   4374
         Width           =   2730
      End
      Begin VB.Label lbNGCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   10
         Left            =   285
         TabIndex        =   101
         Top             =   4898
         Width           =   1230
      End
      Begin VB.Label lbNGCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   0
         Left            =   300
         TabIndex        =   94
         Top             =   368
         Width           =   1230
      End
      Begin VB.Label lbNGCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   1
         Left            =   300
         TabIndex        =   93
         Top             =   821
         Width           =   1230
      End
      Begin VB.Label lbNGCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   2
         Left            =   300
         TabIndex        =   92
         Top             =   1274
         Width           =   1230
      End
      Begin VB.Label lbNGCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   3
         Left            =   300
         TabIndex        =   91
         Top             =   1727
         Width           =   1230
      End
      Begin VB.Label lbNGCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   4
         Left            =   300
         TabIndex        =   90
         Top             =   2180
         Width           =   1230
      End
      Begin VB.Label lbNGCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   5
         Left            =   300
         TabIndex        =   89
         Top             =   2633
         Width           =   1230
      End
      Begin VB.Label lbNGCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   6
         Left            =   300
         TabIndex        =   88
         Top             =   3086
         Width           =   1230
      End
      Begin VB.Label lbNGCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   7
         Left            =   300
         TabIndex        =   87
         Top             =   3539
         Width           =   1230
      End
      Begin VB.Label lbNGCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   8
         Left            =   300
         TabIndex        =   86
         Top             =   3992
         Width           =   1230
      End
      Begin VB.Label lbNGCode 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00000080&
         Caption         =   "항목코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Index           =   9
         Left            =   300
         TabIndex        =   85
         Top             =   4445
         Width           =   1230
      End
      Begin VB.Shape Shape26 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   150
         Shape           =   4  '둥근 사각형
         Top             =   2565
         Width           =   1500
      End
      Begin VB.Shape Shape25 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   150
         Shape           =   4  '둥근 사각형
         Top             =   3018
         Width           =   1500
      End
      Begin VB.Shape Shape24 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   150
         Shape           =   4  '둥근 사각형
         Top             =   3471
         Width           =   1500
      End
      Begin VB.Shape Shape23 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   150
         Shape           =   4  '둥근 사각형
         Top             =   3924
         Width           =   1500
      End
      Begin VB.Shape Shape21 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   165
         Shape           =   4  '둥근 사각형
         Top             =   300
         Width           =   1500
      End
      Begin VB.Shape Shape20 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   165
         Shape           =   4  '둥근 사각형
         Top             =   753
         Width           =   1500
      End
      Begin VB.Shape Shape19 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   165
         Shape           =   4  '둥근 사각형
         Top             =   1206
         Width           =   1500
      End
      Begin VB.Shape Shape18 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   165
         Shape           =   4  '둥근 사각형
         Top             =   1659
         Width           =   1500
      End
      Begin VB.Shape Shape17 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   165
         Shape           =   4  '둥근 사각형
         Top             =   2112
         Width           =   1500
      End
      Begin VB.Shape Shape22 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   150
         Shape           =   4  '둥근 사각형
         Top             =   4377
         Width           =   1500
      End
      Begin VB.Shape Shape29 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   150
         Shape           =   4  '둥근 사각형
         Top             =   4830
         Width           =   1500
      End
   End
   Begin VB.Frame fraFunc 
      BackColor       =   &H8000000E&
      Caption         =   "환경설정"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5445
      Left            =   465
      TabIndex        =   0
      Top             =   300
      Width           =   5085
      Begin VB.ComboBox cboMesNetDrive 
         Height          =   300
         Left            =   300
         TabIndex        =   51
         Text            =   "선택"
         Top             =   4710
         Width           =   1140
      End
      Begin VB.TextBox txtMESNetPass 
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3735
         TabIndex        =   49
         Text            =   "mes"
         Top             =   5040
         Width           =   1140
      End
      Begin VB.TextBox txtMESNetID 
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2145
         TabIndex        =   47
         Text            =   "sblmes"
         Top             =   5040
         Width           =   1125
      End
      Begin VB.TextBox txtMESNetDriveIP 
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   360
         TabIndex        =   45
         Text            =   "17.4.52.125"
         Top             =   5055
         Width           =   1395
      End
      Begin VB.TextBox txtDataRow 
         Alignment       =   2  '가운데 맞춤
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3720
         TabIndex        =   22
         Text            =   "40"
         Top             =   2895
         Visible         =   0   'False
         Width           =   1065
      End
      Begin BHButton.BHImageButton BHBPathSet 
         Height          =   375
         Index           =   0
         Left            =   300
         TabIndex        =   16
         Top             =   3210
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   661
         Caption         =   "경로설정"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin VB.TextBox txtFunc 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   1845
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   2415
         Width           =   2730
      End
      Begin VB.TextBox txtFunc 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   1845
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1920
         Width           =   2730
      End
      Begin VB.TextBox txtFunc 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   1845
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1410
         Width           =   2730
      End
      Begin VB.TextBox txtFunc 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1845
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   915
         Width           =   2730
      End
      Begin VB.TextBox txtFunc 
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1845
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   420
         Width           =   2730
      End
      Begin BHButton.BHImageButton BHBPathSet 
         Height          =   375
         Index           =   1
         Left            =   3405
         TabIndex        =   17
         Top             =   3540
         Visible         =   0   'False
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   661
         Caption         =   "경로설정"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton BHBPathSet 
         Height          =   375
         Index           =   2
         Left            =   300
         TabIndex        =   18
         Top             =   3930
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   661
         Caption         =   "경로설정"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000E&
         Caption         =   "PW"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   3360
         TabIndex        =   50
         Top             =   5115
         Width           =   360
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   1815
         TabIndex        =   48
         Top             =   5115
         Width           =   300
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "IP"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   75
         TabIndex        =   46
         Top             =   5130
         Width           =   300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "DATA 개수"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2655
         TabIndex        =   23
         Top             =   2955
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblLogPath 
         BackColor       =   &H8000000E&
         Caption         =   "D:\MES\LOG"
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
         Height          =   240
         Left            =   1875
         TabIndex        =   21
         Top             =   4035
         Width           =   3090
      End
      Begin VB.Label lblSendPath 
         BackColor       =   &H8000000E&
         Caption         =   "S:\"
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
         Height          =   240
         Left            =   1875
         TabIndex        =   20
         Top             =   4755
         Width           =   2865
      End
      Begin VB.Label lblFilePath 
         BackColor       =   &H8000000E&
         Caption         =   "D:\MES\SEND"
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
         Height          =   240
         Left            =   1890
         TabIndex        =   19
         Top             =   3315
         Width           =   3090
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "MES 로그경로"
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
         Height          =   345
         Left            =   315
         TabIndex        =   15
         Top             =   3675
         Width           =   1470
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "MES NetDrive"
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
         Height          =   210
         Left            =   315
         TabIndex        =   14
         Top             =   4395
         Width           =   1545
      End
      Begin VB.Label lblFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         Caption         =   "프로세스"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   4
         Left            =   510
         TabIndex        =   13
         Top             =   2430
         Width           =   960
      End
      Begin VB.Label lblFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         Caption         =   "공정코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   3
         Left            =   510
         TabIndex        =   11
         Top             =   1935
         Width           =   960
      End
      Begin VB.Label lblFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         Caption         =   "라인넘버"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   2
         Left            =   510
         TabIndex        =   9
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label lblFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         Caption         =   "설비이름"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   510
         TabIndex        =   7
         Top             =   960
         Width           =   960
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "MES 파일경로"
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
         Height          =   345
         Left            =   315
         TabIndex        =   5
         Top             =   2955
         Width           =   1470
      End
      Begin VB.Label lblFunc 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         Caption         =   "설비코드"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   0
         Left            =   510
         TabIndex        =   4
         Top             =   450
         Width           =   960
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   390
         Shape           =   4  '둥근 사각형
         Top             =   375
         Width           =   1245
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   390
         Shape           =   4  '둥근 사각형
         Top             =   870
         Width           =   1245
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   390
         Shape           =   4  '둥근 사각형
         Top             =   1365
         Width           =   1245
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   390
         Shape           =   4  '둥근 사각형
         Top             =   1860
         Width           =   1245
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   420
         Left            =   390
         Shape           =   4  '둥근 사각형
         Top             =   2355
         Width           =   1245
      End
   End
   Begin VB.Shape Shape16 
      BorderColor     =   &H00000080&
      BorderWidth     =   18
      Height          =   6495
      Left            =   150
      Top             =   90
      Width           =   10455
   End
End
Attribute VB_Name = "frmMESFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BHBFunctionEnd_Click()
    Unload Me
End Sub

Private Sub BHBFunctionSave_Click()
Dim i As Integer

    sMESEquipCode = Me.txtFunc(0).Text
    sMESEquipName = Me.txtFunc(1).Text
    sMESLineNum = Me.txtFunc(2).Text
    sMESProgressCode = Me.txtFunc(3).Text
    sMESProcess = Me.txtFunc(4).Text
    sMESFileSavePath = Me.lblFilePath.Caption
    sMESFileSendPath = Me.lblSendPath.Caption
    sMESLogSavePath = Me.lblLogPath.Caption
    sMesPCIP = Me.txtMESNetDriveIP.Text
    sMesPCID = Me.txtMESNetID.Text
    sMesPCPW = Me.txtMESNetPass.Text
    sParamName_SVsj = txtSpecCodeSJ.Text
    For i = 1 To 11
        sParamName_SV(iNowRecipeID, i) = Me.txtSpecCode(i - 1).Text
        sParamName_PV(iNowRecipeID, i) = Me.txtPvCode(i - 1).Text
        sParamName_NG(iNowRecipeID, i) = Me.txtNgCode(i - 1).Text
    Next i
    Call DJ_MESFunctionSave
    
    g_TimeoutInterval = CLng(txtTimeoutInterval.Text)
    g_TimeoutRetry = CLng(txtTimeoutRetry.Text)
    
    Call SaveTimeoutParam
    
End Sub

Private Sub BHBPathSet_Click(Index As Integer)
Dim i As Integer
    For i = 0 To 2
        bPathSelect(i) = False
    Next i
    
    bPathSelect(Index) = True
    
    frmMESDirPath.Show
End Sub

Private Sub BHBSvPvCodeShow_Click()
    If Me.BHBSvPvCodeShow.Caption = "항목코드 보기" Then
        Me.BHBSvPvCodeShow.Caption = "환경설정 보기"
        Me.fraSVCode.Visible = True
        Me.fraPVcode.Visible = True
        Me.fraFunc.Visible = False
        Me.fraFunc.Visible = False
    Else
        Me.BHBSvPvCodeShow.Caption = "항목코드 보기"
        Me.fraSVCode.Visible = False
        Me.fraPVcode.Visible = False
        Me.fraFunc.Visible = True
        Me.fraFunc.Visible = True
    End If
End Sub

Private Sub cboMesNetDrive_click()
    Me.lblSendPath.Caption = Me.cboMesNetDrive.Text
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim i As Integer
    For i = 75 To 90
        cboMesNetDrive.AddItem Chr(i) & ":"
    Next i
    Me.txtFunc(0).Text = sMESEquipCode
    Me.txtFunc(1).Text = sMESEquipName
    Me.txtFunc(2).Text = sMESLineNum
    Me.txtFunc(3).Text = sMESProgressCode
    Me.txtFunc(4).Text = sMESProcess
    Me.lblFilePath.Caption = sMESFileSavePath
    Me.lblSendPath.Caption = sMESFileSendPath
    Me.lblLogPath.Caption = sMESLogSavePath
    Me.txtMESNetDriveIP.Text = sMesPCIP
    Me.txtMESNetID.Text = sMesPCID
    Me.txtMESNetPass.Text = sMesPCPW
    For i = 0 To iToolCount / 2 - 1
        Me.lbSpecCode(i).Caption = sSpecName(i)
        Me.lbPvCode(i).Caption = sSpecName(i)
        Me.lbNGCode(i).Caption = sSpecName(i)
        Me.txtSpecCode(i).Text = sParamName_SV(iNowRecipeID, i + 1)
        Me.txtSpecCodeSJ.Text = sParamName_SVsj
        Me.txtPvCode(i).Text = sParamName_PV(iNowRecipeID, i + 1)
        Me.txtNgCode(i).Text = sParamName_NG(iNowRecipeID, i + 1)
    Next i
    Me.fraSVCode.Visible = False
    Me.fraPVcode.Visible = False
    Me.fraFunc.Visible = True
    Me.fraFunc.Visible = True
    
    txtTimeoutInterval.Text = CStr(g_TimeoutInterval)
    txtTimeoutRetry.Text = CStr(g_TimeoutRetry)
    
End Sub

Private Sub txtTimeoutInterval_Change()

    If CheckTextBox(txtTimeoutInterval, 0, 60000) = True Then
        
    End If
    
End Sub
