VERSION 5.00
Object = "{A306B141-AE98-11D3-83AE-00A024BDBF2B}#3.0#0"; "ActMulti.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0ACBBD11-6E21-11D4-9751-0060089571FC}#1.0#0"; "CogDisplay.ocx"
Object = "{6DF32DBD-B2DD-4895-A028-AE7FCD043771}#1.15#0"; "uEyeCam.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "Samsung SDI - Main (Rev. 2024-08-28)"
   ClientHeight    =   15015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19185
   ControlBox      =   0   'False
   FillColor       =   &H80000005&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   15015
   ScaleWidth      =   19185
   StartUpPosition =   2  '화면 가운데
   WindowState     =   2  '최대화
   Begin VB.Timer tmrLight 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3060
      Top             =   180
   End
   Begin VB.PictureBox picScreenShotSave 
      AutoSize        =   -1  'True
      Height          =   435
      Left            =   510
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   165
      Top             =   180
      Visible         =   0   'False
      Width           =   555
      Begin VB.Image ImageScreen 
         Height          =   285
         Left            =   120
         Top             =   60
         Width           =   300
      End
   End
   Begin VB.Timer TmrPLCSock 
      Interval        =   200
      Left            =   2115
      Top             =   195
   End
   Begin VB.Timer TmrMESSock 
      Interval        =   1000
      Left            =   2580
      Top             =   195
   End
   Begin VB.Timer tmrMelsec 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1650
      Top             =   195
   End
   Begin MSWinsockLib.Winsock WinsockMES 
      Left            =   3510
      Top             =   195
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin BHButton.BHImageButton btnReloadSpec 
      Height          =   525
      Left            =   12840
      TabIndex        =   146
      Top             =   8850
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   926
      Caption         =   "내려 받기"
      CaptionChecked  =   "BHImageButton2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin VB.Frame fraSpecName 
      BackColor       =   &H8000000E&
      Caption         =   "항목 이름 변경"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   14310
      TabIndex        =   4
      Top             =   6180
      Visible         =   0   'False
      Width           =   4785
      Begin VB.Frame fraFunc 
         BackColor       =   &H8000000E&
         Caption         =   "기능설정"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2430
         Left            =   2520
         TabIndex        =   78
         Top             =   270
         Width           =   2160
         Begin VB.Frame Frame12 
            BackColor       =   &H8000000E&
            Caption         =   "저장모드"
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1350
            Left            =   120
            TabIndex        =   81
            Top             =   600
            Width           =   1965
            Begin VB.OptionButton Option1 
               BackColor       =   &H8000000E&
               Caption         =   "Jpg"
               BeginProperty Font 
                  Name            =   "돋움"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   1
               Left            =   1095
               TabIndex        =   85
               Top             =   330
               Width           =   690
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H8000000E&
               Caption         =   "Bmp"
               BeginProperty Font 
                  Name            =   "돋움"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   195
               TabIndex        =   84
               Top             =   330
               Value           =   -1  'True
               Width           =   765
            End
            Begin VB.CheckBox chkNGImageSave 
               BackColor       =   &H8000000E&
               Caption         =   "NG Image 저장"
               BeginProperty Font 
                  Name            =   "돋움"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   83
               Top             =   945
               Width           =   1785
            End
            Begin VB.CheckBox chkOKImageSave 
               BackColor       =   &H8000000E&
               Caption         =   "OK Image 저장"
               BeginProperty Font 
                  Name            =   "돋움"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   82
               Top             =   645
               Width           =   1710
            End
         End
         Begin VB.CheckBox chkWriteDataSave 
            BackColor       =   &H8000000E&
            Caption         =   "결과 데이터 저장"
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   150
            TabIndex        =   80
            ToolTipText     =   "D:\ImageSave\ 이하 날짜별 , 모델별 폴더 아래 확장자 csv 로 결과값이 누적되어 저장됩니다."
            Top             =   2040
            Width           =   1935
         End
         Begin VB.CheckBox chkCamPass 
            BackColor       =   &H8000000E&
            Caption         =   "Vision Pass"
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   150
            TabIndex        =   79
            ToolTipText     =   "1회 촬영 후 검사를 하지 않고 OK 출력신호를 PLC 로 보냅니다."
            Top             =   315
            Width           =   1710
         End
      End
      Begin VB.TextBox txtSpecName 
         Alignment       =   1  '오른쪽 맞춤
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   540
         TabIndex        =   14
         Text            =   "항목 1"
         Top             =   330
         Width           =   1875
      End
      Begin VB.TextBox txtSpecName 
         Alignment       =   1  '오른쪽 맞춤
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   540
         TabIndex        =   13
         Text            =   "항목 1"
         Top             =   720
         Width           =   1875
      End
      Begin VB.TextBox txtSpecName 
         Alignment       =   1  '오른쪽 맞춤
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   540
         TabIndex        =   12
         Text            =   "항목 1"
         Top             =   1110
         Width           =   1875
      End
      Begin VB.TextBox txtSpecName 
         Alignment       =   1  '오른쪽 맞춤
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   540
         TabIndex        =   11
         Text            =   "항목 1"
         Top             =   1500
         Width           =   1875
      End
      Begin VB.TextBox txtSpecName 
         Alignment       =   1  '오른쪽 맞춤
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   540
         TabIndex        =   10
         Text            =   "항목 1"
         Top             =   1890
         Width           =   1875
      End
      Begin VB.TextBox txtSpecName 
         Alignment       =   1  '오른쪽 맞춤
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   540
         TabIndex        =   9
         Text            =   "항목 1"
         Top             =   2280
         Width           =   1875
      End
      Begin VB.TextBox txtSpecName 
         Alignment       =   1  '오른쪽 맞춤
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   540
         TabIndex        =   8
         Text            =   "항목 1"
         Top             =   2670
         Width           =   1875
      End
      Begin VB.TextBox txtSpecName 
         Alignment       =   1  '오른쪽 맞춤
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   540
         TabIndex        =   7
         Text            =   "항목 1"
         Top             =   3060
         Width           =   1875
      End
      Begin VB.TextBox txtSpecName 
         Alignment       =   1  '오른쪽 맞춤
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   540
         TabIndex        =   6
         Text            =   "항목 1"
         Top             =   3450
         Width           =   1875
      End
      Begin VB.TextBox txtSpecName 
         Alignment       =   1  '오른쪽 맞춤
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   540
         TabIndex        =   5
         Text            =   "항목 1"
         Top             =   3840
         Width           =   1875
      End
      Begin BHButton.BHImageButton BHBFuncSave 
         Height          =   540
         Left            =   2820
         TabIndex        =   30
         ToolTipText     =   "저장 - SPEC , 항목이름 변경 , 기능 설정이 저장됩니다. "
         Top             =   2880
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   953
         Caption         =   "저 장"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton BHBFuncCancel 
         Height          =   540
         Left            =   2820
         TabIndex        =   86
         ToolTipText     =   "저장 - SPEC , 항목이름 변경 , 기능 설정이 저장됩니다. "
         Top             =   3540
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   953
         Caption         =   "취 소"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin VB.Label lblSpecNum 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "1 :"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   24
         Top             =   375
         Width           =   255
      End
      Begin VB.Label lblSpecNum 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "2 :"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   23
         Top             =   765
         Width           =   255
      End
      Begin VB.Label lblSpecNum 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "3 :"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   22
         Top             =   1155
         Width           =   255
      End
      Begin VB.Label lblSpecNum 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "4 :"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   21
         Top             =   1545
         Width           =   255
      End
      Begin VB.Label lblSpecNum 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "5 :"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   20
         Top             =   1935
         Width           =   255
      End
      Begin VB.Label lblSpecNum 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "6 :"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   210
         TabIndex        =   19
         Top             =   2325
         Width           =   255
      End
      Begin VB.Label lblSpecNum 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "7 :"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   210
         TabIndex        =   18
         Top             =   2715
         Width           =   255
      End
      Begin VB.Label lblSpecNum 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "8 :"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   210
         TabIndex        =   17
         Top             =   3105
         Width           =   255
      End
      Begin VB.Label lblSpecNum 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "9 :"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   210
         TabIndex        =   16
         Top             =   3495
         Width           =   255
      End
      Begin VB.Label lblSpecNum 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "10 :"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   105
         TabIndex        =   15
         Top             =   3900
         Width           =   360
      End
   End
   Begin VB.CheckBox chkManualSave 
      BackColor       =   &H8000000E&
      Caption         =   "수동검사 클릭시 이미지 및 데이터 저장"
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
      Left            =   12870
      TabIndex        =   126
      Top             =   10200
      Width           =   4065
   End
   Begin VB.CheckBox chkManualAcq 
      BackColor       =   &H8000000E&
      Caption         =   "수동검사 클릭시 촬영 후 검사"
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
      Left            =   12870
      TabIndex        =   125
      ToolTipText     =   "Check 해재 시 수동검사를 클릭하면 마스터 이미지로 검사를 실행합니다."
      Top             =   9930
      Width           =   3495
   End
   Begin VB.Frame Frame11 
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
      Height          =   3195
      Left            =   12840
      TabIndex        =   101
      Top             =   900
      Width           =   6300
      Begin VB.Frame Frame16 
         BackColor       =   &H00FFFFFF&
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
         Height          =   1395
         Left            =   150
         TabIndex        =   111
         Top             =   1620
         Width           =   6015
         Begin VB.Label lblResultNSD 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   5
            Left            =   5100
            TabIndex        =   123
            Top             =   885
            Width           =   165
         End
         Begin VB.Label lblResultTitleNSD 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "우2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   5
            Left            =   4965
            TabIndex        =   122
            Top             =   420
            Width           =   420
         End
         Begin VB.Label lblResultNSD 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   4
            Left            =   3660
            TabIndex        =   121
            Top             =   885
            Width           =   165
         End
         Begin VB.Label lblResultTitleNSD 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "우1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   4
            Left            =   3525
            TabIndex        =   120
            Top             =   420
            Width           =   420
         End
         Begin VB.Label lblResultNSD 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Index           =   3
            Left            =   3255
            TabIndex        =   119
            Top             =   2355
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblResultTitleNSD 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "하2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   3
            Left            =   3105
            TabIndex        =   118
            Top             =   1770
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label lblResultNSD 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   2220
            TabIndex        =   117
            Top             =   885
            Width           =   165
         End
         Begin VB.Label lblResultTitleNSD 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "하1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   2085
            TabIndex        =   116
            Top             =   420
            Width           =   420
         End
         Begin VB.Label lblResultNSD 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Index           =   1
            Left            =   1290
            TabIndex        =   115
            Top             =   2355
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblResultTitleNSD 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "상2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   1
            Left            =   1140
            TabIndex        =   114
            Top             =   1770
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label lblResultNSD 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   780
            TabIndex        =   113
            Top             =   885
            Width           =   165
         End
         Begin VB.Label lblResultTitleNSD 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "상1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   652
            TabIndex        =   112
            Top             =   420
            Width           =   420
         End
         Begin VB.Shape shpResultTitleNSD 
            BackColor       =   &H00000000&
            BackStyle       =   1  '투명하지 않음
            Height          =   405
            Index           =   0
            Left            =   180
            Top             =   360
            Width           =   1365
         End
         Begin VB.Shape shpResultNSD 
            BackStyle       =   1  '투명하지 않음
            Height          =   405
            Index           =   0
            Left            =   180
            Top             =   825
            Width           =   1365
         End
         Begin VB.Shape shpResultNSD 
            BackStyle       =   1  '투명하지 않음
            Height          =   405
            Index           =   5
            Left            =   4500
            Top             =   825
            Width           =   1365
         End
         Begin VB.Shape shpResultTitleNSD 
            BackColor       =   &H00000000&
            BackStyle       =   1  '투명하지 않음
            Height          =   405
            Index           =   5
            Left            =   4500
            Top             =   360
            Width           =   1365
         End
         Begin VB.Shape shpResultNSD 
            BackStyle       =   1  '투명하지 않음
            Height          =   405
            Index           =   4
            Left            =   3060
            Top             =   825
            Width           =   1365
         End
         Begin VB.Shape shpResultTitleNSD 
            BackColor       =   &H00000000&
            BackStyle       =   1  '투명하지 않음
            Height          =   405
            Index           =   4
            Left            =   3060
            Top             =   360
            Width           =   1365
         End
         Begin VB.Shape shpResultNSD 
            BackStyle       =   1  '투명하지 않음
            Height          =   525
            Index           =   3
            Left            =   2880
            Top             =   2265
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Shape shpResultTitleNSD 
            BackColor       =   &H00000000&
            BackStyle       =   1  '투명하지 않음
            Height          =   525
            Index           =   3
            Left            =   2880
            Top             =   1680
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Shape shpResultNSD 
            BackStyle       =   1  '투명하지 않음
            Height          =   405
            Index           =   2
            Left            =   1620
            Top             =   825
            Width           =   1365
         End
         Begin VB.Shape shpResultTitleNSD 
            BackColor       =   &H00000000&
            BackStyle       =   1  '투명하지 않음
            Height          =   405
            Index           =   2
            Left            =   1620
            Top             =   360
            Width           =   1365
         End
         Begin VB.Shape shpResultNSD 
            BackStyle       =   1  '투명하지 않음
            Height          =   525
            Index           =   1
            Left            =   915
            Top             =   2265
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Shape shpResultTitleNSD 
            BackColor       =   &H00000000&
            BackStyle       =   1  '투명하지 않음
            Height          =   525
            Index           =   1
            Left            =   915
            Top             =   1680
            Visible         =   0   'False
            Width           =   945
         End
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "W/H"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Index           =   0
         Left            =   150
         TabIndex        =   102
         Top             =   180
         Width           =   6015
         Begin VB.Label lblResultTitleWH 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "높이2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   3
            Left            =   4845
            TabIndex        =   110
            Top             =   420
            Width           =   660
         End
         Begin VB.Label lblResultWH 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   5100
            TabIndex        =   109
            Top             =   885
            Width           =   165
         End
         Begin VB.Label lblResultTitleWH 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "높이1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   3405
            TabIndex        =   108
            Top             =   420
            Width           =   660
         End
         Begin VB.Label lblResultWH 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   3660
            TabIndex        =   107
            Top             =   885
            Width           =   165
         End
         Begin VB.Label lblResultTitleWH 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "너비2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   1965
            TabIndex        =   106
            Top             =   420
            Width           =   660
         End
         Begin VB.Label lblResultWH 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   2220
            TabIndex        =   105
            Top             =   885
            Width           =   165
         End
         Begin VB.Shape shpResultTitleWH 
            BackColor       =   &H00000000&
            BackStyle       =   1  '투명하지 않음
            Height          =   405
            Index           =   3
            Left            =   4500
            Top             =   360
            Width           =   1365
         End
         Begin VB.Shape shpResultTitleWH 
            BackColor       =   &H00000000&
            BackStyle       =   1  '투명하지 않음
            Height          =   405
            Index           =   2
            Left            =   3060
            Top             =   360
            Width           =   1365
         End
         Begin VB.Shape shpResultTitleWH 
            BackColor       =   &H00000000&
            BackStyle       =   1  '투명하지 않음
            Height          =   405
            Index           =   1
            Left            =   1620
            Top             =   360
            Width           =   1365
         End
         Begin VB.Label lblResultTitleWH 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "너비1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   532
            TabIndex        =   104
            Top             =   420
            Width           =   660
         End
         Begin VB.Label lblResultWH 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  '투명
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   780
            TabIndex        =   103
            Top             =   885
            Width           =   165
         End
         Begin VB.Shape shpResultTitleWH 
            BackColor       =   &H00000000&
            BackStyle       =   1  '투명하지 않음
            Height          =   405
            Index           =   0
            Left            =   180
            Top             =   360
            Width           =   1365
         End
         Begin VB.Shape shpResultWH 
            BackStyle       =   1  '투명하지 않음
            Height          =   405
            Index           =   0
            Left            =   180
            Top             =   825
            Width           =   1365
         End
         Begin VB.Shape shpResultWH 
            BackStyle       =   1  '투명하지 않음
            Height          =   405
            Index           =   3
            Left            =   4500
            Top             =   825
            Width           =   1365
         End
         Begin VB.Shape shpResultWH 
            BackStyle       =   1  '투명하지 않음
            Height          =   405
            Index           =   2
            Left            =   3060
            Top             =   825
            Width           =   1365
         End
         Begin VB.Shape shpResultWH 
            BackStyle       =   1  '투명하지 않음
            Height          =   405
            Index           =   1
            Left            =   1620
            Top             =   825
            Width           =   1365
         End
      End
   End
   Begin uEyeCamLib.uEyeCam uEyeCam1 
      Height          =   1005
      Index           =   4
      Left            =   15480
      Top             =   7815
      Visible         =   0   'False
      Width           =   975
      _Version        =   65551
      _ExtentX        =   1720
      _ExtentY        =   1773
      _StockProps     =   1
      AutoSensorShutterMode=   0
      AutoSensorGainMode=   0
   End
   Begin VB.PictureBox picFavisLogo 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   17340
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   810
      ScaleWidth      =   1815
      TabIndex        =   100
      Top             =   0
      Width           =   1845
   End
   Begin VB.PictureBox picTargetLogo 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   0
      Picture         =   "frmMain.frx":2963
      ScaleHeight     =   810
      ScaleWidth      =   4545
      TabIndex        =   99
      Top             =   0
      Width           =   4575
      Begin VB.Timer tmrMesTimeout 
         Left            =   1170
         Top             =   180
      End
      Begin ACTMULTILibCtl.ActEasyIF ActEasyIF 
         Left            =   3990
         OleObjectBlob   =   "frmMain.frx":594D
         Top             =   150
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Count"
      Height          =   3285
      Left            =   15510
      TabIndex        =   87
      Top             =   11670
      Width           =   3615
      Begin BHButton.BHImageButton BHBCountReset 
         Height          =   525
         Left            =   150
         TabIndex        =   94
         Top             =   2610
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   926
         Caption         =   "Count Reset"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin VB.Label lblCountNG 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   2520
         TabIndex        =   93
         Top             =   2010
         Width           =   255
      End
      Begin VB.Label Label37 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "NG"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   2362
         TabIndex        =   92
         Top             =   1455
         Width           =   570
      End
      Begin VB.Label lblCountOK 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   435
         Left            =   840
         TabIndex        =   91
         Top             =   2010
         Width           =   255
      End
      Begin VB.Label lblCountTotal 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   1680
         TabIndex        =   90
         Top             =   885
         Width           =   255
      End
      Begin VB.Label Label36 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "GOOD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   435
         Left            =   427
         TabIndex        =   89
         Top             =   1455
         Width           =   1080
      End
      Begin VB.Label Label35 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   1215
         TabIndex        =   88
         Top             =   315
         Width           =   1185
      End
      Begin VB.Shape Shape11 
         BackStyle       =   1  '투명하지 않음
         Height          =   525
         Left            =   1830
         Top             =   1965
         Width           =   1635
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  '투명하지 않음
         Height          =   525
         Index           =   0
         Left            =   150
         Top             =   1965
         Width           =   1635
      End
      Begin VB.Shape Shape10 
         BackStyle       =   1  '투명하지 않음
         Height          =   525
         Left            =   150
         Top             =   840
         Width           =   3315
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H00000000&
         BackStyle       =   1  '투명하지 않음
         Height          =   525
         Left            =   1830
         Top             =   1410
         Width           =   1635
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  '투명하지 않음
         Height          =   525
         Left            =   150
         Top             =   270
         Width           =   3315
      End
      Begin VB.Shape Shape19 
         BackColor       =   &H00000000&
         BackStyle       =   1  '투명하지 않음
         Height          =   525
         Index           =   0
         Left            =   150
         Top             =   1410
         Width           =   1635
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   945
      Left            =   3120
      TabIndex        =   70
      Top             =   10620
      Width           =   12915
      Begin BHButton.BHImageButton BHBAutoRun 
         Height          =   690
         Left            =   60
         TabIndex        =   71
         ToolTipText     =   "자동검사 - 자동검사를 준비합니다."
         Top             =   180
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1217
         Caption         =   "자동검사"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton BHBStop 
         Height          =   690
         Left            =   1890
         TabIndex        =   72
         ToolTipText     =   "정지 - 자동검사 또는 동영상을 정지합니다."
         Top             =   180
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1217
         Caption         =   "정지"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton BHBManualRun 
         Height          =   690
         Left            =   3720
         TabIndex        =   73
         ToolTipText     =   "수동검사 - 1회 검사 합니다."
         Top             =   180
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1217
         Caption         =   "수동검사"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton BHBLive 
         Height          =   690
         Left            =   5550
         TabIndex        =   74
         ToolTipText     =   "동영상 - 동영상 촬영을 시작합니다."
         Top             =   180
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1217
         Caption         =   "동영상"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton BHBModel 
         Height          =   690
         Left            =   7380
         TabIndex        =   75
         ToolTipText     =   "모델관리 - 모델을 로드 하거나  복사 , 삭제를 진행합니다."
         Top             =   180
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1217
         Caption         =   "모델관리"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton BHBSetting 
         Height          =   690
         Left            =   9210
         TabIndex        =   76
         ToolTipText     =   "검사 설정 - 검사 포인트 및 검사 툴 을 설정합니다."
         Top             =   180
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1217
         Caption         =   "검사설정"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton BHBEnd 
         Height          =   690
         Left            =   11040
         TabIndex        =   77
         ToolTipText     =   "종료 - 프로그램을 종료 합니다. (PLC 와 MES 의 연결이 끊어집니다.)"
         Top             =   180
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1217
         Caption         =   "종료"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
   End
   Begin VB.PictureBox picCamBaseCaption 
      Appearance      =   0  '평면
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   3
      Left            =   6390
      ScaleHeight     =   450
      ScaleWidth      =   6375
      TabIndex        =   45
      Top             =   5670
      Width           =   6405
      Begin VB.Label lblCamBaseCaption 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Camera4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   75
         TabIndex        =   46
         Top             =   15
         Width           =   1410
      End
      Begin VB.Label lblIDCodeNum 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00C0FFFF&
         Caption         =   "NOID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   3
         Left            =   1140
         TabIndex        =   47
         Top             =   15
         Visible         =   0   'False
         Width           =   5130
      End
   End
   Begin VB.PictureBox picCamBaseCaption 
      Appearance      =   0  '평면
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   0
      ScaleHeight     =   450
      ScaleWidth      =   6375
      TabIndex        =   42
      Top             =   5670
      Width           =   6405
      Begin VB.Label lblCamBaseCaption 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Camera3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   75
         TabIndex        =   43
         Top             =   15
         Width           =   1410
      End
      Begin VB.Label lblIDCodeNum 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00C0FFFF&
         Caption         =   "NOID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   2
         Left            =   -810
         TabIndex        =   44
         Top             =   450
         Visible         =   0   'False
         Width           =   5130
      End
   End
   Begin VB.PictureBox picCamBaseCaption 
      Appearance      =   0  '평면
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   6390
      ScaleHeight     =   450
      ScaleWidth      =   6375
      TabIndex        =   39
      Top             =   840
      Width           =   6405
      Begin VB.Label lblCamBaseCaption 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Camera2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   75
         TabIndex        =   40
         Top             =   15
         Width           =   1410
      End
      Begin VB.Label lblIDCodeNum 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00C0FFFF&
         Caption         =   "NOID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   1
         Left            =   1140
         TabIndex        =   41
         Top             =   15
         Width           =   5130
      End
   End
   Begin VB.PictureBox picCamBaseCaption 
      Appearance      =   0  '평면
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   0
      ScaleHeight     =   450
      ScaleWidth      =   6375
      TabIndex        =   36
      Top             =   840
      Width           =   6405
      Begin VB.Label lblCamBaseCaption 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Camera1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   75
         TabIndex        =   38
         Top             =   15
         Width           =   1410
      End
      Begin VB.Label lblIDCodeNum 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H00C0FFFF&
         Caption         =   "NOID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   0
         Left            =   1170
         TabIndex        =   37
         Top             =   15
         Width           =   5100
      End
   End
   Begin uEyeCamLib.uEyeCam uEyeCam1 
      Height          =   1005
      Index           =   0
      Left            =   2715
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
      _Version        =   65551
      _ExtentX        =   1720
      _ExtentY        =   1773
      _StockProps     =   1
      EnableEvents    =   -1  'True
      AutoSensorShutterMode=   0
      AutoSensorGainMode=   0
   End
   Begin VB.Timer TmrLogin 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2250
      Top             =   210
   End
   Begin VB.TextBox txtBackg 
      BackColor       =   &H00C000C0&
      BorderStyle     =   0  '없음
      Height          =   7590
      Index           =   0
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   -7600
      Width           =   19170
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   60
      TabIndex        =   1
      Top             =   11640
      Width           =   15345
      _ExtentX        =   27067
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "D A T A"
      TabPicture(0)   =   "frmMain.frx":597F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "설정 및 MES"
      TabPicture(1)   =   "frmMain.frx":599B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "IO"
      TabPicture(2)   =   "frmMain.frx":59B7
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame10(1)"
      Tab(2).Control(1)=   "Frame2(1)"
      Tab(2).Control(2)=   "Frame3"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Offset"
      TabPicture(3)   =   "frmMain.frx":59D3
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtOffset(0)"
      Tab(3).Control(1)=   "txtOffset(1)"
      Tab(3).Control(2)=   "txtOffset(2)"
      Tab(3).Control(3)=   "txtOffset(3)"
      Tab(3).Control(4)=   "txtOffset(8)"
      Tab(3).Control(5)=   "txtOffset(9)"
      Tab(3).Control(6)=   "txtOffset(4)"
      Tab(3).Control(7)=   "txtOffset(5)"
      Tab(3).Control(8)=   "txtOffset(6)"
      Tab(3).Control(9)=   "txtOffset(7)"
      Tab(3).Control(10)=   "btnApplyOffset"
      Tab(3).Control(11)=   "btnOffsetDefualt"
      Tab(3).Control(12)=   "lblOffset(0)"
      Tab(3).Control(13)=   "lblOffset(1)"
      Tab(3).Control(14)=   "lblOffset(2)"
      Tab(3).Control(15)=   "lblOffset(3)"
      Tab(3).Control(16)=   "lblOffset(4)"
      Tab(3).Control(17)=   "lblOffset(5)"
      Tab(3).Control(18)=   "lblOffset(6)"
      Tab(3).Control(19)=   "lblOffset(7)"
      Tab(3).Control(20)=   "lblOffset(8)"
      Tab(3).Control(21)=   "lblOffset(9)"
      Tab(3).ControlCount=   22
      TabCaption(4)   =   "판정"
      TabPicture(4)   =   "frmMain.frx":59EF
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "chkJudgement(0)"
      Tab(4).Control(1)=   "chkJudgement(1)"
      Tab(4).Control(2)=   "chkJudgement(2)"
      Tab(4).Control(3)=   "chkJudgement(3)"
      Tab(4).Control(4)=   "chkJudgement(4)"
      Tab(4).Control(5)=   "chkJudgement(5)"
      Tab(4).Control(6)=   "chkJudgement(6)"
      Tab(4).Control(7)=   "chkJudgement(7)"
      Tab(4).Control(8)=   "chkJudgement(8)"
      Tab(4).Control(9)=   "chkJudgement(9)"
      Tab(4).Control(10)=   "btnSaveJudgement"
      Tab(4).ControlCount=   11
      Begin VB.TextBox txtOffset 
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
         Left            =   -73590
         TabIndex        =   179
         Text            =   "0"
         Top             =   480
         Width           =   1425
      End
      Begin VB.TextBox txtOffset 
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
         Left            =   -73590
         TabIndex        =   178
         Text            =   "0"
         Top             =   1065
         Width           =   1425
      End
      Begin VB.TextBox txtOffset 
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
         Left            =   -73590
         TabIndex        =   177
         Text            =   "0"
         Top             =   1665
         Width           =   1425
      End
      Begin VB.TextBox txtOffset 
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
         Index           =   3
         Left            =   -73590
         TabIndex        =   176
         Text            =   "0"
         Top             =   2250
         Width           =   1425
      End
      Begin VB.TextBox txtOffset 
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
         Index           =   8
         Left            =   -65850
         TabIndex        =   175
         Text            =   "0"
         Top             =   450
         Width           =   1425
      End
      Begin VB.TextBox txtOffset 
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
         Index           =   9
         Left            =   -65850
         TabIndex        =   174
         Text            =   "0"
         Top             =   1050
         Width           =   1425
      End
      Begin VB.TextBox txtOffset 
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
         Index           =   4
         Left            =   -69600
         TabIndex        =   173
         Text            =   "0"
         Top             =   420
         Width           =   1425
      End
      Begin VB.TextBox txtOffset 
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
         Index           =   5
         Left            =   -69600
         TabIndex        =   172
         Text            =   "0"
         Top             =   1035
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.TextBox txtOffset 
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
         Index           =   6
         Left            =   -69600
         TabIndex        =   171
         Text            =   "0"
         Top             =   1635
         Width           =   1425
      End
      Begin VB.TextBox txtOffset 
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
         Index           =   7
         Left            =   -69600
         TabIndex        =   170
         Text            =   "0"
         Top             =   2250
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Height          =   2970
         Left            =   -74970
         TabIndex        =   151
         Top             =   345
         Width           =   15270
         Begin VB.ComboBox cboROI 
            Height          =   300
            ItemData        =   "frmMain.frx":5A0B
            Left            =   13200
            List            =   "frmMain.frx":5A24
            TabIndex        =   196
            Text            =   "기본영역"
            Top             =   1635
            Width           =   1425
         End
         Begin VB.ComboBox cboROIBase 
            Height          =   300
            ItemData        =   "frmMain.frx":5A57
            Left            =   10950
            List            =   "frmMain.frx":5A70
            TabIndex        =   195
            Text            =   "기본영역"
            Top             =   1635
            Width           =   1425
         End
         Begin VB.CheckBox chkRetry 
            BackColor       =   &H8000000E&
            Caption         =   "NG시 재검사"
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
            Left            =   8070
            TabIndex        =   194
            Top             =   1650
            Value           =   1  '확인
            Width           =   1875
         End
         Begin VB.CheckBox chkAutoLightOnOff 
            BackColor       =   &H8000000E&
            Caption         =   "조명 상시 ON/OFF 타이머 동작"
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
            Left            =   8070
            TabIndex        =   193
            Top             =   1320
            Value           =   1  '확인
            Width           =   3405
         End
         Begin BHButton.BHImageButton btnLoadSystemData 
            Height          =   855
            Left            =   11940
            TabIndex        =   192
            Top             =   300
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   1508
            Caption         =   "시스템값 불러오기"
            CaptionChecked  =   "BHImageButton1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ImgOutLineSize  =   3
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "MES"
            Height          =   1995
            Left            =   90
            TabIndex        =   166
            Top             =   210
            Width           =   7845
            Begin VB.ListBox lstMESSocket 
               Height          =   420
               Left            =   1170
               TabIndex        =   169
               Top             =   1140
               Width           =   4575
            End
            Begin VB.TextBox txtMESIP 
               Height          =   585
               Left            =   1920
               TabIndex        =   168
               Text            =   "Text1"
               Top             =   270
               Width           =   2175
            End
            Begin VB.Label Label2 
               Caption         =   "Local IP"
               Height          =   465
               Left            =   300
               TabIndex        =   167
               Top             =   330
               Width           =   1245
            End
         End
         Begin VB.Frame Frame13 
            BackColor       =   &H00FFFFFF&
            Caption         =   "HDD Check"
            BeginProperty Font 
               Name            =   "돋움"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   600
            Left            =   90
            TabIndex        =   152
            Top             =   2280
            Width           =   7845
            Begin VB.TextBox txtUsedCapPerS 
               Height          =   345
               Index           =   1
               Left            =   4455
               TabIndex        =   154
               Text            =   "0"
               Top             =   180
               Visible         =   0   'False
               Width           =   1530
            End
            Begin VB.TextBox txtUsedCapPerS 
               Height          =   345
               Index           =   0
               Left            =   600
               TabIndex        =   153
               Text            =   "0"
               Top             =   180
               Visible         =   0   'False
               Width           =   1530
            End
            Begin MSComctlLib.ProgressBar PBDrive 
               Height          =   255
               Index           =   0
               Left            =   405
               TabIndex        =   155
               Top             =   225
               Width           =   1920
               _ExtentX        =   3387
               _ExtentY        =   450
               _Version        =   393216
               Appearance      =   0
               Scrolling       =   1
            End
            Begin MSComctlLib.ProgressBar PBDrive 
               Height          =   255
               Index           =   1
               Left            =   4260
               TabIndex        =   156
               Top             =   255
               Width           =   1920
               _ExtentX        =   3387
               _ExtentY        =   450
               _Version        =   393216
               Appearance      =   0
               Scrolling       =   1
            End
            Begin VB.Label lblOverHdd 
               Alignment       =   2  '가운데 맞춤
               BackColor       =   &H000040C0&
               Caption         =   "양호"
               BeginProperty Font 
                  Name            =   "돋움"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   1
               Left            =   6705
               TabIndex        =   162
               ToolTipText     =   "용량이 60% 이하면 '양호' , 이상이면 '경고' , 80% 이상이면 '삭제요망'"
               Top             =   255
               Width           =   1035
            End
            Begin VB.Label lblOverHdd 
               Alignment       =   2  '가운데 맞춤
               BackColor       =   &H0000C000&
               Caption         =   "양호"
               BeginProperty Font 
                  Name            =   "돋움"
                  Size            =   11.25
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   0
               Left            =   2850
               TabIndex        =   161
               ToolTipText     =   "용량이 60% 이하면 '양호' , 이상이면 '경고' , 80% 이상이면 '삭제요망'"
               Top             =   240
               Width           =   1035
            End
            Begin VB.Label lblDrivePer 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   180
               Index           =   1
               Left            =   6225
               TabIndex        =   160
               Top             =   315
               Width           =   165
            End
            Begin VB.Label lblDrivePer 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   180
               Index           =   0
               Left            =   2355
               TabIndex        =   159
               Top             =   270
               Width           =   165
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "D :"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   180
               Left            =   3945
               TabIndex        =   158
               Top             =   300
               Width           =   285
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000E&
               Caption         =   "C :"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9
                  Charset         =   129
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   180
               Left            =   75
               TabIndex        =   157
               Top             =   255
               Width           =   300
            End
         End
         Begin BHButton.BHImageButton BHBMESNetDriveConnect 
            Height          =   330
            Left            =   13455
            TabIndex        =   163
            Top             =   3690
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   582
            Caption         =   "NET 드라이브 연결"
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
         Begin BHButton.BHImageButton BHBMESMain 
            Height          =   930
            Left            =   8010
            TabIndex        =   164
            ToolTipText     =   "MES MAIN - MES 관련 창을 불러옵니다."
            Top             =   210
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   1640
            Caption         =   "MES MAIN"
            CaptionChecked  =   "BHImageButton2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림"
               Size            =   18
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ImgOutLineSize  =   3
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "재검사:"
            Height          =   180
            Left            =   12510
            TabIndex        =   199
            Top             =   1695
            Width           =   600
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "기본검사:"
            Height          =   180
            Left            =   10110
            TabIndex        =   198
            Top             =   1695
            Width           =   780
         End
         Begin VB.Label lblAutoLightInterval 
            Alignment       =   2  '가운데 맞춤
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "(60초)"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   9.75
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   11670
            TabIndex        =   197
            Top             =   1365
            Width           =   630
         End
      End
      Begin VB.Frame Frame10 
         Appearance      =   0  '평면
         Caption         =   "모델"
         ForeColor       =   &H80000008&
         Height          =   2085
         Index           =   1
         Left            =   -64920
         TabIndex        =   139
         Top             =   630
         Width           =   3735
         Begin VB.TextBox txtModelNumber 
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
            Height          =   525
            Left            =   1560
            TabIndex        =   140
            Text            =   "0"
            Top             =   330
            Width           =   1725
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "No."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   510
            TabIndex        =   144
            Top             =   390
            Width           =   540
         End
         Begin VB.Shape shpModel 
            FillColor       =   &H00808080&
            FillStyle       =   0  '단색
            Height          =   315
            Index           =   0
            Left            =   540
            Shape           =   5  '둥근 정사각형
            Top             =   1110
            Width           =   315
         End
         Begin VB.Label lblPLCBit 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "모델 변경"
            Height          =   180
            Index           =   4
            Left            =   990
            TabIndex        =   143
            Top             =   1200
            Width           =   780
         End
         Begin VB.Shape shpModel 
            FillColor       =   &H00808080&
            FillStyle       =   0  '단색
            Height          =   315
            Index           =   1
            Left            =   2130
            Shape           =   5  '둥근 정사각형
            Top             =   1110
            Width           =   315
         End
         Begin VB.Label lblPLCBit 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "변경 완료"
            Height          =   180
            Index           =   5
            Left            =   2580
            TabIndex        =   142
            Top             =   1200
            Width           =   780
         End
         Begin VB.Label lblChangedModel 
            Alignment       =   1  '오른쪽 맞춤
            Caption         =   "2013.02.15 12:12:12"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   570
            TabIndex        =   141
            Top             =   1650
            Width           =   2925
         End
      End
      Begin VB.CheckBox chkJudgement 
         Caption         =   "폭1"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   -74670
         TabIndex        =   138
         Top             =   570
         Width           =   1635
      End
      Begin VB.CheckBox chkJudgement 
         Caption         =   "폭2"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   -74670
         TabIndex        =   137
         Top             =   1005
         Width           =   1635
      End
      Begin VB.CheckBox chkJudgement 
         Caption         =   "높이1"
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   -74670
         TabIndex        =   136
         Top             =   1425
         Width           =   1635
      End
      Begin VB.CheckBox chkJudgement 
         Caption         =   "높이2"
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   -74670
         TabIndex        =   135
         Top             =   1860
         Width           =   1635
      End
      Begin VB.CheckBox chkJudgement 
         Caption         =   "NSD(상1)"
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   -72660
         TabIndex        =   134
         Top             =   570
         Width           =   1635
      End
      Begin VB.CheckBox chkJudgement 
         Caption         =   "NSD(상2)"
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   -72660
         TabIndex        =   133
         Top             =   1005
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CheckBox chkJudgement 
         Caption         =   "NSD(하1)"
         Enabled         =   0   'False
         Height          =   375
         Index           =   6
         Left            =   -72660
         TabIndex        =   132
         Top             =   1425
         Width           =   1635
      End
      Begin VB.CheckBox chkJudgement 
         Caption         =   "NSD(하2)"
         Enabled         =   0   'False
         Height          =   375
         Index           =   7
         Left            =   -72660
         TabIndex        =   131
         Top             =   1860
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CheckBox chkJudgement 
         Caption         =   "NSD(우1)"
         Enabled         =   0   'False
         Height          =   375
         Index           =   8
         Left            =   -72660
         TabIndex        =   130
         Top             =   2280
         Width           =   1635
      End
      Begin VB.CheckBox chkJudgement 
         Caption         =   "NSD(우2)"
         Enabled         =   0   'False
         Height          =   375
         Index           =   9
         Left            =   -72660
         TabIndex        =   129
         Top             =   2730
         Width           =   1635
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  '평면
         Caption         =   "PC >> PLC"
         ForeColor       =   &H80000008&
         Height          =   2070
         Index           =   1
         Left            =   -71400
         TabIndex        =   53
         Top             =   640
         Width           =   6165
         Begin VB.Shape shpVision2 
            FillColor       =   &H00808080&
            FillStyle       =   0  '단색
            Height          =   315
            Index           =   0
            Left            =   120
            Shape           =   5  '둥근 정사각형
            Top             =   330
            Width           =   315
         End
         Begin VB.Label lblVision2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   " 준비신호"
            Height          =   180
            Index           =   0
            Left            =   495
            TabIndex        =   69
            Top             =   375
            Width           =   780
         End
         Begin VB.Label lblVision2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   " 검사중"
            Height          =   180
            Index           =   1
            Left            =   495
            TabIndex        =   68
            Top             =   780
            Width           =   600
         End
         Begin VB.Shape shpVision2 
            FillColor       =   &H00808080&
            FillStyle       =   0  '단색
            Height          =   315
            Index           =   1
            Left            =   120
            Shape           =   5  '둥근 정사각형
            Top             =   735
            Width           =   315
         End
         Begin VB.Shape shpVision2 
            FillColor       =   &H00808080&
            FillStyle       =   0  '단색
            Height          =   315
            Index           =   2
            Left            =   120
            Shape           =   5  '둥근 정사각형
            Top             =   1140
            Width           =   315
         End
         Begin VB.Label lblVision2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   " -"
            Height          =   180
            Index           =   3
            Left            =   495
            TabIndex        =   67
            Top             =   1590
            Width           =   150
         End
         Begin VB.Shape shpVision2 
            FillColor       =   &H00808080&
            FillStyle       =   0  '단색
            Height          =   315
            Index           =   11
            Left            =   120
            Shape           =   5  '둥근 정사각형
            Top             =   1545
            Width           =   315
         End
         Begin VB.Label lblVision2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "1번 OK"
            Height          =   180
            Index           =   4
            Left            =   2340
            TabIndex        =   66
            Top             =   375
            Width           =   585
         End
         Begin VB.Shape shpVision2 
            FillColor       =   &H00808080&
            FillStyle       =   0  '단색
            Height          =   315
            Index           =   3
            Left            =   1965
            Shape           =   5  '둥근 정사각형
            Top             =   330
            Width           =   315
         End
         Begin VB.Label lblVision2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "1번 NG"
            Height          =   180
            Index           =   5
            Left            =   2340
            TabIndex        =   65
            Top             =   780
            Width           =   600
         End
         Begin VB.Shape shpVision2 
            FillColor       =   &H00808080&
            FillStyle       =   0  '단색
            Height          =   315
            Index           =   4
            Left            =   1965
            Shape           =   5  '둥근 정사각형
            Top             =   735
            Width           =   315
         End
         Begin VB.Label lblVision2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "2번 OK"
            Height          =   180
            Index           =   6
            Left            =   2340
            TabIndex        =   64
            Top             =   1185
            Width           =   585
         End
         Begin VB.Shape shpVision2 
            FillColor       =   &H00808080&
            FillStyle       =   0  '단색
            Height          =   315
            Index           =   5
            Left            =   1965
            Shape           =   5  '둥근 정사각형
            Top             =   1140
            Width           =   315
         End
         Begin VB.Label lblVision2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "2번 NG"
            Height          =   180
            Index           =   7
            Left            =   2340
            TabIndex        =   63
            Top             =   1590
            Width           =   600
         End
         Begin VB.Shape shpVision2 
            FillColor       =   &H00808080&
            FillStyle       =   0  '단색
            Height          =   315
            Index           =   6
            Left            =   1965
            Shape           =   5  '둥근 정사각형
            Top             =   1545
            Width           =   315
         End
         Begin VB.Label lblVision2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   " 검사종료"
            Height          =   180
            Index           =   2
            Left            =   495
            TabIndex        =   62
            Top             =   1185
            Width           =   780
         End
         Begin VB.Shape shpVision2 
            FillColor       =   &H00808080&
            FillStyle       =   0  '단색
            Height          =   315
            Index           =   7
            Left            =   4035
            Shape           =   5  '둥근 정사각형
            Top             =   330
            Width           =   315
         End
         Begin VB.Label lblVision2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "3번 OK"
            Height          =   180
            Index           =   8
            Left            =   4410
            TabIndex        =   61
            Top             =   375
            Width           =   585
         End
         Begin VB.Label lblVision2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "3번 NG"
            Height          =   180
            Index           =   9
            Left            =   4410
            TabIndex        =   60
            Top             =   780
            Width           =   600
         End
         Begin VB.Shape shpVision2 
            FillColor       =   &H00808080&
            FillStyle       =   0  '단색
            Height          =   315
            Index           =   8
            Left            =   4035
            Shape           =   5  '둥근 정사각형
            Top             =   735
            Width           =   315
         End
         Begin VB.Shape shpVision2 
            FillColor       =   &H00808080&
            FillStyle       =   0  '단색
            Height          =   315
            Index           =   9
            Left            =   4035
            Shape           =   5  '둥근 정사각형
            Top             =   1140
            Width           =   315
         End
         Begin VB.Label lblVision2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "4번 NG"
            Height          =   180
            Index           =   11
            Left            =   4410
            TabIndex        =   59
            Top             =   1590
            Width           =   600
         End
         Begin VB.Shape shpVision2 
            FillColor       =   &H00808080&
            FillStyle       =   0  '단색
            Height          =   315
            Index           =   10
            Left            =   4035
            Shape           =   5  '둥근 정사각형
            Top             =   1545
            Width           =   315
         End
         Begin VB.Label lblVision2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "-"
            Height          =   180
            Index           =   12
            Left            =   6630
            TabIndex        =   58
            Top             =   375
            Width           =   90
         End
         Begin VB.Shape shpVision2 
            FillColor       =   &H00808080&
            FillStyle       =   0  '단색
            Height          =   315
            Index           =   12
            Left            =   6255
            Shape           =   5  '둥근 정사각형
            Top             =   330
            Width           =   315
         End
         Begin VB.Label lblVision2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "-"
            Height          =   180
            Index           =   13
            Left            =   6630
            TabIndex        =   57
            Top             =   780
            Width           =   90
         End
         Begin VB.Shape shpVision2 
            FillColor       =   &H00808080&
            FillStyle       =   0  '단색
            Height          =   315
            Index           =   13
            Left            =   6255
            Shape           =   5  '둥근 정사각형
            Top             =   735
            Width           =   315
         End
         Begin VB.Label lblVision2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "-"
            Height          =   180
            Index           =   14
            Left            =   6630
            TabIndex        =   56
            Top             =   1185
            Width           =   90
         End
         Begin VB.Shape shpVision2 
            FillColor       =   &H00808080&
            FillStyle       =   0  '단색
            Height          =   315
            Index           =   14
            Left            =   6255
            Shape           =   5  '둥근 정사각형
            Top             =   1140
            Width           =   315
         End
         Begin VB.Label lblVision2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "-"
            Height          =   180
            Index           =   15
            Left            =   6630
            TabIndex        =   55
            Top             =   1590
            Width           =   90
         End
         Begin VB.Shape shpVision2 
            FillColor       =   &H00808080&
            FillStyle       =   0  '단색
            Height          =   315
            Index           =   15
            Left            =   6255
            Shape           =   5  '둥근 정사각형
            Top             =   1545
            Width           =   315
         End
         Begin VB.Label lblVision2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "4번 OK"
            Height          =   180
            Index           =   10
            Left            =   4410
            TabIndex        =   54
            Top             =   1185
            Width           =   585
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  '평면
         Caption         =   "PLC >> Vision"
         ForeColor       =   &H80000008&
         Height          =   2070
         Left            =   -74700
         TabIndex        =   48
         Top             =   640
         Width           =   3015
         Begin VB.Shape shpInput 
            BackColor       =   &H00E0E0E0&
            FillColor       =   &H00808080&
            FillStyle       =   0  '단색
            Height          =   315
            Index           =   0
            Left            =   120
            Shape           =   5  '둥근 정사각형
            Top             =   330
            Width           =   315
         End
         Begin VB.Label lblPLCBit 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   " 검사 트리거1"
            Height          =   180
            Index           =   0
            Left            =   495
            TabIndex        =   52
            Top             =   375
            Width           =   1110
         End
         Begin VB.Label lblPLCBit 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   " 검사 트리거2"
            Height          =   180
            Index           =   1
            Left            =   495
            TabIndex        =   51
            Top             =   780
            Width           =   1110
         End
         Begin VB.Shape shpInput 
            FillColor       =   &H00808080&
            FillStyle       =   0  '단색
            Height          =   315
            Index           =   1
            Left            =   120
            Shape           =   5  '둥근 정사각형
            Top             =   735
            Width           =   315
         End
         Begin VB.Shape shpInput 
            FillColor       =   &H00808080&
            FillStyle       =   0  '단색
            Height          =   315
            Index           =   2
            Left            =   120
            Shape           =   5  '둥근 정사각형
            Top             =   1140
            Width           =   315
         End
         Begin VB.Label lblPLCBit 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   " 검사 트리거4"
            Height          =   180
            Index           =   3
            Left            =   495
            TabIndex        =   50
            Top             =   1590
            Width           =   1110
         End
         Begin VB.Shape shpInput 
            FillColor       =   &H00808080&
            FillStyle       =   0  '단색
            Height          =   315
            Index           =   3
            Left            =   120
            Shape           =   5  '둥근 정사각형
            Top             =   1545
            Width           =   315
         End
         Begin VB.Label lblPLCBit 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   " 검사 트리거3"
            Height          =   180
            Index           =   2
            Left            =   495
            TabIndex        =   49
            Top             =   1185
            Width           =   1110
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H80000004&
         Height          =   2985
         Left            =   60
         TabIndex        =   25
         Top             =   360
         Width           =   15660
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   2685
            Left            =   90
            TabIndex        =   26
            Top             =   210
            Width           =   15000
            _ExtentX        =   26458
            _ExtentY        =   4736
            _Version        =   393216
            Cols            =   15
            RowHeightMin    =   400
            BackColorFixed  =   16761024
            ForeColorFixed  =   0
            BackColorBkg    =   -2147483633
            WordWrap        =   -1  'True
            FormatString    =   $"frmMain.frx":5AA3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin BHButton.BHImageButton btnSaveJudgement 
         Height          =   645
         Left            =   -68790
         TabIndex        =   128
         Top             =   510
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1138
         Caption         =   "저 장"
         CaptionChecked  =   "BHImageButton2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton btnApplyOffset 
         Height          =   735
         Left            =   -63930
         TabIndex        =   190
         Top             =   480
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         Caption         =   "적용"
         CaptionChecked  =   "BHImageButton2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton btnOffsetDefualt 
         Height          =   735
         Left            =   -63930
         TabIndex        =   191
         Top             =   1350
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1296
         Caption         =   "초기화"
         CaptionChecked  =   "BHImageButton2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin VB.Label lblOffset 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "폭1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -74730
         TabIndex        =   189
         Top             =   540
         Width           =   525
      End
      Begin VB.Label lblOffset 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "폭2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   -74730
         TabIndex        =   188
         Top             =   1125
         Width           =   525
      End
      Begin VB.Label lblOffset 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "높이1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -74730
         TabIndex        =   187
         Top             =   1725
         Width           =   855
      End
      Begin VB.Label lblOffset 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "높이2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   -74730
         TabIndex        =   186
         Top             =   2310
         Width           =   855
      End
      Begin VB.Label lblOffset 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "NSD(상1)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   -71550
         TabIndex        =   185
         Top             =   480
         Width           =   1500
      End
      Begin VB.Label lblOffset 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "NSD(상2)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   -71550
         TabIndex        =   184
         Top             =   1095
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label lblOffset 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "NSD(하1)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   -71550
         TabIndex        =   183
         Top             =   1695
         Width           =   1500
      End
      Begin VB.Label lblOffset 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "NSD(하2)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   -71550
         TabIndex        =   182
         Top             =   2310
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label lblOffset 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "NSD(우1)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   -67590
         TabIndex        =   181
         Top             =   510
         Width           =   1500
      End
      Begin VB.Label lblOffset 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "NSD(우2)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   -67590
         TabIndex        =   180
         Top             =   1110
         Width           =   1500
      End
   End
   Begin VB.TextBox txt_FocusStealer 
      Height          =   270
      Left            =   105
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   11685
      Width           =   345
   End
   Begin uEyeCamLib.uEyeCam uEyeCam1 
      Height          =   1005
      Index           =   1
      Left            =   9105
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
      _Version        =   65551
      _ExtentX        =   1720
      _ExtentY        =   1773
      _StockProps     =   1
      EnableEvents    =   -1  'True
      AutoSensorShutterMode=   0
      AutoSensorGainMode=   0
   End
   Begin uEyeCamLib.uEyeCam uEyeCam1 
      Height          =   1005
      Index           =   2
      Left            =   2715
      Top             =   7815
      Visible         =   0   'False
      Width           =   975
      _Version        =   65551
      _ExtentX        =   1720
      _ExtentY        =   1773
      _StockProps     =   1
      EnableEvents    =   -1  'True
      AutoSensorShutterMode=   0
      AutoSensorGainMode=   0
   End
   Begin uEyeCamLib.uEyeCam uEyeCam1 
      Height          =   1005
      Index           =   3
      Left            =   9105
      Top             =   7815
      Visible         =   0   'False
      Width           =   975
      _Version        =   65551
      _ExtentX        =   1720
      _ExtentY        =   1773
      _StockProps     =   1
      EnableEvents    =   -1  'True
      AutoSensorShutterMode=   0
      AutoSensorGainMode=   0
   End
   Begin VB.PictureBox picBase 
      Appearance      =   0  '평면
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   825
      Index           =   1
      Left            =   4560
      ScaleHeight     =   795
      ScaleWidth      =   8520
      TabIndex        =   31
      Top             =   0
      Width           =   8550
      Begin VB.Label lblInspecTime 
         Alignment       =   1  '오른쪽 맞춤
         BackColor       =   &H80000009&
         BackStyle       =   0  '투명
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   360
         Left            =   7680
         TabIndex        =   124
         Top             =   30
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblProgramTitle 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "터미널 CAP 측정 VISION"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   1845
         TabIndex        =   32
         Top             =   150
         Width           =   4845
      End
   End
   Begin VB.PictureBox picBase 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   840
      Index           =   2
      Left            =   13110
      ScaleHeight     =   810
      ScaleWidth      =   4215
      TabIndex        =   33
      Top             =   0
      Width           =   4245
      Begin VB.Label lblPLCConnect 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "PLC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3555
         TabIndex        =   149
         Top             =   263
         Width           =   480
      End
      Begin VB.Shape shpPLCSock 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   690
         Left            =   3450
         Shape           =   4  '둥근 사각형
         Top             =   60
         Width           =   690
      End
      Begin VB.Label lblMESServer 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "MES"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2790
         TabIndex        =   148
         Top             =   270
         Width           =   510
      End
      Begin VB.Shape shpMESSock 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   690
         Left            =   2700
         Shape           =   4  '둥근 사각형
         Top             =   60
         Width           =   690
      End
      Begin VB.Label lblMESNetDrive 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "N/D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2055
         TabIndex        =   147
         Top             =   270
         Width           =   510
      End
      Begin VB.Shape shpMESNetDrive 
         BackColor       =   &H0000FF00&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   690
         Left            =   1965
         Shape           =   4  '둥근 사각형
         Top             =   60
         Width           =   690
      End
      Begin VB.Label lblModelNameMain 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Inspection Model Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   105
         TabIndex        =   35
         Top             =   180
         Width           =   3990
      End
      Begin VB.Label aaa 
         Caption         =   "Label11"
         Height          =   375
         Left            =   4200
         TabIndex        =   34
         Top             =   240
         Width           =   1695
      End
   End
   Begin CognexDisplay.CogDisplay CogDisplay 
      Height          =   4365
      Index           =   0
      Left            =   0
      TabIndex        =   95
      Top             =   1320
      Width           =   6405
      _cx             =   11298
      _cy             =   7699
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
      AutoFitWithGraphics=   0   'False
      MaintainZoom    =   0   'False
   End
   Begin CognexDisplay.CogDisplay CogDisplay 
      Height          =   4365
      Index           =   1
      Left            =   6390
      TabIndex        =   96
      Top             =   1320
      Width           =   6405
      _cx             =   11298
      _cy             =   7699
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
      AutoFitWithGraphics=   0   'False
      MaintainZoom    =   0   'False
   End
   Begin CognexDisplay.CogDisplay CogDisplay 
      Height          =   4335
      Index           =   2
      Left            =   0
      TabIndex        =   97
      Top             =   6150
      Width           =   6405
      _cx             =   11298
      _cy             =   7646
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
      AutoFitWithGraphics=   0   'False
      MaintainZoom    =   0   'False
   End
   Begin CognexDisplay.CogDisplay CogDisplay 
      Height          =   4335
      Index           =   3
      Left            =   6390
      TabIndex        =   98
      Top             =   6150
      Width           =   6405
      _cx             =   11298
      _cy             =   7646
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
      AutoFitWithGraphics=   0   'False
      MaintainZoom    =   0   'False
   End
   Begin BHButton.BHImageButton BHB_EtcSetting 
      Height          =   540
      Left            =   17400
      TabIndex        =   127
      ToolTipText     =   "저장 - SPEC , 항목이름 변경 , 기능 설정이 저장됩니다. "
      Top             =   9900
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   953
      Caption         =   "환경설정"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin MSFlexGridLib.MSFlexGrid grdSpec 
      Height          =   4635
      Left            =   12840
      TabIndex        =   145
      Top             =   4140
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   8176
      _Version        =   393216
      Rows            =   11
      Cols            =   4
      RowHeightMin    =   400
      BackColorFixed  =   16761024
      ForeColorFixed  =   0
      BackColorBkg    =   -2147483633
      ScrollBars      =   0
      FormatString    =   "^                    항목|^        하한값   |^        기준값    |^       상한값  "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox chkResultImageSaving 
      BackColor       =   &H8000000E&
      Caption         =   "자동 검사시 결과파일 저장/미저장 선택"
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
      Left            =   12870
      TabIndex        =   200
      Top             =   9660
      Width           =   4245
   End
   Begin VB.Label lblAutoStop 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "자동검사중"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   20.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   518
      TabIndex        =   150
      ToolTipText     =   "자동검사 상태 와 설비 정지 상태를 나타냅니다."
      Top             =   10905
      Width           =   2115
   End
   Begin VB.Label Label11 
      Alignment       =   1  '오른쪽 맞춤
      BackColor       =   &H80000012&
      Caption         =   "ms"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   360
      Left            =   11025
      TabIndex        =   3
      Top             =   45
      Width           =   435
   End
   Begin VB.Label lblProgramName 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "물류"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   20.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   6180
      TabIndex        =   2
      Top             =   180
      Width           =   840
   End
   Begin VB.Label lblResults 
      Alignment       =   2  '가운데 맞춤
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   17168
      TabIndex        =   0
      ToolTipText     =   "검사 결과가 OK / NG 로 표시 됩니다."
      Top             =   10777
      Width           =   795
   End
   Begin VB.Shape shpAutoStop 
      BackColor       =   &H00008000&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   75
      Shape           =   4  '둥근 사각형
      Top             =   10725
      Width           =   3000
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000C&
      BackStyle       =   1  '투명하지 않음
      Height          =   105
      Left            =   -30
      Top             =   735
      Width           =   19200
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   1005
      Left            =   15
      Shape           =   4  '둥근 사각형
      Top             =   10605
      Width           =   3195
   End
   Begin VB.Shape ShpResult 
      BackColor       =   &H00008000&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   16095
      Shape           =   4  '둥근 사각형
      Top             =   10710
      Width           =   2940
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   960
      Index           =   1
      Left            =   15930
      Shape           =   4  '둥근 사각형
      Top             =   10605
      Width           =   3165
   End
   Begin VB.Label lbl_IDcodeCleaner 
      Caption         =   "Label26"
      Height          =   255
      Left            =   11400
      TabIndex        =   29
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mstr_PathVisionImg As String

Public Sub Counter(OKNG As String)

    If OKNG <> "Pass" Then
        lToTalCount = lToTalCount + 1
    End If
    
    Select Case OKNG
    Case "OK"
        lOKCount = lOKCount + 1
    Case "NG"
        lNGCount = lNGCount + 1
    End Select
    
    lblCountTotal.Caption = CStr(lToTalCount)
    lblCountOK.Caption = CStr(lOKCount)
    lblCountNG.Caption = CStr(lNGCount)
    
    Call SaveCount
    
End Sub
Private Sub ActQJ71E71TCP_OnDeviceStatus(ByVal szDevice As String, ByVal lData As Long, ByVal lReturnCode As Long)
Dim temp As String
Dim temp1 As String
Dim temp2 As String
temp = szDevice
temp1 = lData
temp2 = lReturnCode
End Sub

Private Sub BHB_EtcSetting_Click()
    Me.fraSpecName.Visible = True
End Sub

Private Sub BHBAutoRun_Click()
On Error GoTo err:
Dim bRet As Boolean

    mb_IfstopBtnClked = False
    
    If bAutoRunOn = False Then
    
        Me.shpAutoStop.BackColor = &H8000&
        Me.lblAutoStop.Caption = "자동검사중"
        
        Me.BHBLive.Enabled = False
        Me.BHBManualRun.Enabled = False
        Me.BHBModel.Enabled = False
        Me.BHBSetting.Enabled = False
        Me.BHBEnd.Enabled = False
        Me.BHBAutoRun.Enabled = False
        Me.BHBStop.Enabled = True
        bAutoRunOn = True
        
        Call MES_DATASEND_FUNC("EQ_STATE_EVENT", "AUTO", "")          '
        
        Call ClearMelsecResult
        Call SendSignalToMelsec(0, 1)
        
        If g_UseLightTimer = 1 Then
            PWM_LightAll True
            g_LightTimerCount = g_LightTimerInterval
        End If
        
        Call Terminal_AutoRun
        
        Call ClearMelsecResult
        Call SendSignalToMelsec(0, 0)
        
    End If
    

Exit Sub
err:
    MsgBox "PLC 와 통신이 끊어졌습니다. 잠시후 재접속 하십시오.", vbCritical, "PLC 통신 확인"
    frmMain.shpAutoStop.BackColor = &H40C0&
    frmMain.lblAutoStop.Caption = "정지상태"
    frmMain.BHBLive.Enabled = True
    frmMain.BHBManualRun.Enabled = True
    frmMain.BHBModel.Enabled = True
    frmMain.BHBSetting.Enabled = True
    frmMain.BHBEnd.Enabled = True
    frmMain.BHBAutoRun.Enabled = True
    frmMain.BHBStop.Enabled = True
    bAutoRunOn = False
    
End Sub

Private Sub BHBCountReset_Click()
    
    lToTalCount = 0
    lOKCount = 0
    lNGCount = 0
    
    Me.lblCountTotal.Caption = "0"
    Me.lblCountOK.Caption = "0"
    Me.lblCountNG.Caption = "0"
    
    lInspectionNum = 1
    
    frmMain.MSFlexGrid1.Rows = 1
    
    Call SaveCount
    Call LoadCount
    
End Sub

Private Sub BHBEnd_Click()

    If MsgBox("프로그램을 종료 하시겠습니까?", vbOKCancel, "프로그램 종료") = vbOK Then
        PWM_LightAll False
        End
    End If

End Sub

Private Sub BHBFuncCancel_Click()
    
    Me.fraSpecName.Visible = False
    
End Sub

Private Sub BHBFuncSAVE_Click()
On Error Resume Next
Dim i As Integer
    If MsgBox("변경 사항을 저장 하시겠습니까?", vbOKCancel, "저장") = vbOK Then
        Call SpecName_Save(sModelName)
        Call SpecName_Load(sModelName)
        Call SpecAllValue_Save(sModelName)
        Call FunctionValue_Save(sModelName)
        frmMain.MSFlexGrid1.Rows = 1
        
        For i = 0 To 9
            frmMain.txtSpecName(i).Text = sSpecName(i)
'            frmMain.chkSpecPass(i).Caption = sSpecName(i)
'            frmMain.lblSpecNameoff(i).Caption = sSpecName(i)
'            frmMain.lblResultName(i).Caption = sSpecName(i)
            frmSetting.lblResultName(i).Caption = sSpecName(i)
        Next i
        Me.fraSpecName.Visible = False
    End If

End Sub

Private Sub BHBLive_Click()

Dim i As Integer
Dim bRet As Boolean
  
    Me.shpAutoStop.BackColor = &HFFFF&
    Me.lblAutoStop.Caption = "동영상"
    Me.BHBLive.Enabled = False
    Me.BHBManualRun.Enabled = False
    Me.BHBModel.Enabled = False
    Me.BHBSetting.Enabled = False
    Me.BHBEnd.Enabled = False
    Me.BHBAutoRun.Enabled = False
    Me.BHBStop.Enabled = True
    bLiveOn = True
    
    '조명 켬
    Call PWM_LightAll(True, 100)
    
    Do
        DoEvents
        For i = 0 To kMaxCamera - 1
            Set g_CogImage(i) = IDS_AcquireCognex(uEyeCam1(i), CogDisplay(i))
            Sleep 1
        Next i
    Loop Until bLiveOn = False
    
    '조명 끔
    Call PWM_LightAll(False)
    
End Sub




Private Sub BHBManualRun_Click()
    On Error GoTo err
Dim i As Integer
Dim starttime As Long
Dim endtime As Long
Dim bRet As Boolean
Dim ImageFolderName As String
Dim ImageFolderName2 As String
Dim sDate As String
Dim stime As String
Dim sDataTemp As String
Dim sMesDate1 As String
Dim sMesTime1 As String
Dim tempstr As String
Dim sMesSendJPGPath As String

    sDate = Format(Date, "yy-mm-dd")
    stime = Format(Time, "hh-mm-ss")
    sMESDate = Format(Date, "YYYYMMDD")
    sMesTime = Format(Time, "HHMMSS")
    sDateTimeCheck = sMESDate & sMesTime
    
    ImageFolderName = "D:\Imagesave\" & sDate & "\" & sModelName & "\"
    Call Create_DIR(ImageFolderName)

    starttime = GetTickCount
    
    ' 조명 켬.
    Call PWM_LightAll(True, 100)

    sZigID = "NOID"
    
    '영상 획득
    For i = 0 To kMaxCamera - 1
        sIDCode(i) = "NOID"
        
        If frmMain.chkManualAcq.Value = 1 Then
            Set g_CogImage(i) = IDS_AcquireCognex(uEyeCam1(i), CogDisplay(i))
        Else
            Set g_CogImage(i) = LoadCogImage(App.Path & "\Model\" & sModelName & "\" & "Master" & CStr(i) & ".bmp")
            
            CogDisplayClear CogDisplay(i)
            Set CogDisplay(i).Image = g_CogImage(i)
        End If
    Next i
    
    ' 조명 끔.
    Call PWM_LightAll(False)
    g_LightTimerCount = 0
        
    '검사
    Call PreWelding_RunWidthHeight(frmMain.CogDisplay(0), frmMain.CogDisplay(1), frmMain.CogDisplay(2), frmMain.CogDisplay(3))
    Call PreWelding_RunNSD(frmMain.CogDisplay(0), frmMain.CogDisplay(1), frmMain.CogDisplay(2), frmMain.CogDisplay(3), frmMain.CogDisplay(g_CogBlobIndex))
    
    '판정
    Call PreWelding_Judgement
    
    Dim Judge As String
    
    If g_NGCount > 0 Then
        If bNGimageSave = True Then
            Call Create_DIR(ImageFolderName & "NG")
            For i = 0 To 3
                Call SaveCogImage(ImageFolderName & "NG" & "\" & sMESDate & "_" & sMesTime & "_" & sIDCode(0) & "_" & "CAM" & CStr(i + 1) & IIf(iImageFileMode = 1, ".bmp", ".jpg"), g_CogImage(i))
            Next i
        End If
        Judge = "NG"
    Else
        If bOKimageSave = True Then
            Call Create_DIR(ImageFolderName & "OK")
            For i = 0 To 3
                Call SaveCogImage(ImageFolderName & "OK" & "\" & sMESDate & "_" & sMesTime & "_" & sIDCode(0) & "_" & "CAM" & CStr(i + 1) & IIf(iImageFileMode = 1, ".bmp", ".jpg"), g_CogImage(i))
            Next i
        End If
        Judge = "OK"
    End If
    
    '결과 출력
    Call Terminal_WriteDataToGrid(0)
    
    '결과 표시
    frmMain.lblResults.Caption = Judge
    
    For i = 0 To 3
        Call lblResultWH_Change(i)
    Next i
    
    For i = 0 To 5
        Call lblResultNSD_Change(i)
    Next i
    
    If chkManualSave.Value = 1 Then
        DoEvents
        sMesSendJPGPath = "D:\MES\SEND\" & sIDCode(0) & "_" & sMESEquipCode & "_" & 1 & "_" & sDateTimeCheck & "_001" & ".JPG"
        Call SH_ScreenSave(sMesSendJPGPath, ImageFolderName & sIDCode(0) & "_" & sMESEquipCode & "_" & 1 & "_" & sDateTimeCheck & "_001" & ".JPG")

        'QCP 파일 백업
        sDataTemp = DJ_DataFileADD(0)
        Call DataFileSave(0, sDataTemp, ImageFolderName & sIDCode(0) & "_" & sMESEquipCode & "_" & 1 & "_" & sDateTimeCheck & ".QCP")     '저장되는 데이터 생성
    End If

    endtime = GetTickCount
    frmMain.lblInspecTime.Caption = CStr(endtime - starttime) & " ㎳"
    
    Exit Sub
err:
    MsgBox "BHBManualRun Error!!"
    
End Sub

Private Sub BHBMESMain_Click()

    frmMESMain.Show
    
End Sub

Private Sub BHBMESNetDriveConnect_Click()
    Call MES_NetDriveConnect
End Sub

Private Sub BHBMESServerResetM_Click()
    Call MES_ServerOpen
    ListBox_Append Time & "서버가 OPEN 되었습니다.", 1
End Sub

Private Sub BHBModel_Click()
If MsgBox("모델 관리창을 실행 하시겠습니까?", vbOKCancel, "모델 관리창 열기") = vbOK Then
    frmModelAuto.Top = -14500
    frmModelAuto.Show
    Do
        DoEvents
        frmModelAuto.Top = frmModelAuto.Top + 200
        
    Loop Until frmModelAuto.Top >= 480
End If

End Sub

Private Sub BHBPLCConnect_Click()
On Error GoTo err:
'
'    If iProName = 2 Then
'        Call QJ71E71DisConnect
'        Call QJ71E71Connect
'    Else
'        Me.WinsockPLC.Close
'        Me.lstPLCSocket.Clear
'        ListBox_Append Time & "접속을 종료합니다.", 0
'
'        Call Win_Connect(0)
'    End If
Exit Sub
err:
    ListBox_Append Time & "연결 시도 실패....", 0
End Sub

Private Sub BHBPLCConnectM_Click()
On Error GoTo err:

'    If iProName = 2 Then
'        Call QJ71E71DisConnect
'        Call QJ71E71Connect
'    Else
'        Me.WinsockPLC.Close
'        Me.lstPLCSocket.Clear
'        ListBox_Append Time & "접속을 종료합니다.", 0
'
'        Call Win_Connect(0)
'    End If
Exit Sub
err:
    ListBox_Append Time & "연결 시도 실패....", 0
End Sub

Private Sub BHBSetting_Click()

    If MsgBox("검사 설정창을 실행 하시겠습니까?", vbOKCancel, "검사 설정창 열기") = vbOK Then
        Call LoadCogTool(sModelName)
        Call Calibration_Load(sModelName)
        Call LoadCameraPosition(sModelName)
        Call LoadSystemData
        
        frmSetting.Show 1
        
        Call LoadCogTool(sModelName)
        Call Calibration_Load(sModelName)
        Call LoadCameraPosition(sModelName)
        Call LoadSystemData
    End If

End Sub

Private Sub BHBSocketSetSave_Click()
    Call SocketSET_Save
    Call SocketSET_Load
End Sub

Private Sub BHBStop_Click()
On Error GoTo err:

Dim i As Integer
Dim bRet As Boolean
    
    mb_IfstopBtnClked = True

    If bAutoRunOn = True Then
        If MsgBox("자동 검사 상태를 해지 하시겠습니까?", vbOKCancel, "자동 검사실행") = vbOK Then
            If m_Snd_Bit_1(outReadyVision) = 1 Then
                m_Snd_Bit_1(outReadyVision) = 0
                Call Write_Send_Word(addVisionInspect1, Make_Send_Word(addVisionInspect1, True))
                ListBox_Append Time & " Vision Ready OFF 신호 전송완료", 0
                Me.shpAutoStop.BackColor = &H40C0&
                Me.lblAutoStop.Caption = "정지상태"
                Me.BHBLive.Enabled = True
                Me.BHBManualRun.Enabled = True
                Me.BHBModel.Enabled = True
                Me.BHBSetting.Enabled = True
                Me.BHBEnd.Enabled = True
                Me.BHBAutoRun.Enabled = True
                Me.BHBStop.Enabled = True
                bAutoRunOn = False
                
            Else
                GoTo comerr:
            End If
                                   
            Call MES_DATASEND_FUNC("EQ_STATE_EVENT", "MANUAL", "")
        Else
            bAutoRunOn = True
        End If
    Else
        Me.shpAutoStop.BackColor = &H40C0&
        Me.lblAutoStop.Caption = "정지상태"
        Me.tmrMelsec = False
        bLiveOn = False
        Me.BHBLive.Enabled = True
        Me.BHBManualRun.Enabled = True
        Me.BHBModel.Enabled = True
        Me.BHBSetting.Enabled = True
        Me.BHBEnd.Enabled = True
        Me.BHBAutoRun.Enabled = True
        Me.BHBStop.Enabled = True
        bAutoRunOn = False
        
    End If
    '조명끄기
     Call LightControl(0, False)
    Call LightControl(1, False)
    Call LightControl(2, False)
    Call LightControl(3, False)
    
    g_LightTimerCount = 0
    
Exit Sub
err:
    Me.shpAutoStop.BackColor = &H40C0&
    Me.lblAutoStop.Caption = "정지상태"
    bLiveOn = False
    'bAutoRunOn = False
    Me.BHBLive.Enabled = True
    Me.BHBManualRun.Enabled = True
    Me.BHBModel.Enabled = True
    Me.BHBSetting.Enabled = True
    Me.BHBEnd.Enabled = True
    Me.BHBAutoRun.Enabled = True
    Me.BHBStop.Enabled = True
    bAutoRunOn = False
    
Exit Sub
comerr:
    ListBox_Append Time & " Vision Ready OFF 신호 전송실패", 0
    Me.shpAutoStop.BackColor = &H40C0&
    Me.lblAutoStop.Caption = "정지상태"
    Me.BHBLive.Enabled = True
    Me.BHBManualRun.Enabled = True
    Me.BHBModel.Enabled = True
    Me.BHBSetting.Enabled = True
    Me.BHBEnd.Enabled = True
    Me.BHBAutoRun.Enabled = True
    Me.BHBStop.Enabled = True
    bAutoRunOn = False
    
    Call PWM_LightAll(False)
    
End Sub


Private Sub btnApplyOffset_Click()

    Dim i As Integer
    
    If MsgBox("옵셋을 적용 하시겠습니까?", vbOKCancel + vbQuestion, "옵셋 저장") <> vbOK Then
        Exit Sub
    End If
    
    For i = 0 To 9
        dSpecOffset(i) = CDbl(txtOffset(i).Text)
        txtOffset(i).Text = Format(dSpecOffset(i), "#0.00")
    Next i
    
    Call SpecAllValue_Save(sModelName)

End Sub

Private Sub btnLoadSystemData_Click()

    Call LoadSystemData
    Call MelsecAddressLoad
    
End Sub

Private Sub btnOffsetDefualt_Click()

    If MsgBox("옵셋 초기화 하시겠습니까?", vbOKCancel + vbQuestion, "옵셋 초기화") <> vbOK Then
        Exit Sub
    End If

    For i = 0 To 9
        dSpecOffset(i) = 0#
        txtOffset(i).Text = Format(dSpecOffset(i), "#0.00")
    Next i
    
    Call SpecAllValue_Save(sModelName)
    
End Sub

Private Sub btnReloadSpec_Click()
    
    Call ReadDataFromPLC
    Call Terminal_SpecPrint
    
End Sub

Private Sub btnSaveJudgement_Click()

    Dim i As Integer
    
    For i = 0 To 9
        bSpecPass(i) = IIf(chkJudgement(i).Value = 1, False, True)
    Next i
    
    bSpecPass(5) = True
    bSpecPass(7) = True
    
    Call SpecAllValue_Save(sModelName)
    
End Sub



Private Sub cboROI_Click()

    g_RetryROI = cboROI.listIndex
    Call SaveRetryParameters(sModelName)
    
End Sub



Private Sub cboROIBase_Click()

    g_RetryBase = cboROIBase.listIndex
    Call SaveRetryParameters(sModelName)

End Sub

Private Sub chkAutoLightOnOff_Click()
    Dim inputString As String
    
    If chkAutoLightOnOff.Value = 1 Then
        inputString = InputBox("조명 꺼짐 시간(초) 설정", "시간설정")
        If inputString = "" Then
            chkAutoLightOnOff.Value = 0
            Exit Sub
        End If
        If IsNumeric(inputString) = False Then
            MsgBox "숫자를 입력하시오", vbCritical, "입력에러"
            Exit Sub
        End If
        g_LightTimerInterval = CLng(inputString)
        lblAutoLightInterval.Caption = Format(g_LightTimerInterval, "(0초)")
        g_LightTimerCount = g_LightTimerInterval
        g_UseLightTimer = 1
        tmrLight.Enabled = True
        
        PWM_LightAll True
        Debug.Print "[자동조명] 켜기"
    Else
        g_UseLightTimer = 0
        tmrLight.Enabled = False
        lblAutoLightInterval.Caption = ""
        PWM_LightAll False
        Debug.Print "[자동조명] 끄기"
    End If
    
    Call SaveAutoLightParameters
    
End Sub

Private Sub chkCamPass_Click()
    
        If Me.chkCamPass.Value = 1 Then
            bCamPass = True
        Else
            bCamPass = False
        End If

    
    
End Sub
Private Sub chkNGImageSave_Click()

        If Me.chkNGImageSave.Value = 1 Then
            bNGimageSave = True
        Else
            bNGimageSave = False
        End If

End Sub
Private Sub chkOKImageSave_Click()

        If Me.chkOKImageSave.Value = 1 Then
            bOKimageSave = True
        Else
            bOKimageSave = False
        End If

End Sub

Private Sub chkResultImageSaving_Click()

    g_SaveResultImage = chkResultImageSaving.Value
    
    Call SaveResultSaving(sModelName)

End Sub

Private Sub chkRetry_Click()

    If chkRetry.Value = 1 Then
        g_UseRetry = 1
        cboROIBase.Enabled = True
        cboROI.Enabled = True
    Else
        g_UseRetry = 0
        cboROIBase.Enabled = False
        cboROI.Enabled = False
    End If
    
    Call SaveRetryParameters(sModelName)
    
End Sub

Private Sub chkWriteDataSave_Click()

        If Me.chkWriteDataSave.Value = 1 Then
            bWriteDataSave = True
        Else
            bWriteDataSave = False
        End If

End Sub

Private Sub Command1_Click()
Dim tempstr As String
Dim sDate As String
Dim stime As String
Dim sMesdata As String

    tempstr = "d:\MES\SEND\" & sIDCode(0) & "_" & sMESEquipCode & "_" & "1_" & sDateTimeCheck & "_001" & ".JPG"
    Call SH_ScreenSave(tempstr)

End Sub

Private Sub Form_Load()

Dim i As Integer
Dim CameraInit As Integer
Dim lngReturnBoardCount     As Long
Dim ret As Long
Dim ret2 As Long

    Dim Color As ColorConstants
    Dim Message As String
    
    Dim bRet As Boolean

    Load frmSplash
    
    Call frmSplash.SetPos(0)
    Call frmSplash.SetText("Vision Program 을 시작합니다!!", vbBlack)
    Call Sleep(500)
    
    '조명 초기화
    Call frmSplash.SetPos(10)
    Call frmSplash.SetText("조명 초기화 작업 실행...", vbBlack)
    m_bLightExist = PWM_Init
    Call Sleep(500)
    
    If m_bLightExist = True Then
        Color = vbBlue
        Message = "조명 초기화 작업 성공!!"
    Else
        Color = vbRed
        Message = "조명 초기화 작업 실패!!"
    End If
    
    Call frmSplash.SetPos(10)
    Call frmSplash.SetText(Message, Color)
    Call Sleep(500)
    
    '자동조명
    Call LoadAutoLightParameters
    If g_UseLightTimer = 1 Then
        lblAutoLightInterval.Caption = Format(g_LightTimerInterval, "(0초)")
        PWM_LightAll True
        g_LightTimerCount = g_LightTimerInterval
        tmrLight.Enabled = True
    Else
        chkAutoLightOnOff.Value = 0
    End If
    
    ' 카메라 초기화
    '카메라 개수 설정
    Call frmSplash.SetPos(20)
    Call frmSplash.SetText("카메라 초기화 작업 실행...", vbBlack)
    Call Sleep(500)
    
    Call frmSplash.SetPos(30)
    
    g_bCameraInitialized = InitCamera()
    If g_bCameraInitialized = True Then
        Color = vbBlue
        Message = "카메라 초기화 작업 성공!!"
        Call frmSplash.SetText("카메라가 " & CStr(iCamNumber) & "개 연결.", Color)
        Call frmSplash.SetText("uEye IDS 5480CP-M : " & CStr(XRES) & " x " & CStr(YRES), Color)
    Else
        Color = vbRed
        Message = "카메라 초기화 작업 실패!!"
    End If
    Call frmSplash.SetText(Message, Color)
    Call Sleep(500)
    
    'Tool 초기화
    Call frmSplash.SetPos(40)
    Call frmSplash.SetText("Vision Tool 초기화 작업...", vbBlack)
    Call InitCogTool
    Call Sleep(500)
    
    
    Call frmSplash.SetPos(60)
    Call frmSplash.SetText("모델 목록 로딩...", vbBlack)
    Call ModelList_LOAD
    Call Sleep(500)
    
    Call frmSplash.SetPos(70)
    Call frmSplash.SetText("최근 모델 정보 로딩...", vbBlack)
    Call LastModelRead
    txtModelNumber.Text = CStr(g_ModelNumber)
    lblChangedModel.Caption = g_ModelChangedDate
    Call LoadSystemData
    Call LoadModel(sModelName)
    Call LoadCogTool(sModelName)
    Call Sleep(500)
    
    '재검사 설정값 불러오기
    Call LoadRetryParameters(sModelName)
    chkRetry.Value = g_UseRetry
    cboROIBase.listIndex = g_RetryBase
    cboROI.listIndex = g_RetryROI
    
    
    '멜섹초기화
    'Melsec 주소번지 읽어옴
    Call MelsecAddressLoad
    Call frmSplash.SetPos(80)
    Call frmSplash.SetText("Melsec Address 로드...", vbBlack)
    Call Sleep(500)
    
    Call frmSplash.SetPos(90)
    Call frmSplash.SetText("Melsec 초기화...", vbBlack)
    m_bMelsecConnected = MelsecSocketInit
    If m_bMelsecConnected = True Then
        Color = vbBlue
        Message = "Melsec 초기화 작업 완료!!"
    Else
        'frmMain.tmrMelsec.Enabled = False
        Color = vbRed
        Message = "Melsec 초기화 작업 실패!!"
    End If
    Call frmSplash.SetText(Message, Color)
    Call Sleep(100)
    
    Call frmSplash.SetPos(95)
    Call frmSplash.SetText("넷드라이브 연결 중...", vbBlack)
    
    Call DJ_MESRecipeIDCountLoad
    Call DJ_MESMowRecipeLoad
    Call DJ_MESFunctionLoad
    Call Sleep(200)
    
    iPgCount = 0
    
    For i = 0 To 1
        bWinsock(i) = False
    Next i
    lInspectionNum = 0
    iBlobToolCount = 0
    
    bTriggerOn = False '트리거On 초기화
    bDArrival = False

    Me.lblModelNameMain.Caption = sModelName
    
    
    '카메라 알람
    If g_bCameraInitialized = True Then
        Call WriteMelsec(frmMain.ActEasyIF, sMelsecAddrAlarmCamera, 0)
        'WriteLog "[PC→PLC] CAMERA ALARM 클리어"
    Else
        Call WriteMelsec(frmMain.ActEasyIF, sMelsecAddrAlarmCamera, 1)
        'WriteLog "[PC→PLC] CAMERA ALARM 출력"
    End If
        
    
    
    Call FormControlShow                     '파일처리된 Value 값들 로드 후 폼컨트롤에 value 값 써주기
    
    sSpecName(0) = "폭1"
    sSpecName(1) = "폭2"
    sSpecName(2) = "높이1"
    sSpecName(3) = "높이2"
    sSpecName(4) = "NSD1"
    sSpecName(5) = "NSD2"
    sSpecName(6) = "NSD3"
    sSpecName(7) = "NSD4"
    sSpecName(8) = "NSD5"
    sSpecName(9) = "NSD6"
    
    Call SocketSET_Load '소켓통신 관련 IP 와 Port 를 읽어옴
    Me.txtMESIP.Text = Me.WinsockMES.LocalIP
    'Call Terminal_InitGrid         'Data Grid 초기화
    MSFlexGrid1.Rows = 1
    Call LoadCount      '최근 Count 불러오기
    Me.shpAutoStop.BackColor = &H40C0&
    Me.lblAutoStop.Caption = "정지상태"
    bAutoRunOn = False

    For i = 0 To 1
        Call SH_HDDCheking(i)
    Next i
    Me.TmrMESSock.Enabled = True               'MES 상태 Timer 작동
    
    Call MES_NetDriveConnect
    Call MES_ServerOpen                  '폼 로드시 서버 오픈
    'Call MES_ImageFile_Send
    
    Call MES_DATASEND_FUNC("EQ_STATE_EVENT", "MANUAL", "")

    '스펙 초기화
    Call ReadDataFromPLC
    Call Terminal_SpecPrint
    
    Call frmSplash.SetPos(100)
    Call frmSplash.SetText("Vision Program 구동 준비가 완료!!", vbBlack)
    Call Sleep(200)
    
    Unload frmSplash
    
    iToolCount = 20
    
    '옵셋
    For i = 0 To 9
        txtOffset(i).Text = Format(dSpecOffset(i), "#0.00")
    Next i

End Sub

Private Sub Form_Unload(Cancel As Integer)                '폼 우측 상단 "X" 클릭시
Dim i As Integer
    If MsgBox("프로그램을 종료하시겠습니까?", vbOKCancel, "프로그램 종료") = vbOK Then
        PWM_LightAll False
        End
    Else
        Cancel = 1               '안해주면 취소 눌러도 프로그램 종료됨
    End If
End Sub

Private Sub grdSpec_Click()

    Dim Row As Integer
    Dim Col As Integer
    
    Dim Color As ColorConstants
    Dim i As Integer
    
    Row = grdSpec.Row
    Col = grdSpec.Col
    
    If bSpecPass(Row - 1) = False Then
        Color = vbYellow
        bSpecPass(Row - 1) = True
    Else
        Color = vbWhite
        bSpecPass(Row - 1) = False
    End If
        
    For i = 1 To grdSpec.Cols - 1
        grdSpec.Col = i
        grdSpec.CellBackColor = Color
    Next i
    
End Sub


Private Sub lblProgramTitle_DblClick()
    
    lblInspecTime.Visible = Not lblInspecTime.Visible
    
End Sub

Private Sub lblResultNSD_Change(Index As Integer)

   If CheckLabel(lblResultNSD(Index), dSpecOriMin(Index + 4), dSpecOriMax(Index + 4), vbBlack, vbBlack) = False And bSpecPass(Index + 4) = False Then
        shpResultNSD(Index).BackColor = vbRed
        lblResultTitleNSD(Index).ForeColor = vbRed
    Else
        shpResultNSD(Index).BackColor = vbWhite
        lblResultTitleNSD(Index).ForeColor = vbWhite
    End If
 
End Sub

Private Sub lblResults_Change()

    Select Case lblResults.Caption
    Case "OK"
        ShpResult.BackColor = &H8000&
    Case "NG"
        ShpResult.BackColor = vbRed
    Case "Pass"
        ShpResult.BackColor = vbYellow
    End Select

End Sub

Private Sub lblResultWH_Change(Index As Integer)

    If CheckLabel(lblResultWH(Index), dSpecOriMin(Index), dSpecOriMax(Index), vbBlack, vbBlack) = False And bSpecPass(Index) = False Then
        shpResultWH(Index).BackColor = vbRed
        lblResultTitleWH(Index).ForeColor = vbRed
    Else
        shpResultWH(Index).BackColor = vbWhite
        lblResultTitleWH(Index).ForeColor = vbWhite
    End If

End Sub

Private Sub Option1_Click(Index As Integer)          '이미지 저장 모드 설정
 
        Select Case Index
        Case 0            'bmp
            iImageFileMode = 1
        Case 1            'jpg
            iImageFileMode = 2
            
        End Select

    
End Sub

Private Sub tmrLight_Timer()
    
    If g_LightTimerCount = 0 Then
        '조명끄기
        PWM_LightAll False
        Debug.Print "[자동조명] 끄기"
    End If
    
    g_LightTimerCount = g_LightTimerCount - 1

End Sub

Private Sub TmrLogin_Timer()

    iTmrLogin = iTmrLogin + 1
    If iTmrLogin > 360 Then
        TmrLogin.Enabled = False
        frmTmpLogin.mb_CertificationOfLogin = False
    End If

End Sub

Private Sub tmrMelsec_Timer()
On Error GoTo ErrHandler

    Dim strDeviceList As String
    Dim nSize As Long
    Dim nData() As Long
    Dim nResult As Long
    Dim i As Integer
    
    DoEvents
    'strDeviceList = lMelsecAddrInput
    strDeviceList = GetAddressString(lMelsecAddrInput, 4)
    nSize = 4
    
    ReDim nData(nSize)
    
    nResult = ActEasyIF.ReadDeviceRandom(strDeviceList, nSize, nData(0))
    
    For i = 1 To 4
        If nData(i) > 0 Then
            nData(0) = nData(0) + (2 ^ i)
        End If
    Next i

    If nResult = 0 Then
        Call Read_Recieve_Bit(nData(0))
    End If
    
    Exit Sub
ErrHandler:
End Sub

Private Sub TmrMESSock_Timer()
    
    If Me.WinsockMES.State = 7 Then
        Me.shpMESSock.BackColor = vbGreen
        If sMelsecAddrAlarm <> "" Then
            Call WriteMelsec(frmMain.ActEasyIF, sMelsecAddrAlarm, 0)
        End If
    ElseIf Me.WinsockMES.State = 2 Then
        Call MES_ServerOpen
        Me.shpMESSock.BackColor = vbYellow
        If sMelsecAddrAlarm <> "" Then
            Call WriteMelsec(frmMain.ActEasyIF, sMelsecAddrAlarm, 1)
        End If
    Else
        Me.shpMESSock.BackColor = vbRed
        Call MES_ServerOpen
        If sMelsecAddrAlarm <> "" Then
            Call WriteMelsec(frmMain.ActEasyIF, sMelsecAddrAlarm, 1)
        End If
    End If

End Sub

Private Sub tmrMesTimeout_Timer()
          
    g_Timeout = g_Timeout + 1
    
    If g_Timeout > g_TimeoutRetry Then
        tmrMesTimeout.Enabled = False
        g_Timeout = 0
        Exit Sub
    End If
    
    Call MES_DATASEND_FUNC("TIMEOUT_EVENT", "", "", True)
    Call MES_DATASEND_FUNC(g_LastMesMsgId, g_LastMesMsgItem1, g_LastMesMsgItem2)
    'Call DJ_MESmsgLogSave(g_LastMesMsg)
    'Call MES_SendData(g_LastMesMsg)
    

    
End Sub

Private Sub TmrPLCSock_Timer()
'Dim temp As String

'If iProName = 2 Then
'    temp = Me.ActQJ71E71TCP.FreeDeviceStatus
'    If temp = 0 Then
'        Me.shpPLCSock.BackColor = vbGreen
'        Me.lblPLCConnect.Caption = "PLC Connect"
'    Else
'        Me.shpPLCSock.BackColor = vbRed
'       ' Call QJ71E71DisConnect '덕화
'       ' Call QJ71E71Connect
'        Me.lblPLCConnect.Caption = "PLC Disconnect"
'    End If
'Else
'    If Me.WinsockPLC.State = 7 Then
'        Me.shpPLCSock.BackColor = vbGreen
'        Me.lblPLCConnect.Caption = "PLC Connect"
'    Else
'        Me.shpPLCSock.BackColor = vbRed
'        If IsNetworkAlive(1) = 1 Then
'        '    Call Win_Disable(0)
'        '    Call Win_Connect(0)
'            Me.lblPLCConnect.Caption = "PLC Disconnect"
'        Else
'            Me.lblPLCConnect.Caption = "No Cable"
'        End If
'    End If
'End If
End Sub



Private Sub txtSpecMax_Click(Index As Integer)
    If frmTmpLogin.mb_CertificationOfLogin = False Then
        frmMain.txt_FocusStealer.SetFocus
        frmTmpLogin.Show
    End If
End Sub

Private Sub txtSpecMin_Click(Index As Integer)
    If frmTmpLogin.mb_CertificationOfLogin = False Then
        frmMain.txt_FocusStealer.SetFocus
        frmTmpLogin.Show
    End If
End Sub



Private Sub txtSpecName_Click(Index As Integer)
'    If frmTmpLogin.mb_CertificationOfLogin = False Then
'        frmMain.txt_FocusStealer.SetFocus
'        frmTmpLogin.Show
'    End If
End Sub

Private Sub txtSpecOffset_Click(Index As Integer)
'    If frmTmpLogin.mb_CertificationOfLogin = False Then
'        frmMain.txt_FocusStealer.SetFocus
'        frmMESLogin.Show
'    End If
End Sub

Private Sub txtSpecOri_Click(Index As Integer)
    If frmTmpLogin.mb_CertificationOfLogin = False Then
        frmMain.txt_FocusStealer.SetFocus
        frmTmpLogin.Show
    End If
End Sub

Private Sub txtUsedCapPerS_Change(Index As Integer)
    Me.PBDrive(Index).Value = Int(frmMain.txtUsedCapPerS(Index).Text)
    
    Me.lblDrivePer(Index).Caption = frmMain.txtUsedCapPerS(Index).Text & "%"
End Sub



Private Sub uEyeCam1_EventOnDeviceReconnected(Index As Integer)
On Error Resume Next

    Call CogDisplay(Index).StaticGraphics.Remove("ALARM")

    Call WriteMelsec(frmMain.ActEasyIF, sMelsecAddrAlarmCamera, 0)
    
End Sub

Private Sub uEyeCam1_EventOnDeviceRemoved(Index As Integer)
On Error Resume Next

    Dim Label As New CogGraphicLabel
    Label.Font.Bold = True
    Label.Font.Name = "Tahoma"
    Label.Font.size = 32
    Label.Alignment = cogGraphicLabelAlignmentTopLeft
    Label.Color = cogColorRed
    Label.SetXYText 20, 20, "카메라 연결 끊어짐!!!!!"
    
    Call CogDisplay(Index).StaticGraphics.Add(Label, "ALARM")
    
    Call WriteMelsec(frmMain.ActEasyIF, sMelsecAddrAlarmCamera, 1)
    
End Sub

Private Sub WinsockMES_Close()
    bWinsock(1) = False
    Me.shpMESSock.BackColor = vbRed
    Call MES_ServerOpen
End Sub

Private Sub WinsockMES_Connect()
    ListBox_Append Me.WinsockMES.RemoteHostIP & "컴퓨터에 " & Me.WinsockMES.RemotePort & "포트에 연결 되었습니다.", 1
    Me.lstMESSocket.Refresh
    bWinsock(1) = True
End Sub

Private Sub WinsockMES_ConnectionRequest(ByVal requestID As Long)
Dim temp(0 To 1) As String

    If WinsockMES.State <> sckClosed Then WinsockMES.Close
    WinsockMES.Accept requestID
    'temp(1) = "( " & WinsockMES.LocalIP & " )"
    bWinsock(1) = True
    temp(0) = Time & requestID & "   가 접속 되었습니다."
    ListBox_Append temp(0), 1
    Me.shpMESSock.BackColor = vbGreen
'    Me.lblMESServer.Caption = "MES Connect"
    
    Dim State As String
    
    If bAutoRunOn = True Then
        State = "AUTO"
    Else
        State = "MANUAL"
    End If
    
    Call MES_DATASEND_FUNC("EQ_STATE_EVENT", State, "")
    
End Sub

Private Sub WinsockMES_DataArrival(ByVal bytesTotal As Long)
On Error GoTo err:
    Dim sMesdata As String
    Dim sMesPC As String
    Dim iMesPC As Integer
    Dim temp_endLen As Integer
    Dim i As Integer
    
    frmMESMain.txtReciveMES.Text = ""
    WinsockMES.GetData sMesdata
    frmMESMain.txtReciveMES.Text = sMesdata
    
    Select Case DJSJ_XMLData_Find(1, "<MSG_ID>", "</MSG_ID>", sMesdata, temp_endLen)
        Case "LOGIN_REPLY"
            If DJSJ_XMLData_Find(1, "<RETURN_VALUE>", "</RETURN_VALUE>", sMesdata, temp_endLen) = 1 Then
                TmrLogin.Enabled = False
                mb_CertificationOfLogin = True
                frmMain.TmrLogin.Enabled = True
                frmTmpLogin.mb_CertificationOfLogin = True
                MsgBox "로그인 인증 성공", vbOKOnly, "로그인 정보"
                bMESReply = True
            ElseIf DJSJ_XMLData_Find(1, "<RETURN_VALUE>", "</RETURN_VALUE>", sMesdata, temp_endLen) = 0 Then
                MsgBox "로그인 인증 실패" & vbCrLf & "ID : " & sMesUserID & vbCrLf & "PW : " & sMesUserPass, vbCritical, "로그인 정보"
                bMESReply = True
            Else
                MsgBox "로그인 인증 요구 실패!! MES 로 확인하세요" & DJSJ_XMLData_Find(1, "<ERROR_MSG>", "</ERROR_MSG>", sMesdata, temp_endLen), vbCritical, "로그인 정보"
                bMESReply = True
            End If
        Case "DATE_EVENT"
            Dim temp As String
            Dim temp2 As String
            
            sMESDate = DJSJ_XMLData_Find(InStr(1, sMesdata, "<DATA>"), "<DATE>", "</DATE>", sMesdata, temp_endLen)
            temp = Format(Left(sMESDate, 10), "YYYY-MM-DD")
            temp2 = Format(Right(sMESDate, 8), "HH:MM:SS")
            Date = Format(Left(sMESDate, 10), "YYYY-MM-DD")
            Time = Format(Right(sMESDate, 8), "HH:MM:SS")
            
            iMesSysbyteR = DJSJ_XMLData_Find(1, "<SYSTEM_BYTES>", "</SYSTEM_BYTES>", sMesdata, temp_endLen)
            Call MES_DATASEND_FUNC("DATE_REPLY", "", CStr(iMesSysbyteR))
        Case "LINKTEST_EVENT"
            iMesSysbyteR = DJSJ_XMLData_Find(1, "<SYSTEM_BYTES>", "</SYSTEM_BYTES>", sMesdata, temp_endLen)
            Call MES_DATASEND_FUNC("LINKTEST_REPLY", "", CStr(iMesSysbyteR))
        Case "RECIPE_REPLY"
            
            If iRecipeIDcount < 11 Then
                sRecipeID(iRecipeIDcount) = DJSJ_XMLData_Find(InStr(1, sMesdata, "<DEFAULT>"), "<RECIPE_ID>", "</RECIPE_ID>", sMesdata, temp_endLen)
            
                sMesPC = DJSJ_XMLData_Find(InStr(1, sMesdata, "<DATA>"), "<PARAM_COUNT>", "</PARAM_COUNT>", sMesdata, temp_endLen)
                sParamCount(iRecipeIDcount) = CStr(sMesPC)
                iParamCount(iRecipeIDcount) = CInt(sMesPC)   '저장용
                iMesPC = CInt(sMesPC)
                temp_endLen = 1     '여기서 끝자리 반환을 초기화
                If iMesPC > 0 Then
                    For i = 1 To iMesPC
                        'sParamName_SV(iRecipeIDcount, i) = DJSJ_XMLData_Find(temp_endLen, "<PARAM_NAME>", "</PARAM_NAME>", sMesdata, temp_endLen)
                        sParamValue(iRecipeIDcount, i) = DJSJ_XMLData_Find(temp_endLen, "<PARAM_VALUE>", "</PARAM_VALUE>", sMesdata, temp_endLen)
                        sParamMinValue(iRecipeIDcount, i) = DJSJ_XMLData_Find(temp_endLen, "<PARAM_MINVALUE>", "</PARAM_MINVALUE>", sMesdata, temp_endLen)
                        sParamMaxValue(iRecipeIDcount, i) = DJSJ_XMLData_Find(temp_endLen, "<PARAM_MAXVALUE>", "</PARAM_MAXVALUE>", sMesdata, temp_endLen)
                    Next i
                End If
                Call MESRecipeRecieve
                Call DJ_MESRecipeSave(iRecipeIDcount)
                If iRecipeIDcount <= 10 Then
                    iRecipeIDcount = iRecipeIDcount + 1
                End If
                Call DJ_MESRecipeIDCountSave
            End If
            bMESReply = True
            
        Case "RECIPE_CHANGE_REPLY"
            If DJSJ_XMLData_Find(1, "<RETURN_VALUE>", "</RETURN_VALUE>", sMesdata, temp_endLen) = 1 Then
                MsgBox "RECIPE 변경 성공", vbOKOnly, "RECIPE CHANGE"
                Call MESRecipeChange_OK   '양승조추가해
                Call DJ_EquipSpecApply_NG '양승조추가해
                bMESReply = True
            ElseIf DJSJ_XMLData_Find(1, "<RETURN_VALUE>", "</RETURN_VALUE>", sMesdata, temp_endLen) = 0 Then
                MsgBox "RECIPE 변경 실패", vbCritical, "RECIPE CHANGE"
                Call MESRecipeChange_NG   '양승조추가해
                bMESReply = True
            Else
                MsgBox "RECIPE 변경 요구 실패!! MES 로 확인하세요" & DJSJ_XMLData_Find(1, "<ERROR_MSG>", "</ERROR_MSG>", sMesdata, temp_endLen), vbCritical, "RECIPE CHANGE"
                Call MESRecipeChange_NG   '양승조추가해
                bMESReply = True
            End If
        
        Case "RECIPE_SV_CHANGE_REPLY"
            If DJSJ_XMLData_Find(1, "<RETURN_VALUE>", "</RETURN_VALUE>", sMesdata, temp_endLen) = 1 Then
                MsgBox "RECIPE SV 변경 성공", vbOKOnly, "RECIPE SV CHANGE"
                Call DJ_EquipSpecApply_OK
                bMESReply = True
            ElseIf DJSJ_XMLData_Find(1, "<RETURN_VALUE>", "</RETURN_VALUE>", sMesdata, temp_endLen) = 0 Then
                MsgBox "RECIPE SV 변경 실패", vbCritical, "RECIPE SV CHANGE"
                Call DJ_EquipSpecApply_NG
                bMESReply = True
            Else
                MsgBox "RECIPE SV 변경 요구 실패!! MES 로 확인하세요" & DJSJ_XMLData_Find(1, "<ERROR_MSG>", "</ERROR_MSG>", sMesdata, temp_endLen), vbCritical, "RECIPE SV CHANGE"
                Call DJ_EquipSpecApply_NG
                bMESReply = True
            End If
        
    End Select
    
    If DJSJ_XMLData_Find(1, "<RETURN_VALUE>", "</RETURN_VALUE>", sMesdata, temp_endLen) = 1 Then
        frmMain.tmrMesTimeout.Enabled = False
        g_Timeout = 0
    End If
    
Exit Sub
err:
    ListBox_Append "Socket 통신중 Error 가 발생 하였습니다.", 1
End Sub

Private Sub WinsockMES_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error GoTo err

    ListBox_Append "WinsockMES_Error, Number ; " & Number & ", Description ; " & Description & ", Scode ; " & Scode, 1
    Call MES_ServerOpen
    
Exit Sub
'--------------------------------------------------------------------------------------------
err:
    ListBox_Append "Error, WinsockMES_Error " & err.Description, 1
Resume Next
End Sub
