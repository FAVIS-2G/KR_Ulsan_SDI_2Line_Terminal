VERSION 5.00
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmStart 
   Appearance      =   0  '평면
   BackColor       =   &H00000000&
   BorderStyle     =   0  '없음
   Caption         =   "Start"
   ClientHeight    =   4515
   ClientLeft      =   5190
   ClientTop       =   5340
   ClientWidth     =   8895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrProgress 
      Interval        =   100
      Left            =   8385
      Top             =   75
   End
   Begin VB.TextBox txtStartMsg 
      BackColor       =   &H00000080&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2160
      Left            =   315
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1050
      Width           =   4665
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  '없음
      Height          =   405
      Left            =   5205
      TabIndex        =   8
      Top             =   2700
      Width           =   3510
      Begin BHButton.BHImageButton BHBProgress 
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   9
         Top             =   150
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   344
         Caption         =   ""
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         BackColor       =   128
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton BHBProgress 
         Height          =   195
         Index           =   1
         Left            =   435
         TabIndex        =   10
         Top             =   150
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   344
         Caption         =   ""
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         BackColor       =   16777215
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton BHBProgress 
         Height          =   195
         Index           =   2
         Left            =   765
         TabIndex        =   11
         Top             =   150
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   344
         Caption         =   ""
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         BackColor       =   16777215
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton BHBProgress 
         Height          =   195
         Index           =   3
         Left            =   1095
         TabIndex        =   12
         Top             =   150
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   344
         Caption         =   ""
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         BackColor       =   16777215
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton BHBProgress 
         Height          =   195
         Index           =   4
         Left            =   1425
         TabIndex        =   13
         Top             =   150
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   344
         Caption         =   ""
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         BackColor       =   16777215
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton BHBProgress 
         Height          =   195
         Index           =   5
         Left            =   1755
         TabIndex        =   14
         Top             =   150
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   344
         Caption         =   ""
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         BackColor       =   16777215
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton BHBProgress 
         Height          =   195
         Index           =   6
         Left            =   2085
         TabIndex        =   15
         Top             =   150
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   344
         Caption         =   ""
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         BackColor       =   16777215
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton BHBProgress 
         Height          =   195
         Index           =   7
         Left            =   2415
         TabIndex        =   16
         Top             =   150
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   344
         Caption         =   ""
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         BackColor       =   16777215
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton BHBProgress 
         Height          =   195
         Index           =   8
         Left            =   2745
         TabIndex        =   17
         Top             =   150
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   344
         Caption         =   ""
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         BackColor       =   16777215
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton BHBProgress 
         Height          =   195
         Index           =   9
         Left            =   3075
         TabIndex        =   18
         Top             =   150
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   344
         Caption         =   ""
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         BackColor       =   16777215
         AlphaColor      =   16777215
         ImgOutLineSize  =   3
      End
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  '없음
      Height          =   135
      Left            =   8400
      TabIndex        =   19
      Top             =   2880
      Width           =   165
   End
   Begin VB.Label lblNowTime 
      BackColor       =   &H8000000E&
      Caption         =   "24:35:27"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   26.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   600
      Left            =   5865
      TabIndex        =   20
      Top             =   945
      Width           =   2235
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  '투명하지 않음
      BorderStyle     =   0  '투명
      Height          =   150
      Left            =   1725
      Top             =   3405
      Width           =   1740
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00000080&
      BackStyle       =   1  '투명하지 않음
      Height          =   2370
      Left            =   195
      Shape           =   4  '둥근 사각형
      Top             =   960
      Width           =   4905
   End
   Begin VB.Label lblName2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Vision"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   840
      Index           =   1
      Left            =   6015
      TabIndex        =   3
      Top             =   3495
      Width           =   2145
   End
   Begin VB.Label lblName2 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "Vision"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   840
      Index           =   0
      Left            =   6075
      TabIndex        =   1
      Top             =   3540
      Width           =   2145
   End
   Begin VB.Label lblName1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "Hybrid J/R "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   840
      Index           =   1
      Left            =   270
      TabIndex        =   2
      Top             =   -105
      Width           =   3795
   End
   Begin VB.Label lblName1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "Hybrid J/R "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   840
      Index           =   0
      Left            =   315
      TabIndex        =   0
      Top             =   -45
      Width           =   3795
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000080&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      Shape           =   4  '둥근 사각형
      Top             =   180
      Width           =   5820
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000080&
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Left            =   4980
      Shape           =   4  '둥근 사각형
      Top             =   3765
      Width           =   3645
   End
   Begin VB.Image Image1 
      Height          =   810
      Left            =   5295
      Stretch         =   -1  'True
      Top             =   1620
      Width           =   3090
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "DEJAY 2012 . 5 . 20 (Ver 1.1)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   5775
      TabIndex        =   4
      Top             =   3120
      Width           =   2430
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  '투명하지 않음
      BorderColor     =   &H00000080&
      BorderWidth     =   5
      FillColor       =   &H00000080&
      Height          =   2715
      Left            =   45
      Shape           =   4  '둥근 사각형
      Top             =   780
      Width           =   8805
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "VIS"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1005
      Left            =   1725
      TabIndex        =   7
      Top             =   3285
      Width           =   2010
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "FA"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1005
      Left            =   330
      TabIndex        =   6
      Top             =   3285
      Width           =   1590
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iPgCount As Integer
Private Sub Form_Load()

Dim ret As Long
Dim ret2 As Long
Dim i As Integer
    tmrProgress.Enabled = True
    frmStart.Show
    'Me.BackColor = vbYellow
    Me.Text1.SetFocus
    iRecipeIDcount = 1
    iNowRecipeID = 0
    frmStart.Refresh
    
    ret2 = ret2 Or WS_EX_LAYERED
    Call SetWindowLong(frmStart.hWnd, GWL_EXSTYLE, ret2)
    Call SetLayeredWindowAttributes(frmStart.hWnd, vbBlack, 80, LWA_COLORKEY)
    Me.txtStartMsg.Text = "Vision Program 을 시작합니다." & vbCrLf

    iImageFileMode = 1
    
    '조명 초기화
    m_bLightExist = InitLightIO
   ' 카메라 초기화
    
    'Melsec 주소번지 읽어옴
    Call MelsecAddressLoad
    
    '카메라 개수 설정
    Call ProgramSelect_Load
    Me.txtStartMsg.Text = Me.txtStartMsg.Text & "Camera 가 1개 연결 되었습니다." & vbCrLf
    Call Dlay_T(0.5)
    '카메라 초기화
    Me.txtStartMsg.Text = Me.txtStartMsg.Text & "Camera 를 초기화 합니다." & vbCrLf
    For i = 0 To 3
        Call AllocateNew(i, 3000, 3000)
        Call AllocateFav(i, 3000, 3000)
    Next i
    Me.txtStartMsg.Text = Me.txtStartMsg.Text & "Ueye IDS 1480 : 2560 x 1920" & vbCrLf
    'Call Dlay_T(0.1)
    Me.txtStartMsg.Text = Me.txtStartMsg.Text & "Vision Tool 을 초기화 합니다." & vbCrLf
    Call Tool_Init         'Tool 초기화
    Call Dlay_T(0.1)
    Me.txtStartMsg.Text = Me.txtStartMsg.Text & "전체 ModelList 를 불러옵니다." & vbCrLf
    Call ModelList_LOAD
    Me.txtStartMsg.Text = Me.txtStartMsg.Text & "최근 Model 을 불러옵니다." & vbCrLf
    Call LastModelRead
    'Call loa(sModelName)
    Call InitCogTool
    Me.txtStartMsg.Text = Me.txtStartMsg.Text & "최근 통신관련 설정을 불러옵니다." & vbCrLf
    '멜섹초기화
    m_bMelsecConnected = MelsecSocketInit
    If m_bMelsecConnected = True Then
        frmMain.tmrMelsec.Enabled = True
    Else
        frmMain.tmrMelsec.Enabled = False
    End If
    ''''''''''''
    Call Dlay_T(0.5)
    Me.txtStartMsg.Text = Me.txtStartMsg.Text & "Vision Program 구동 준비가 완료 되었습니다." & vbCrLf
    tmrProgress.Enabled = False
    Call Dlay_T(0.5)
    Call DJ_MESRecipeIDCountLoad
    Call DJ_MESMowRecipeLoad
    Call DJ_MESFunctionLoad
    Unload Me
    frmMain.Show
    iPgCount = 0
End Sub

Private Sub tmrProgress_Timer()
    frmStart.lblNowTime.Caption = Format(Time, "hh:mm:ss")
    Me.BHBProgress(iPgCount).BackColor = &H80&
    iPgCount = iPgCount + 1
    
    If iPgCount > 9 Then
        tmrProgress.Enabled = False
    End If
    
End Sub
