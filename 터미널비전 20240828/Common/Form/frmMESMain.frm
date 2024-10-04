VERSION 5.00
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmMESMain 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  '단일 고정
   Caption         =   "MES MAIN"
   ClientHeight    =   14790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   14790
   ScaleWidth      =   11370
   Begin VB.Frame fraMESBHBSelect 
      BackColor       =   &H8000000E&
      Caption         =   "정보"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7065
      Left            =   195
      TabIndex        =   8
      Top             =   990
      Width           =   11025
      Begin VB.PictureBox picSection 
         Appearance      =   0  '평면
         BackColor       =   &H8000000E&
         BorderStyle     =   0  '없음
         ForeColor       =   &H80000008&
         Height          =   6750
         Left            =   90
         ScaleHeight     =   6750
         ScaleWidth      =   10890
         TabIndex        =   9
         Top             =   225
         Width           =   10890
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "Log"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6465
      Left            =   195
      TabIndex        =   0
      Top             =   8190
      Width           =   11010
      Begin VB.TextBox txtReciveMES 
         Height          =   5715
         Left            =   5580
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   630
         Width           =   5325
      End
      Begin VB.TextBox txtSendMES 
         Height          =   5715
         Left            =   105
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   630
         Width           =   5325
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  '투명
         Caption         =   "< 받기 >"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   1
         Left            =   5610
         TabIndex        =   4
         Top             =   345
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  '투명
         Caption         =   "< 보내기 >"
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   3
         Top             =   345
         Width           =   1035
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   330
         Left            =   105
         Shape           =   4  '둥근 사각형
         Top             =   285
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000080&
         BackStyle       =   1  '투명하지 않음
         BorderStyle     =   0  '투명
         Height          =   330
         Left            =   5580
         Shape           =   4  '둥근 사각형
         Top             =   285
         Width           =   900
      End
   End
   Begin BHButton.BHImageButton BHBMESLogin 
      Height          =   840
      Left            =   150
      TabIndex        =   5
      Top             =   75
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   1482
      Caption         =   "로그인"
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
   Begin BHButton.BHImageButton BHBDateSet 
      Height          =   840
      Left            =   2580
      TabIndex        =   6
      Top             =   75
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   1482
      Caption         =   "시간설정"
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
   Begin BHButton.BHImageButton BHBRecipeSet 
      Height          =   840
      Left            =   5010
      TabIndex        =   7
      Top             =   75
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   1482
      Caption         =   "Recipe"
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
   Begin BHButton.BHImageButton BHBMESFunc 
      Height          =   840
      Left            =   7440
      TabIndex        =   10
      Top             =   75
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   1482
      Caption         =   "환경설정"
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
   Begin BHButton.BHImageButton BHBMESUnload 
      Height          =   840
      Left            =   9870
      TabIndex        =   11
      Top             =   75
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1482
      Caption         =   "닫기"
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
Attribute VB_Name = "frmMESMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BHBDateSet_Click()
    Unload frmMESDate
    Unload frmMESRecipe
    Unload frmMESLogin
    Unload frmMESRecipePM
    Unload frmMESFunction
    Me.fraMESBHBSelect.Caption = Me.BHBDateSet.Caption
    Call ChangeViewSection(frmMESDate)
End Sub

Private Sub BHBMESFunc_Click()
    Unload frmMESDate
    Unload frmMESRecipe
    Unload frmMESLogin
    Unload frmMESRecipePM
    Unload frmMESFunction
    Me.fraMESBHBSelect.Caption = Me.BHBMESFunc.Caption
    Call ChangeViewSection(frmMESFunction)
End Sub

Private Sub BHBMESLogin_Click()
    Unload frmMESDate
    Unload frmMESRecipe
    Unload frmMESLogin
    Unload frmMESRecipePM
    Unload frmMESFunction
    Me.fraMESBHBSelect.Caption = Me.BHBMESLogin.Caption
    Call ChangeViewSection(frmMESLogin)
End Sub

Private Sub BHBMESUnload_Click()
    Unload Me
End Sub

Private Sub BHBRecipeSet_Click()
    Unload frmMESDate
    Unload frmMESRecipe
    Unload frmMESLogin
    Unload frmMESRecipePM
    Unload frmMESFunction
    Me.fraMESBHBSelect.Caption = Me.BHBRecipeSet.Caption
    Call ChangeViewSection(frmMESRecipe)
End Sub
