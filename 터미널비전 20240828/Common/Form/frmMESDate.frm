VERSION 5.00
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmMESDate 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  '없음
   Caption         =   "시간설정"
   ClientHeight    =   6690
   ClientLeft      =   390
   ClientTop       =   1740
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TmrDate 
      Interval        =   1000
      Left            =   10170
      Top             =   90
   End
   Begin VB.Frame fraDate 
      BackColor       =   &H8000000E&
      Caption         =   "설비 시간설정"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   2970
      TabIndex        =   1
      Top             =   1995
      Width           =   4905
      Begin VB.TextBox txtMesTime 
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "돋움"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1695
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   405
         Width           =   2805
      End
      Begin VB.TextBox txtPCTime 
         BorderStyle     =   0  '없음
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
         Left            =   1695
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   930
         Width           =   2805
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "MES시간 :"
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
         Height          =   270
         Left            =   330
         TabIndex        =   5
         Top             =   435
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "설비시간 :"
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
         Height          =   270
         Left            =   330
         TabIndex        =   4
         Top             =   960
         Width           =   1200
      End
   End
   Begin BHButton.BHImageButton BHBMESTimeSet 
      Height          =   840
      Left            =   2985
      TabIndex        =   0
      Top             =   3705
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   1482
      Caption         =   "시간 동기화"
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
   Begin BHButton.BHImageButton BHBTimeCancel 
      Height          =   840
      Left            =   5475
      TabIndex        =   6
      Top             =   3705
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   1482
      Caption         =   "닫 기"
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000080&
      BorderWidth     =   18
      Height          =   6495
      Left            =   135
      Top             =   90
      Width           =   10470
   End
End
Attribute VB_Name = "frmMESDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BHBTimeCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.TmrDate.Enabled = False
    frmMESDate.txtMesTime.Text = Format(Time, "HH:MM:SS")
    frmMESDate.txtPCTime.Text = Format(Time, "HH:MM:SS")
End Sub

Private Sub TmrDate_Timer()
    iTmrDate = iTmrDate + 1
    If iTmrDate = 3 Then
        If bMESReply = False Then
            MsgBox "MES로 부터 응답이 없습니다.", vbCritical, "타임아웃 에러"
            TmrDate.Enabled = False
        Else
            TmrDate.Enabled = False
        End If
    End If
End Sub
