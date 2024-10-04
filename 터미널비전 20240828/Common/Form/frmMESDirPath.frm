VERSION 5.00
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmMESDirPath 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  '단일 고정
   Caption         =   "Drive"
   ClientHeight    =   4815
   ClientLeft      =   8610
   ClientTop       =   3630
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   4215
   Begin BHButton.BHImageButton BHBCreateFD 
      Height          =   420
      Left            =   375
      TabIndex        =   3
      Top             =   4020
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   741
      Caption         =   "새폴더"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin VB.DirListBox Dir1 
      Height          =   3240
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   3495
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1350
      TabIndex        =   0
      Top             =   330
      Width           =   2325
   End
   Begin BHButton.BHImageButton BHBFDSave 
      Height          =   420
      Left            =   1560
      TabIndex        =   4
      Top             =   4020
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   741
      Caption         =   "경로 저장"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton BHImageButton1 
      Height          =   420
      Left            =   2745
      TabIndex        =   5
      Top             =   4020
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   741
      Caption         =   "닫 기"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
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
      Height          =   4560
      Left            =   135
      Top             =   135
      Width           =   3960
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "드라이브"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   420
      TabIndex        =   1
      Top             =   405
      Width           =   870
   End
End
Attribute VB_Name = "frmMESDirPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BHBCreateFD_Click()
Dim temp As String

    temp = InputBox("특수문자를 제외한 새로운 폴더이름을 입력해주세요", "폴더이름")
    
    Call Create_DIR(Dir1.Path & "\" & temp)
    Dir1.Refresh
    
End Sub

Private Sub BHBFDSave_Click()
    If bPathSelect(0) = True Then
        sMESFileSavePath = Dir1.Path
    ElseIf bPathSelect(1) = True Then
        sMESFileSendPath = Dir1.Path
    ElseIf bPathSelect(2) = True Then
        sMESLogSavePath = Dir1.Path
    End If
End Sub

Private Sub BHImageButton1_Click()
    Unload Me
End Sub

Private Sub Dir1_Change()
    If bPathSelect(0) = True Then
        frmMESFunction.lblFilePath.Caption = Dir1.Path
    ElseIf bPathSelect(1) = True Then
        frmMESFunction.lblSendPath.Caption = Dir1.Path
    ElseIf bPathSelect(2) = True Then
        frmMESFunction.lblLogPath.Caption = Dir1.Path
    End If

End Sub

Private Sub Drive1_Change()
    frmMESDirPath.Dir1.Path = frmMESDirPath.Drive1.Drive
    
End Sub
