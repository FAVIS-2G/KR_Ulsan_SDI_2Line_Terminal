VERSION 5.00
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmModelAuto 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '���� ����
   Caption         =   "MODEL CHANGE"
   ClientHeight    =   13875
   ClientLeft      =   75
   ClientTop       =   1305
   ClientWidth     =   19125
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13875
   ScaleWidth      =   19125
   StartUpPosition =   1  '������ ���
   Begin BHButton.BHImageButton cmdModel 
      Height          =   690
      Index           =   0
      Left            =   5430
      TabIndex        =   215
      Top             =   12780
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   1217
      Caption         =   "SAVE"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      ForeColor       =   -2147483634
      BackColor       =   128
      AlphaColor      =   128
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton cmdModelRoom 
      Height          =   390
      Index           =   0
      Left            =   15870
      TabIndex        =   14
      Top             =   255
      Visible         =   0   'False
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   688
      Caption         =   "BHImageButton1"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   11745
      Left            =   615
      TabIndex        =   0
      Top             =   870
      Width           =   17835
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   1
         Left            =   690
         TabIndex        =   115
         Top             =   825
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   2
         Left            =   690
         TabIndex        =   116
         Top             =   1365
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   3
         Left            =   690
         TabIndex        =   117
         Top             =   1905
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   4
         Left            =   690
         TabIndex        =   118
         Top             =   2460
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   5
         Left            =   690
         TabIndex        =   119
         Top             =   2985
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   6
         Left            =   690
         TabIndex        =   120
         Top             =   3525
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   7
         Left            =   690
         TabIndex        =   121
         Top             =   4065
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   8
         Left            =   690
         TabIndex        =   122
         Top             =   4605
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   9
         Left            =   690
         TabIndex        =   123
         Top             =   5145
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   10
         Left            =   690
         TabIndex        =   124
         Top             =   5685
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   11
         Left            =   690
         TabIndex        =   125
         Top             =   6225
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   12
         Left            =   690
         TabIndex        =   126
         Top             =   6765
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   13
         Left            =   690
         TabIndex        =   127
         Top             =   7305
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   14
         Left            =   690
         TabIndex        =   128
         Top             =   7845
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   15
         Left            =   690
         TabIndex        =   129
         Top             =   8385
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   16
         Left            =   690
         TabIndex        =   130
         Top             =   8925
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   17
         Left            =   690
         TabIndex        =   131
         Top             =   9465
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   18
         Left            =   690
         TabIndex        =   132
         Top             =   10005
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   19
         Left            =   690
         TabIndex        =   133
         Top             =   10545
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   20
         Left            =   690
         TabIndex        =   134
         Top             =   11085
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   21
         Left            =   4245
         TabIndex        =   135
         Top             =   825
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   22
         Left            =   4245
         TabIndex        =   136
         Top             =   1365
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   23
         Left            =   4245
         TabIndex        =   137
         Top             =   1905
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   24
         Left            =   4245
         TabIndex        =   138
         Top             =   2445
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   25
         Left            =   4245
         TabIndex        =   139
         Top             =   2985
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   26
         Left            =   4245
         TabIndex        =   140
         Top             =   3525
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   27
         Left            =   4245
         TabIndex        =   141
         Top             =   4050
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   28
         Left            =   4245
         TabIndex        =   142
         Top             =   4605
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   29
         Left            =   4245
         TabIndex        =   143
         Top             =   5145
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   30
         Left            =   4245
         TabIndex        =   144
         Top             =   5685
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   31
         Left            =   4245
         TabIndex        =   145
         Top             =   6225
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   32
         Left            =   4245
         TabIndex        =   146
         Top             =   6765
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   33
         Left            =   4245
         TabIndex        =   147
         Top             =   7305
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   34
         Left            =   4245
         TabIndex        =   148
         Top             =   7845
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   35
         Left            =   4245
         TabIndex        =   149
         Top             =   8385
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   36
         Left            =   4245
         TabIndex        =   150
         Top             =   8925
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   37
         Left            =   4245
         TabIndex        =   151
         Top             =   9465
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   38
         Left            =   4245
         TabIndex        =   152
         Top             =   10005
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   39
         Left            =   4245
         TabIndex        =   153
         Top             =   10545
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   40
         Left            =   4245
         TabIndex        =   154
         Top             =   11085
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   41
         Left            =   7800
         TabIndex        =   155
         Top             =   825
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   42
         Left            =   7800
         TabIndex        =   156
         Top             =   1365
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   43
         Left            =   7800
         TabIndex        =   157
         Top             =   1905
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   44
         Left            =   7800
         TabIndex        =   158
         Top             =   2445
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   45
         Left            =   7800
         TabIndex        =   159
         Top             =   2985
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   46
         Left            =   7800
         TabIndex        =   160
         Top             =   3525
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   47
         Left            =   7800
         TabIndex        =   161
         Top             =   4050
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   48
         Left            =   7800
         TabIndex        =   162
         Top             =   4605
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   49
         Left            =   7800
         TabIndex        =   163
         Top             =   5145
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   50
         Left            =   7800
         TabIndex        =   164
         Top             =   5685
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   51
         Left            =   7800
         TabIndex        =   165
         Top             =   6225
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   52
         Left            =   7800
         TabIndex        =   166
         Top             =   6765
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   53
         Left            =   7800
         TabIndex        =   167
         Top             =   7305
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   54
         Left            =   7800
         TabIndex        =   168
         Top             =   7845
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   55
         Left            =   7800
         TabIndex        =   169
         Top             =   8385
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   56
         Left            =   7800
         TabIndex        =   170
         Top             =   8925
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   57
         Left            =   7800
         TabIndex        =   171
         Top             =   9465
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   58
         Left            =   7800
         TabIndex        =   172
         Top             =   10005
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   59
         Left            =   7800
         TabIndex        =   173
         Top             =   10545
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   60
         Left            =   7815
         TabIndex        =   174
         Top             =   11085
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   61
         Left            =   11355
         TabIndex        =   175
         Top             =   825
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   62
         Left            =   11355
         TabIndex        =   176
         Top             =   1365
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   63
         Left            =   11355
         TabIndex        =   177
         Top             =   1905
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   64
         Left            =   11355
         TabIndex        =   178
         Top             =   2460
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   65
         Left            =   11355
         TabIndex        =   179
         Top             =   2985
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   66
         Left            =   11355
         TabIndex        =   180
         Top             =   3525
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   67
         Left            =   11355
         TabIndex        =   181
         Top             =   4065
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   68
         Left            =   11355
         TabIndex        =   182
         Top             =   4605
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   69
         Left            =   11355
         TabIndex        =   183
         Top             =   5145
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   70
         Left            =   11355
         TabIndex        =   184
         Top             =   5685
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   71
         Left            =   11355
         TabIndex        =   185
         Top             =   6225
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   72
         Left            =   11355
         TabIndex        =   186
         Top             =   6765
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   73
         Left            =   11355
         TabIndex        =   187
         Top             =   7305
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   74
         Left            =   11355
         TabIndex        =   188
         Top             =   7845
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   75
         Left            =   11355
         TabIndex        =   189
         Top             =   8385
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   76
         Left            =   11355
         TabIndex        =   190
         Top             =   8925
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   77
         Left            =   11355
         TabIndex        =   191
         Top             =   9465
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   78
         Left            =   11355
         TabIndex        =   192
         Top             =   10005
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   79
         Left            =   11355
         TabIndex        =   193
         Top             =   10545
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   80
         Left            =   11355
         TabIndex        =   194
         Top             =   11085
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   81
         Left            =   14910
         TabIndex        =   195
         Top             =   825
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   82
         Left            =   14910
         TabIndex        =   196
         Top             =   1365
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   83
         Left            =   14910
         TabIndex        =   197
         Top             =   1905
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   84
         Left            =   14910
         TabIndex        =   198
         Top             =   2460
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   85
         Left            =   14910
         TabIndex        =   199
         Top             =   2985
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   86
         Left            =   14910
         TabIndex        =   200
         Top             =   3525
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   87
         Left            =   14910
         TabIndex        =   201
         Top             =   4065
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   88
         Left            =   14910
         TabIndex        =   202
         Top             =   4605
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   89
         Left            =   14910
         TabIndex        =   203
         Top             =   5145
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   90
         Left            =   14910
         TabIndex        =   204
         Top             =   5685
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   91
         Left            =   14910
         TabIndex        =   205
         Top             =   6225
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   92
         Left            =   14910
         TabIndex        =   206
         Top             =   6765
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   93
         Left            =   14910
         TabIndex        =   207
         Top             =   7305
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   94
         Left            =   14910
         TabIndex        =   208
         Top             =   7845
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   95
         Left            =   14910
         TabIndex        =   209
         Top             =   8385
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   96
         Left            =   14910
         TabIndex        =   210
         Top             =   8925
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   97
         Left            =   14910
         TabIndex        =   211
         Top             =   9465
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   98
         Left            =   14910
         TabIndex        =   212
         Top             =   10005
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   99
         Left            =   14910
         TabIndex        =   213
         Top             =   10545
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin BHButton.BHImageButton cmdModelRoom 
         Height          =   390
         Index           =   100
         Left            =   14910
         TabIndex        =   214
         Top             =   11085
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   688
         Caption         =   "Empty"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonAttrib    =   2
         CheckOption     =   1
         ForeColor       =   128
         BackColor       =   65280
         ImgOutLineSize  =   3
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   75
         TabIndex        =   114
         Top             =   1362
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   4
         Left            =   75
         TabIndex        =   113
         Top             =   1905
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   5
         Left            =   75
         TabIndex        =   112
         Top             =   2445
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   6
         Left            =   75
         TabIndex        =   111
         Top             =   2985
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   7
         Left            =   75
         TabIndex        =   110
         Top             =   3525
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   8
         Left            =   75
         TabIndex        =   109
         Top             =   4065
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   9
         Left            =   75
         TabIndex        =   108
         Top             =   4605
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   10
         Left            =   75
         TabIndex        =   107
         Top             =   5145
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   11
         Left            =   75
         TabIndex        =   106
         Top             =   5685
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   12
         Left            =   75
         TabIndex        =   105
         Top             =   6225
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   13
         Left            =   75
         TabIndex        =   104
         Top             =   6765
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   14
         Left            =   75
         TabIndex        =   103
         Top             =   7305
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   15
         Left            =   75
         TabIndex        =   102
         Top             =   7845
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   16
         Left            =   75
         TabIndex        =   101
         Top             =   8385
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   17
         Left            =   75
         TabIndex        =   100
         Top             =   8925
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   18
         Left            =   75
         TabIndex        =   99
         Top             =   9465
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   19
         Left            =   75
         TabIndex        =   98
         Top             =   10005
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   20
         Left            =   75
         TabIndex        =   97
         Top             =   10545
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   21
         Left            =   75
         TabIndex        =   96
         Top             =   11085
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   3
         Left            =   75
         TabIndex        =   95
         Top             =   825
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "22"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   30
         Left            =   3645
         TabIndex        =   94
         Top             =   1365
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "23"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   31
         Left            =   3645
         TabIndex        =   93
         Top             =   1905
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "24"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   32
         Left            =   3645
         TabIndex        =   92
         Top             =   2445
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   33
         Left            =   3645
         TabIndex        =   91
         Top             =   2985
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "26"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   34
         Left            =   3645
         TabIndex        =   90
         Top             =   3525
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "27"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   35
         Left            =   3645
         TabIndex        =   89
         Top             =   4065
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "28"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   36
         Left            =   3645
         TabIndex        =   88
         Top             =   4605
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "29"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   37
         Left            =   3645
         TabIndex        =   87
         Top             =   5145
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "30"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   38
         Left            =   3645
         TabIndex        =   86
         Top             =   5685
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   39
         Left            =   3645
         TabIndex        =   85
         Top             =   6225
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "32"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   40
         Left            =   3645
         TabIndex        =   84
         Top             =   6765
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "33"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   41
         Left            =   3645
         TabIndex        =   83
         Top             =   7305
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "34"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   42
         Left            =   3645
         TabIndex        =   82
         Top             =   7845
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "35"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   43
         Left            =   3645
         TabIndex        =   81
         Top             =   8385
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "36"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   44
         Left            =   3645
         TabIndex        =   80
         Top             =   8925
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "37"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   45
         Left            =   3645
         TabIndex        =   79
         Top             =   9465
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "38"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   46
         Left            =   3645
         TabIndex        =   78
         Top             =   10005
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "39"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   47
         Left            =   3645
         TabIndex        =   77
         Top             =   10545
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "40"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   48
         Left            =   3645
         TabIndex        =   76
         Top             =   11085
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   49
         Left            =   3645
         TabIndex        =   75
         Top             =   825
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "42"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   50
         Left            =   7200
         TabIndex        =   74
         Top             =   1365
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "43"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   51
         Left            =   7200
         TabIndex        =   73
         Top             =   1905
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "44"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   52
         Left            =   7200
         TabIndex        =   72
         Top             =   2445
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "45"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   53
         Left            =   7200
         TabIndex        =   71
         Top             =   2985
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "46"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   54
         Left            =   7200
         TabIndex        =   70
         Top             =   3525
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "47"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   55
         Left            =   7200
         TabIndex        =   69
         Top             =   4065
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "48"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   56
         Left            =   7200
         TabIndex        =   68
         Top             =   4605
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "49"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   57
         Left            =   7200
         TabIndex        =   67
         Top             =   5145
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   58
         Left            =   7200
         TabIndex        =   66
         Top             =   5685
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "51"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   59
         Left            =   7200
         TabIndex        =   65
         Top             =   6225
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "52"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   60
         Left            =   7200
         TabIndex        =   64
         Top             =   6765
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "53"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   61
         Left            =   7200
         TabIndex        =   63
         Top             =   7305
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "54"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   62
         Left            =   7200
         TabIndex        =   62
         Top             =   7845
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "55"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   63
         Left            =   7200
         TabIndex        =   61
         Top             =   8385
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "56"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   64
         Left            =   7200
         TabIndex        =   60
         Top             =   8925
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "57"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   65
         Left            =   7200
         TabIndex        =   59
         Top             =   9465
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "58"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   66
         Left            =   7200
         TabIndex        =   58
         Top             =   10005
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "59"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   67
         Left            =   7200
         TabIndex        =   57
         Top             =   10545
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "60"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   68
         Left            =   7200
         TabIndex        =   56
         Top             =   11085
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "41"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   69
         Left            =   7200
         TabIndex        =   55
         Top             =   825
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "62"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   70
         Left            =   10755
         TabIndex        =   54
         Top             =   1365
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "63"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   71
         Left            =   10755
         TabIndex        =   53
         Top             =   1905
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "64"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   72
         Left            =   10755
         TabIndex        =   52
         Top             =   2445
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "65"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   73
         Left            =   10755
         TabIndex        =   51
         Top             =   2985
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "66"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   74
         Left            =   10755
         TabIndex        =   50
         Top             =   3525
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "67"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   75
         Left            =   10755
         TabIndex        =   49
         Top             =   4065
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "68"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   76
         Left            =   10755
         TabIndex        =   48
         Top             =   4605
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "69"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   77
         Left            =   10755
         TabIndex        =   47
         Top             =   5145
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "70"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   78
         Left            =   10755
         TabIndex        =   46
         Top             =   5685
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "71"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   79
         Left            =   10755
         TabIndex        =   45
         Top             =   6225
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "72"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   80
         Left            =   10755
         TabIndex        =   44
         Top             =   6765
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "73"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   81
         Left            =   10755
         TabIndex        =   43
         Top             =   7305
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "74"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   82
         Left            =   10755
         TabIndex        =   42
         Top             =   7845
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "75"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   83
         Left            =   10755
         TabIndex        =   41
         Top             =   8385
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "76"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   84
         Left            =   10755
         TabIndex        =   40
         Top             =   8925
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "77"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   85
         Left            =   10755
         TabIndex        =   39
         Top             =   9465
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "78"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   86
         Left            =   10755
         TabIndex        =   38
         Top             =   10005
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "79"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   87
         Left            =   10755
         TabIndex        =   37
         Top             =   10545
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "80"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   88
         Left            =   10755
         TabIndex        =   36
         Top             =   11085
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "61"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   89
         Left            =   10755
         TabIndex        =   35
         Top             =   825
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "82"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   90
         Left            =   14310
         TabIndex        =   34
         Top             =   1365
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "83"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   91
         Left            =   14310
         TabIndex        =   33
         Top             =   1905
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "84"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   92
         Left            =   14310
         TabIndex        =   32
         Top             =   2445
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "85"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   93
         Left            =   14310
         TabIndex        =   31
         Top             =   2985
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "86"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   94
         Left            =   14310
         TabIndex        =   30
         Top             =   3525
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "87"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   95
         Left            =   14310
         TabIndex        =   29
         Top             =   4065
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "88"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   96
         Left            =   14310
         TabIndex        =   28
         Top             =   4605
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "89"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   97
         Left            =   14310
         TabIndex        =   27
         Top             =   5145
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "90"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   98
         Left            =   14310
         TabIndex        =   26
         Top             =   5685
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "91"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   99
         Left            =   14310
         TabIndex        =   25
         Top             =   6225
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "92"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   100
         Left            =   14310
         TabIndex        =   24
         Top             =   6765
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "93"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   101
         Left            =   14310
         TabIndex        =   23
         Top             =   7305
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "94"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   102
         Left            =   14310
         TabIndex        =   22
         Top             =   7845
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "95"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   103
         Left            =   14310
         TabIndex        =   21
         Top             =   8385
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "96"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   104
         Left            =   14310
         TabIndex        =   20
         Top             =   8925
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "97"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   105
         Left            =   14310
         TabIndex        =   19
         Top             =   9465
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "98"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   106
         Left            =   14310
         TabIndex        =   18
         Top             =   10005
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   107
         Left            =   14310
         TabIndex        =   17
         Top             =   10545
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   108
         Left            =   14175
         TabIndex        =   16
         Top             =   11085
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H0000FFFF&
         Caption         =   "81"
         BeginProperty Font 
            Name            =   "����"
            Size            =   20.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   109
         Left            =   14310
         TabIndex        =   15
         Top             =   825
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H00FF0000&
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   29
         Left            =   14295
         TabIndex        =   12
         Top             =   330
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H00FF0000&
         Caption         =   "Model Name"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   28
         Left            =   14910
         TabIndex        =   11
         Top             =   330
         Width           =   2745
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H00FF0000&
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   27
         Left            =   10740
         TabIndex        =   10
         Top             =   330
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H00FF0000&
         Caption         =   "Model Name"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   26
         Left            =   11355
         TabIndex        =   9
         Top             =   330
         Width           =   2745
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H00FF0000&
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   25
         Left            =   7185
         TabIndex        =   8
         Top             =   330
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H00FF0000&
         Caption         =   "Model Name"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   24
         Left            =   7800
         TabIndex        =   7
         Top             =   330
         Width           =   2745
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H00FF0000&
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   23
         Left            =   3630
         TabIndex        =   6
         Top             =   330
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H00FF0000&
         Caption         =   "Model Name"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   22
         Left            =   4245
         TabIndex        =   5
         Top             =   330
         Width           =   2745
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H00FF0000&
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   1
         Left            =   60
         TabIndex        =   2
         Top             =   330
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  '��� ����
         BackColor       =   &H00FF0000&
         Caption         =   "Model Name"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   2
         Left            =   675
         TabIndex        =   1
         Top             =   330
         Width           =   2745
      End
   End
   Begin BHButton.BHImageButton cmdModel 
      Height          =   690
      Index           =   1
      Left            =   7650
      TabIndex        =   216
      Top             =   12780
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   1217
      Caption         =   "LOAD"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      ForeColor       =   -2147483634
      BackColor       =   128
      AlphaColor      =   128
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton cmdModel 
      Height          =   690
      Index           =   2
      Left            =   9855
      TabIndex        =   217
      Top             =   12780
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   1217
      Caption         =   "DELETE"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      ForeColor       =   -2147483634
      BackColor       =   128
      AlphaColor      =   128
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton cmdModel 
      Height          =   690
      Index           =   3
      Left            =   12060
      TabIndex        =   218
      Top             =   12780
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   1217
      Caption         =   "EXIT"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonAttrib    =   2
      ForeColor       =   -2147483634
      BackColor       =   128
      AlphaColor      =   128
      ImgOutLineSize  =   3
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000080&
      BorderWidth     =   20
      Height          =   13815
      Left            =   30
      Top             =   60
      Width           =   19065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "Current Model  ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   20.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   450
      Left            =   780
      TabIndex        =   4
      Top             =   390
      Width           =   3735
   End
   Begin VB.Label lblCurrentModelName 
      Alignment       =   2  '��� ����
      Appearance      =   0  '���
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  '����
      Caption         =   "Model Name"
      BeginProperty Font 
         Name            =   "����"
         Size            =   24
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   540
      Left            =   5265
      TabIndex        =   3
      Top             =   285
      Width           =   2850
   End
   Begin VB.Label Label3 
      Appearance      =   0  '���
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   4995
      TabIndex        =   13
      Top             =   540
      Width           =   3435
   End
End
Attribute VB_Name = "frmModelAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TempClickModelNo As Integer '��â���� �𵨹��� Ŭ��������� �ӽ������� �� ��ȣ�� ������ ���� ����
Dim TempClickModelName As String '��â���� Ŭ��������� �ӽ������� �� �̸� ���� ����

Private Sub cmdModel_Click(Index As Integer)

Dim i As Integer
Dim SaveSW As Boolean
Dim loadsw As Boolean
Dim KillSW As Boolean
Dim TempInputName As String
Dim sOldModelName As String

    Select Case Index
    
        Case 0  'save
        
            If TempClickModelName = "" Or TempClickModelNo = 0 Then
                MsgBox "������ ��ġ�� ���õ��� �ʾҽ��ϴ�.", vbCritical, "��ġ ����"
            Else
                If TempClickModelName <> "Empty" Then
                    If MsgBox("���� �����Ͻ� ���� ������ " & TempClickModelName & " ������ ���� �մϴ�." & Chr(10) & "����Ͻðڽ��ϱ�?", vbOKCancel) = vbOK Then
                        Call LOGWrite("��ȣ[" & TempClickModelNo & "]" & TempClickModelName & "���� ���� ���� �մϴ�.(����� �����ϼ̽��ϴ�.")
                        SaveSW = Model_Create_DIR(TempClickModelName)  '�𵨸����� ���� ����
                        If SaveSW = False Then
                            MsgBox "�� ���� ������ ���� �߽��ϴ�.", vbCritical, "�� ���� ����"
                            Call LOGWrite(TempClickModelName & "�� ���� ������ �����Ͽ� �����۾��� �Ϸ����� ���Ͽ����ϴ�.")
                            Exit Sub
                        End If
                        sOldModelName = Trim(Modelinfo.ModelName)
                        Call JR_ModelSave(Trim(TempClickModelName))
                        'Call Modeldata_Filesave(Trim(TempClickModelName))   '������ ����
                        sModelName = Trim(Modelinfo.ModelName)
                        
                        Call LastModelWrite
                        'Call ModelData_SetTool
                        For i = 1 To 100
                            sModelRoom(i) = Trim(cmdModelRoom(i).Caption)
                        Next i
                        Call ModelList_SAVE                                 '�� ����Ʈ ������Ʈ
                        iNowModelNo = TempClickModelNo
                        sModelName = TempClickModelName
                        Call MasterImage_Copy(sOldModelName, sModelName)
                        Call LOGWrite("��ȣ[" & TempClickModelNo & "]" & TempClickModelName & "�� ������ ���������� �����Ͽ�����, �� �������� ��" & "��ȣ[" & TempClickModelNo & "]" & TempClickModelName & "�𵨷� ���� �˴ϴ�.")
                    End If
                Else
                    TempInputName = InputBox("�𵨸��� ������ �ּ���!", "�� �̸� �Է�")
                    If MsgBox("�Է��Ͻ� �𵨸� " & TempInputName & " ���� " & TempClickModelNo & " ������ ���� �Ͻðڽ��ϱ�?", vbOKCancel, "����") = vbOK Then
                        Call LOGWrite("��ȣ[" & TempClickModelNo & "]" & TempInputName & "�� �� �����Ͻ� �� ������ ���� �մϴ�.(�ű����� �����ϼ̽��ϴ�.)")
                        SaveSW = Model_Create_DIR(TempInputName)  '�𵨰������� ����
                        If SaveSW = False Then
                            MsgBox "�� ���� ������ ���� �߽��ϴ�.", vbCritical, "�� ���� ����"
                            Call LOGWrite(TempInputName & "�� ���� ������ �����Ͽ� �����۾��� �Ϸ����� ���Ͽ����ϴ�.")
                            Exit Sub
                        End If
                        sOldModelName = Trim(Modelinfo.ModelName)
                        Call JR_ModelSave(Trim(TempInputName))
                        'Call Modeldata_Filesave(Trim(TempInputName))   '������ ����
                        sModelName = Trim(Modelinfo.ModelName)
                        
                        Call LastModelWrite
                        'Call ModelData_SetTool
                        cmdModelRoom(TempClickModelNo).Caption = TempInputName  '�� �� ��ư ĸ�� ������Ʈ
                        For i = 1 To 100
                            sModelRoom(i) = Trim(cmdModelRoom(i).Caption)
                        Next i
                        Call ModelList_SAVE                                        '�� ����Ʈ ������Ʈ
                        iNowModelNo = TempClickModelNo
                        sModelName = TempInputName
                        TempClickModelName = TempInputName
                        Call MasterImage_Copy(sOldModelName, sModelName)
                        Call LOGWrite("��ȣ[" & TempClickModelNo & "]" & TempInputName & "�� ������ ���������� �����Ͽ�����, �� �������� ��" & "��ȣ[" & TempClickModelNo & "]" & TempClickModelName & "�𵨷� ���� �˴ϴ�.")
                    End If
                End If
            End If

        Case 1  'load
        
            If TempClickModelName = "" Or TempClickModelName = "Empty" Or TempClickModelNo = 0 Then
                MsgBox "���� ���Դϴ�. �ε��� �� �����ϴ�.", vbCritical, "�� �ҷ����� ����"
            Else
'                loadsw = ModelData_FileLoad(Trim(TempClickModelName))     '�𵨷ε�(���ϰ� �Ҹ���)
'                If loadsw = True Then           '�𵨷ε��� ���������� �Ǿ��ٸ� ���� ���� �ϹǷ�
                    sModelName = Trim(Modelinfo.ModelName)
                    Call JR_ModelLoad(sModelName)
                    Call FormControlShow
                    Call LastModelWrite             '������ �۾������� ����
                    frmMain.lblModelNameMain.Caption = sModelName
'                End If
'                If loadsw = False Then
'                    MsgBox "�� �ε忡 ���� �߽��ϴ�.", vbCritical, "�� �ҷ����� ����"
'                    Call LOGWrite("��ȣ[" & TempClickModelNo & "]" & TempClickModelName & "�� �ε��� �ڵ��˻翡 �ʿ��� ������ �ε����� ���Ͽ����ϴ�. �� �ε� �۾��� �ߴ� �Ǿ����ϴ�.")
'                    Exit Sub
'                End If
                iNowModelNo = TempClickModelNo
                sModelName = Trim(Modelinfo.ModelName)
                MsgBox "�����Ͻ� [" & TempClickModelNo & "]" & TempClickModelName & " ���� ���������� �ε� �߽��ϴ�.", vbInformation, "�� �ε� ����"
                Call LOGWrite("��ȣ[" & TempClickModelNo & "]" & TempClickModelName & "�� �ε��۾��� ���������� ���� �Ͽ����ϴ�. ���� ���� ��ȣ[" & TempClickModelNo & "]" & TempClickModelName & "�� ���� �˴ϴ�.")
                nPreJobNum = TempClickModelNo
            End If
            
        Case 2  'delete
        
            If TempClickModelNo = 0 Or TempClickModelName = "Empty" Or TempClickModelName = "None" Then
                MsgBox TempClickModelNo & " ������ �ִ� " & TempClickModelName & " ���� �ùٸ� ���� �ƴմϴ�.", vbInformation, "�� ����"
            Else
                If MsgBox(TempClickModelNo & " ������ �ִ� " & TempClickModelName & " ���� ���� �Ͻðڽ��ϱ�?", vbOKCancel, "�� ����") = vbOK Then
                
                    Call LOGWrite("��ȣ[" & TempClickModelNo & "]" & TempClickModelName & "�� ���� �۾��� ���� �ϼ̽��ϴ�.")
                    KillSW = Delete_Model_File(TempClickModelName)
                    If KillSW = True Then
                        'MsgBox TempClickModelNo & " ������ �ִ� " & TempClickModelName & " �� ������ �Ϸ� �Ǿ����ϴ�.", vbInformation
                    Else
                        MsgBox TempClickModelNo & " ������ �ִ� " & TempClickModelName & " �� ���� �����۾��� ������ �߻� �Ͽ����ϴ�.", vbCritical, "�� ���� ����"
                        Call LOGWrite("��ȣ[" & TempClickModelNo & "]" & TempClickModelName & "�� ������ ���ϻ��� ������ �߻��Ͽ� �۾��� �Ϸ����� �� �Ͽ����ϴ�.")
                        Exit Sub
                    End If
                    KillSW = Delete_Model_Dir(TempClickModelName)
                    If KillSW = True Then
                        MsgBox TempClickModelNo & " ������ �ִ� " & TempClickModelName & " �� ������ �Ϸ� �Ǿ����ϴ�.", vbInformation, "�� ���� �Ϸ�"
                        Call LOGWrite("��ȣ[" & TempClickModelNo & "]" & TempClickModelName & "�� �����۾��� ���������� ���� �Ǿ����ϴ�.")
                    Else
                        MsgBox TempClickModelNo & " ������ �ִ� " & TempClickModelName & " �� ���丮 �����۾��� ������ �߻� �Ͽ����ϴ�.", vbCritical, "�� ���� ����"
                        Call LOGWrite("��ȣ[" & TempClickModelNo & "]" & TempClickModelName & "�� ������ �������� ������ �߻��Ͽ� �۾��� �Ϸ����� �� �Ͽ����ϴ�.")
                        Exit Sub
                    End If
                    cmdModelRoom(TempClickModelNo).Caption = "Empty"  '�� �� ��ư ĸ�� ������Ʈ
                    For i = 1 To 100
                        sModelRoom(i) = Trim(cmdModelRoom(i).Caption)
                    Next i
                    Call ModelList_SAVE                                        '�� ����Ʈ ������Ʈ
                    If iNowModelNo = TempClickModelNo Then
                        TempClickModelNo = 0
                        TempClickModelName = "Empty"
                        iNowModelNo = 0
                        sModelName = "None"
                    Else
                        TempClickModelNo = 0
                        TempClickModelName = "Empty"
                    End If
                    For i = 1 To 100
                        cmdModelRoom(i).BackColor = vbWhite
                    Next i
                End If
            End If
        
        Case 3  'close
            Unload Me

    End Select
    lblCurrentModelName.Caption = "[" & iNowModelNo & "] " & sModelName
    
End Sub

Private Sub cmdModelRoom_Click(Index As Integer)
    Dim i As Integer
    
    TempClickModelNo = Index                                    '��â���� ��ư������ ���õ� �� ��ȣ
    TempClickModelName = Trim(cmdModelRoom(Index).Caption)      '��â���� ��ư������ ���õ� �� �̸�
    For i = 1 To 100
        cmdModelRoom(i).BackColor = vbWhite
    Next i
    cmdModelRoom(Index).BackColor = vbGreen
    
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim ret As Long

''    ret = ret Or WS_EX_LAYERED
''    Call SetWindowLong(frmModelAuto.hWnd, GWL_EXSTYLE, ret)
''    Call SetLayeredWindowAttributes(frmModelAuto.hWnd, vbRed, 80, LWA_COLORKEY)
    TempClickModelName = sModelName
    For i = 1 To 100
        cmdModelRoom(i).BackColor = vbWhite
        
        If Trim(sModelRoom(i)) = "" Then sModelRoom(i) = "Empty"
        Me.cmdModelRoom(i).Caption = sModelRoom(i)
        If Trim(sModelRoom(i)) = sModelName Then
            iNowModelNo = i
        End If
        If Me.cmdModelRoom(i).Caption = "" Or Me.cmdModelRoom(i).Caption = "Empty" Then
            Me.cmdModelRoom(i).ForeColor = vbBlack
        End If
    Next i
    
    If iNowModelNo <> 0 Then
        TempClickModelNo = iNowModelNo                                      '��â���� ��ư������ ���õ� �� ��ȣ
        TempClickModelName = sModelName                                  '��â���� ��ư������ ���õ� �� �̸�
        cmdModelRoom(iNowModelNo).BackColor = vbGreen
    End If

    lblCurrentModelName.Caption = "[" & iNowModelNo & "] " & sModelName

End Sub

