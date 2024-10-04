VERSION 5.00
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmMESRecipePM 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  '¾øÀ½
   Caption         =   "RecipeParameter"
   ClientHeight    =   6690
   ClientLeft      =   390
   ClientTop       =   1740
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6690
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TmrRecipePM 
      Interval        =   1000
      Left            =   10185
      Top             =   75
   End
   Begin VB.Frame fraRecipPM 
      BackColor       =   &H8000000E&
      Caption         =   "Recipe Parameter"
      BeginProperty Font 
         Name            =   "µ¸¿ò"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4860
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   9990
      Begin BHButton.BHImageButton BHBSV_CHANGE 
         Height          =   375
         Left            =   5910
         TabIndex        =   89
         Top             =   1290
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   661
         Caption         =   "º¯°æ¿äÃ»"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin VB.TextBox txtMESValueMin 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   7815
         TabIndex        =   86
         Top             =   3945
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueMax 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   8835
         TabIndex        =   85
         Top             =   3945
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueMin 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   7815
         TabIndex        =   84
         Top             =   3585
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueMax 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   8835
         TabIndex        =   83
         Top             =   3585
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueMin 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   7815
         TabIndex        =   82
         Top             =   3225
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueMax 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   8835
         TabIndex        =   81
         Top             =   3225
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueMin 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   7815
         TabIndex        =   80
         Top             =   2865
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueMax 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   8835
         TabIndex        =   79
         Top             =   2865
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueMin 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   7815
         TabIndex        =   78
         Top             =   2505
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueMax 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   8835
         TabIndex        =   77
         Top             =   2505
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueMin 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   7815
         TabIndex        =   76
         Top             =   2145
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueMax 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   8835
         TabIndex        =   75
         Top             =   2145
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueMin 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   7815
         TabIndex        =   74
         Top             =   1785
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueMax 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   8835
         TabIndex        =   73
         Top             =   1785
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueMin 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   7815
         TabIndex        =   72
         Top             =   1425
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueMax 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   8835
         TabIndex        =   71
         Top             =   1425
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueMin 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   7815
         TabIndex        =   70
         Top             =   1065
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueMax 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   8835
         TabIndex        =   69
         Top             =   1065
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueMin 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   7815
         TabIndex        =   68
         Top             =   705
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueMax 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   8835
         TabIndex        =   67
         Top             =   705
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueMin 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -885
         TabIndex        =   66
         Top             =   585
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueMax 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -885
         TabIndex        =   65
         Top             =   780
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueMax 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   4800
         TabIndex        =   62
         Top             =   3945
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueMax 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   4800
         TabIndex        =   61
         Top             =   3585
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueMax 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   4800
         TabIndex        =   60
         Top             =   3225
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueMax 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   4800
         TabIndex        =   59
         Top             =   2865
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueMax 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   4800
         TabIndex        =   58
         Top             =   2505
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueMax 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   4800
         TabIndex        =   57
         Top             =   2145
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueMax 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   4800
         TabIndex        =   56
         Top             =   1785
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueMax 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   4800
         TabIndex        =   55
         Top             =   1425
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueMax 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   4800
         TabIndex        =   54
         Top             =   1065
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueMax 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4800
         TabIndex        =   52
         Top             =   705
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueMin 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   3795
         TabIndex        =   51
         Top             =   3945
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueMin 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   3795
         TabIndex        =   50
         Top             =   3585
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueMin 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   3795
         TabIndex        =   49
         Top             =   3225
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueMin 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   3795
         TabIndex        =   48
         Top             =   2865
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueMin 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   3795
         TabIndex        =   47
         Top             =   2505
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueMin 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   3795
         TabIndex        =   46
         Top             =   2145
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueMin 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   3795
         TabIndex        =   45
         Top             =   1785
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueMin 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3795
         TabIndex        =   44
         Top             =   1425
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueMin 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3795
         TabIndex        =   43
         Top             =   1065
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueMax 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -840
         TabIndex        =   42
         Top             =   45
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueMin 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3795
         TabIndex        =   41
         Top             =   705
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueMin 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -870
         TabIndex        =   39
         Top             =   195
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueOri 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   6825
         TabIndex        =   37
         Top             =   3945
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueOri 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   2790
         TabIndex        =   36
         Top             =   3945
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueOri 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   6825
         TabIndex        =   34
         Top             =   3585
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueOri 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   2790
         TabIndex        =   33
         Top             =   3585
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueOri 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   6825
         TabIndex        =   31
         Top             =   3225
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueOri 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   2790
         TabIndex        =   30
         Top             =   3225
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueOri 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   6825
         TabIndex        =   28
         Top             =   2865
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueOri 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   2790
         TabIndex        =   27
         Top             =   2865
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueOri 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   6825
         TabIndex        =   25
         Top             =   2505
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueOri 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   2790
         TabIndex        =   24
         Top             =   2505
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueOri 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   6825
         TabIndex        =   22
         Top             =   2145
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueOri 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   2790
         TabIndex        =   21
         Top             =   2145
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueOri 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   6825
         TabIndex        =   19
         Top             =   1785
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueOri 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   2790
         TabIndex        =   18
         Top             =   1785
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueOri 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   6825
         TabIndex        =   16
         Top             =   1425
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueOri 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   2790
         TabIndex        =   15
         Top             =   1425
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueOri 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   6825
         TabIndex        =   13
         Top             =   1065
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueOri 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2790
         TabIndex        =   12
         Top             =   1065
         Width           =   1020
      End
      Begin VB.TextBox txtMESValueOri 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   6825
         TabIndex        =   10
         Top             =   705
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox txtPCValueOri 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2790
         TabIndex        =   9
         Top             =   705
         Width           =   1020
      End
      Begin VB.TextBox txtMESValue 
         BeginProperty Font 
            Name            =   "µ¸¿ò"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -855
         TabIndex        =   5
         Top             =   -90
         Visible         =   0   'False
         Width           =   1020
      End
      Begin BHButton.BHImageButton BHBMESValSave 
         Height          =   375
         Left            =   5910
         TabIndex        =   90
         Top             =   3465
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   661
         Caption         =   "¹Ù·ÎÀû¿ë"
         CaptionChecked  =   "BHImageButton1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "±¼¸²"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ImgOutLineSize  =   3
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Åõ¸í
         Caption         =   "<-"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   6195
         TabIndex        =   88
         Top             =   3120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Åõ¸í
         Caption         =   "->"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   6210
         TabIndex        =   87
         Top             =   945
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label7 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00000080&
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
         Caption         =   "MESMax"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   8835
         TabIndex        =   64
         Top             =   345
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label6 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00000080&
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
         Caption         =   "MESMin"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   7830
         TabIndex        =   63
         Top             =   345
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label5 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00000080&
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
         Caption         =   "¼³ºñMax"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   4800
         TabIndex        =   53
         Top             =   345
         Width           =   1020
      End
      Begin VB.Label Label4 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00000080&
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
         Caption         =   "¼³ºñMin"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   3795
         TabIndex        =   40
         Top             =   345
         Width           =   1020
      End
      Begin VB.Label lblMESParameter 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00000080&
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   10
         Left            =   165
         TabIndex        =   35
         Top             =   3945
         Width           =   2640
      End
      Begin VB.Label lblMESParameter 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00000080&
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   9
         Left            =   165
         TabIndex        =   32
         Top             =   3585
         Width           =   2640
      End
      Begin VB.Label lblMESParameter 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00000080&
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   8
         Left            =   165
         TabIndex        =   29
         Top             =   3225
         Width           =   2640
      End
      Begin VB.Label lblMESParameter 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00000080&
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   7
         Left            =   165
         TabIndex        =   26
         Top             =   2865
         Width           =   2640
      End
      Begin VB.Label lblMESParameter 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00000080&
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   6
         Left            =   165
         TabIndex        =   23
         Top             =   2505
         Width           =   2640
      End
      Begin VB.Label lblMESParameter 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00000080&
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   5
         Left            =   165
         TabIndex        =   20
         Top             =   2145
         Width           =   2640
      End
      Begin VB.Label lblMESParameter 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00000080&
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   4
         Left            =   165
         TabIndex        =   17
         Top             =   1785
         Width           =   2640
      End
      Begin VB.Label lblMESParameter 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00000080&
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   3
         Left            =   165
         TabIndex        =   14
         Top             =   1425
         Width           =   2640
      End
      Begin VB.Label lblMESParameter 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00000080&
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   2
         Left            =   165
         TabIndex        =   11
         Top             =   1065
         Width           =   2640
      End
      Begin VB.Label lblMESParameter 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00000080&
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   1
         Left            =   165
         TabIndex        =   8
         Top             =   705
         Width           =   2640
      End
      Begin VB.Label Label3 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00000080&
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
         Caption         =   "MES±âÁØ"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   6825
         TabIndex        =   7
         Top             =   345
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label2 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00000080&
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
         Caption         =   "¼³ºñ±âÁØ"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   2790
         TabIndex        =   6
         Top             =   345
         Width           =   1020
      End
      Begin VB.Label lblMESParameter 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         BackColor       =   &H00000080&
         BorderStyle     =   1  '´ÜÀÏ °íÁ¤
         Caption         =   "Ç× ¸ñ"
         BeginProperty Font 
            Name            =   "±¼¸²"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   0
         Left            =   165
         TabIndex        =   4
         Top             =   345
         Width           =   2640
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00000080&
         BackStyle       =   1  'Åõ¸íÇÏÁö ¾ÊÀ½
         BorderStyle     =   0  'Åõ¸í
         Height          =   345
         Left            =   5970
         Shape           =   4  'µÕ±Ù »ç°¢Çü
         Top             =   915
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000080&
         BackStyle       =   1  'Åõ¸íÇÏÁö ¾ÊÀ½
         BorderStyle     =   0  'Åõ¸í
         Height          =   345
         Left            =   5985
         Shape           =   4  'µÕ±Ù »ç°¢Çü
         Top             =   3090
         Visible         =   0   'False
         Width           =   690
      End
   End
   Begin BHButton.BHImageButton BHBRecipeSelect 
      Height          =   705
      Left            =   360
      TabIndex        =   3
      Top             =   5580
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1244
      Caption         =   "Recip ¼±ÅÃ"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   128
      AlphaColor      =   128
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton BHBRecipeValSave 
      Height          =   705
      Left            =   2280
      TabIndex        =   38
      Top             =   5580
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   1244
      Caption         =   "Recipe °ª ÀúÀå"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   128
      AlphaColor      =   128
      ImgOutLineSize  =   3
   End
   Begin VB.Label lblNowRecipeName 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "RECIPE NAME SELECT"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   315
      Left            =   2565
      TabIndex        =   2
      Top             =   315
      Width           =   3180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "¼±ÅÃµÈ Recipe :"
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   420
      TabIndex        =   1
      Top             =   315
      Width           =   1980
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000080&
      BorderWidth     =   18
      Height          =   6495
      Left            =   135
      Top             =   105
      Width           =   10470
   End
End
Attribute VB_Name = "frmMESRecipePM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BHBMESValSave_Click()
    Call DJ_EquipSpecApply_NG
End Sub

Private Sub BHBRecipeSelect_Click()
    
    Unload frmMESDate
    Unload frmMESRecipe
    Unload frmMESLogin
    Unload frmMESRecipePM
    Unload frmMESFunction
    'Me.TmrRecipePM.Enabled = True
    Call ChangeViewSection(frmMESRecipe)
End Sub

Private Sub BHImageButton1_Click()
    If MsgBox("Recipe ¼öÁ¤¿äÃ»À» ÇÏ½Ã°Ú½À´Ï±î?", vbOKCancel, "RECIPE ¼öÁ¤") = vbOK Then
    
    End If
    
End Sub

Private Sub BHBSV_CHANGE_Click()
    Me.TmrRecipePM.Enabled = True
    Call MES_DATASEND_FUNC("RECIPE_SV_CHANGE_EVENT", "", "")
End Sub

Private Sub Form_Load()
Dim i As Integer
    Me.TmrRecipePM.Enabled = False
    Me.lblNowRecipeName.Caption = sNowRecipeID
    For i = 1 To iParamCount(iNowRecipeID)
        'Call DJ_MESRecipeLoad(i)
        'Call DJ_EquipSpecLoad(sParamName_SV(iNowRecipeID, i), i, iNowRecipeID)
        'Call DJ_ComparePN(sParamName_SV(iNowRecipeID, i), i, iNowRecipeID)            'MES SPEC ·Îµå (ÆÄ¶ó¹ÌÅÍ ¸ÞÄª ÈÄ)
    Next i
End Sub

Private Sub TmrRecipePM_Timer()
    iTmrRecipePM = iTmrRecipePM + 1
    If iTmrRecipePM = 3 Then
        If bMESReply = False Then
            MsgBox "MES·Î ºÎÅÍ ÀÀ´äÀÌ ¾ø½À´Ï´Ù.", vbCritical, "Å¸ÀÓ¾Æ¿ô ¿¡·¯"
            TmrRecipePM.Enabled = False
            iTmrRecipe = 0
        Else
            TmrRecipePM.Enabled = False
            iTmrRecipe = 0
        End If
    End If
End Sub
