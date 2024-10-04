VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   0  '없음
   Caption         =   "KMS Engine Shop Vision System"
   ClientHeight    =   6345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10395
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   423
   ScaleMode       =   3  '픽셀
   ScaleWidth      =   693
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   360
      Left            =   720
      TabIndex        =   1
      Top             =   3135
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Image imgCustomerLogo 
      Height          =   1365
      Index           =   0
      Left            =   1350
      Picture         =   "frmSplash.frx":0000
      Top             =   480
      Width           =   7800
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  '투명
      Caption         =   "Samsung SDI Terminal Vision Program"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Index           =   0
      Left            =   735
      TabIndex        =   3
      Top             =   2040
      Width           =   9015
   End
   Begin VB.Label LblState 
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   720
      TabIndex        =   2
      Top             =   3630
      Width           =   8925
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  '투명
      Caption         =   "Samsung SDI Terminal Vision Program"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   1200
      Index           =   1
      Left            =   765
      TabIndex        =   0
      Top             =   2085
      Width           =   9015
   End
   Begin VB.Image imgCustomerLogo 
      Height          =   6330
      Index           =   1
      Left            =   0
      Picture         =   "frmSplash.frx":3390
      Top             =   0
      Width           =   10395
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Top = 5000
    Me.Left = 3000
    
    'Call SetProgramTile(lblTitle(0), conStationTitle & " " & conProgramVersion)
    'Call SetProgramTile(lblTitle(1), conStationTitle & " " & conProgramVersion)
    
    Me.Show
    Me.Refresh
End Sub

Public Sub SetText(strMsg As String, Color As Variant, Optional ByVal Delay As Long = 0)
On Error GoTo ErrHandler

    LblState.ForeColor = Color
    LblState.Caption = strMsg
    
    Me.Refresh
    
    Call Sleep(Delay)
    
    Exit Sub
ErrHandler:
     Call WriteErrorLog("frmSplash SetText" & " : " & err.Description)
End Sub

Public Sub SetPos(nPos As Long)
On Error GoTo ErrHandler
    ProgressBar.Value = nPos
    Exit Sub
ErrHandler:
     Call WriteErrorLog("frmSplash SetPos" & " : " & err.Description)
End Sub

