VERSION 5.00
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Begin VB.Form frmTmpLogin 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  '����
   Caption         =   "����� �α���"
   ClientHeight    =   6690
   ClientLeft      =   390
   ClientTop       =   1740
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Begin BHButton.BHImageButton BHBMESLogin 
      Height          =   840
      Left            =   2985
      TabIndex        =   5
      Top             =   3705
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   1482
      Caption         =   "�α��� ��û"
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
      ImgOutLineSize  =   3
   End
   Begin VB.Frame fraLogin 
      BackColor       =   &H8000000E&
      Caption         =   "�α��� ��û"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   2970
      TabIndex        =   0
      Top             =   1995
      Width           =   4905
      Begin VB.TextBox txtMesPW 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1695
         TabIndex        =   4
         Top             =   930
         Width           =   2805
      End
      Begin VB.TextBox txtMesID 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1695
         TabIndex        =   2
         Top             =   405
         Width           =   2805
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "��й�ȣ :"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   330
         TabIndex        =   3
         Top             =   960
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "��� ��ȣ :"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   330
         TabIndex        =   1
         Top             =   435
         Width           =   1305
      End
   End
   Begin BHButton.BHImageButton BHBLoginCancel 
      Height          =   840
      Left            =   5475
      TabIndex        =   6
      Top             =   3705
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   1482
      Caption         =   "�� ��"
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
      ImgOutLineSize  =   3
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000080&
      BorderWidth     =   18
      Height          =   6495
      Left            =   135
      Top             =   105
      Width           =   10470
   End
End
Attribute VB_Name = "frmTmpLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mb_CertificationOfLogin As Boolean
Private Sub BHBLoginCancel_Click()
    Unload Me
    
End Sub

Private Sub BHBMESLogin_Click()

On Error GoTo err

    If Len(CStr(txtMesID.Text)) = 7 And Len(CStr(txtMesPW.Text)) <> 0 Then
        mb_CertificationOfLogin = True
        
        '�α��� �α� ����
        Dim fp As Integer
        Dim i As Integer
        Dim SHDate As String
        Dim SHTime As String
        Dim FileName_Result As String
        
        SHDate = Format(Date, "yy-mm-dd")
        SHTime = Format(Time, "hh_mm_ss")
        
        Call Create_DIR("D:\Login_Log\")
        FileName_Result = "D:\Login_Log\" & "Login_Log(" & SHDate & "," & SHTime & ").fav"
        fp = FreeFile
            
        Open FileName_Result For Output As fp
        
                Print #fp, CStr(frmTmpLogin.txtMesID)
                Print #fp, CStr(frmTmpLogin.txtMesPW)
        Close fp
        iTmrLogin = 0
        frmMain.TmrLogin.Enabled = True
        MsgBox "�α��� ���� ����", vbCritical
    Else
        MsgBox CStr(frmTmpLogin.txtMesID) & "    �α��� ����", vbCritical
    End If
    
    Exit Sub
    
err:
    Close fp
    MsgBox CStr(frmTmpLogin.txtMesID) & "    �α��� ����", vbCritical
    mb_CertificationOfLogin = False
    
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_Load()
    frmMain.TmrLogin.Enabled = False
    Me.txtMesID.Text = sMesUserID
End Sub


