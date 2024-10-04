VERSION 5.00
Begin VB.Form frmROI 
   Caption         =   "ROI 선택"
   ClientHeight    =   4575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton btnCancel 
      Caption         =   "취소"
      Height          =   555
      Left            =   3000
      TabIndex        =   8
      Top             =   3750
      Width           =   1545
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "확인"
      Height          =   555
      Left            =   1110
      TabIndex        =   7
      Top             =   3750
      Width           =   1545
   End
   Begin VB.OptionButton optROI 
      Caption         =   "Dummy"
      Height          =   315
      Index           =   6
      Left            =   660
      TabIndex        =   6
      Top             =   3000
      Width           =   4605
   End
   Begin VB.OptionButton optROI 
      Caption         =   "ROI 5"
      Height          =   315
      Index           =   5
      Left            =   660
      TabIndex        =   5
      Top             =   2545
      Width           =   4605
   End
   Begin VB.OptionButton optROI 
      Caption         =   "ROI 4"
      Height          =   315
      Index           =   4
      Left            =   660
      TabIndex        =   4
      Top             =   2090
      Width           =   4605
   End
   Begin VB.OptionButton optROI 
      Caption         =   "ROI 3"
      Height          =   315
      Index           =   3
      Left            =   660
      TabIndex        =   3
      Top             =   1635
      Width           =   4605
   End
   Begin VB.OptionButton optROI 
      Caption         =   "ROI 2"
      Height          =   315
      Index           =   2
      Left            =   660
      TabIndex        =   2
      Top             =   1180
      Width           =   4605
   End
   Begin VB.OptionButton optROI 
      Caption         =   "ROI 1"
      Height          =   315
      Index           =   1
      Left            =   660
      TabIndex        =   1
      Top             =   725
      Width           =   4605
   End
   Begin VB.OptionButton optROI 
      Caption         =   "기존영역"
      Height          =   315
      Index           =   0
      Left            =   660
      TabIndex        =   0
      Top             =   270
      Value           =   -1  'True
      Width           =   4605
   End
End
Attribute VB_Name = "frmROI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ROI As Long

Private Sub btnCancel_Click()

    ROI = -1
    Me.Hide
    
End Sub

Private Sub btnOK_Click()

    Me.Hide
    
End Sub

Private Sub Form_Load()

    If ROI < 0 Then
        ROI = 0
    End If
    optROI(ROI).Value = 1
    
End Sub

Private Sub optROI_Click(Index As Integer)

    ROI = Index
    

End Sub

