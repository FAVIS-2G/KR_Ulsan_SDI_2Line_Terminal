VERSION 5.00
Object = "{1C636623-3093-4147-A822-EBF40B4E415C}#6.0#0"; "BHButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMESRecipe 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  '없음
   Caption         =   "Recipe"
   ClientHeight    =   6690
   ClientLeft      =   390
   ClientTop       =   1740
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6690
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer TmrRecipe 
      Interval        =   1000
      Left            =   10095
      Top             =   75
   End
   Begin BHButton.BHImageButton BHBRecipeDown 
      Height          =   705
      Left            =   405
      TabIndex        =   3
      Top             =   5670
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1244
      Caption         =   "Recip 받기"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin VB.Frame fraRecipe 
      BackColor       =   &H8000000E&
      Caption         =   "Recipe Select"
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
      Height          =   5025
      Left            =   405
      TabIndex        =   0
      Top             =   585
      Width           =   9930
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   4650
         Left            =   150
         TabIndex        =   8
         Top             =   285
         Width           =   9630
         _ExtentX        =   16986
         _ExtentY        =   8202
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   128
         ForeColorFixed  =   16777215
         BackColorSel    =   -2147483643
         BackColorBkg    =   16777215
         AllowBigSelection=   0   'False
         HighLight       =   0
         GridLines       =   2
         ScrollBars      =   0
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin BHButton.BHImageButton BHBRecipeDelete 
      Height          =   705
      Left            =   4275
      TabIndex        =   4
      Top             =   5670
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1244
      Caption         =   "Recip 삭제"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton BHBRecipeSpec 
      Height          =   705
      Left            =   8145
      TabIndex        =   5
      Top             =   5670
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1244
      Caption         =   "Recip 내용"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton BHBRecipeSave 
      Height          =   705
      Left            =   6210
      TabIndex        =   6
      Top             =   5670
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1244
      Caption         =   "Recip 저장"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin BHButton.BHImageButton BHBRecipeChoise 
      Height          =   705
      Left            =   2340
      TabIndex        =   7
      Top             =   5670
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1244
      Caption         =   "Recip 선택"
      CaptionChecked  =   "BHImageButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImgOutLineSize  =   3
   End
   Begin VB.Label lblNowRecipeName 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "RECIPE NAME SELECT"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   315
      Left            =   2445
      TabIndex        =   2
      Top             =   270
      Width           =   3180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "현재 Recipe :"
      BeginProperty Font 
         Name            =   "굴림"
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
      Top             =   270
      Width           =   1725
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
Attribute VB_Name = "frmMESRecipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BHBRecipeChoise_Click()
Dim i As Integer
Dim j As Integer
    If MsgBox(iMESGridClick & "번 " & "_" & sRecipeID(iMESGridClick) & " 를 선택 하시겠습니까?", vbOKCancel) = vbOK Then
        Me.TmrRecipe.Enabled = True
        sBufRecipeID = sRecipeID(iMESGridClick)
        iBufRecipeID = iMESGridClick
        
        Call MES_DATASEND_FUNC("RECIPE_CHANGE_EVENT", "", "")
        Me.BHBRecipeChoise.Enabled = False
        Me.BHBRecipeDelete.Enabled = False
        'Me.BHBRecipeSpec.Enabled = False
    End If

End Sub

Private Sub BHBRecipeDelete_Click()
On Error GoTo err:
Dim i As Integer
Dim j As Integer
Dim temp As Integer
Dim temp2 As Integer
    If iRecipeIDcount > 1 Then
        temp2 = Me.MSFlexGrid1.Row
        temp = Me.MSFlexGrid1.TextMatrix(temp2, 0)
        Me.MSFlexGrid1.RemoveItem (temp2)
        Call MESRecipeShift(temp)
        iRecipeIDcount = iRecipeIDcount - 1
        Call DJ_MESRecipeIDCountSave
        If iRecipeIDcount - 1 > 0 Then
            For i = iRecipeIDcount - 1 To 1 Step -1
                Me.MSFlexGrid1.TextMatrix(i, 0) = iRecipeIDcount - i
            Next i
        End If
        
        
        For i = 1 To iRecipeIDcount
            Call DJ_MESRecipeSave(i)
        Next i
    End If
Exit Sub
err:
MsgBox "삭제 할 RECIPE 가 없습니다.", vbCritical, "RECIPE 삭제"
End Sub

Private Sub BHBRecipeDown_Click()
On Error GoTo err:
    Me.TmrRecipe.Enabled = True

    Call MES_DATASEND_FUNC("RECIPE_EVENT", "", "")
Exit Sub
err:
    MsgBox "RECIPE 는 10개 까지 받을 수 있습니다." & vbCrLf & "새로 받으시려면 기존 RECIPE 를 삭제 하십시오.", vbCritical, "RECIPE 받기 실패"
End Sub

Private Sub BHBRecipeSave_Click()
Dim i As Integer
    For i = 1 To iRecipeIDcount - 1          '밑에랑 합치면 안되니 주의
        sRecipeComment(iRecipeIDcount - i) = Me.MSFlexGrid1.TextMatrix(i, 2)
    Next i
    
    For i = 1 To iRecipeIDcount - 1
        Call DJ_MESRecipeSave(i)
    Next i
End Sub

Private Sub BHBRecipeSpec_Click()
    Unload frmMESDate
    'Unload frmMESRecipe
    Unload frmMESLogin
    Unload frmMESFunction
    Call ChangeViewSection(frmMESRecipePM)
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim itmpCount As Integer
    itmpCount = 1
    iRecipeIDcount = 1
    Call DJ_MESRecipeIDCountLoad
    Me.TmrRecipe.Enabled = False
    
    frmMESRecipe.MSFlexGrid1.Rows = 1
    frmMESRecipe.MSFlexGrid1.Cols = 3
    frmMESRecipe.MSFlexGrid1.FormatString = "^No.    |" & "^RECIPE ID                |" & "^COMMENT                  "

    frmMESRecipe.MSFlexGrid1.RowHeight(0) = 500
    frmMESRecipe.MSFlexGrid1.ColWidth(0) = 1000
    frmMESRecipe.MSFlexGrid1.ColWidth(1) = 5100
    frmMESRecipe.MSFlexGrid1.ColWidth(2) = 3400
    For i = 1 To iRecipeIDcount - 1
        Call DJ_MESRecipeLoad(i)
    Next i
    Call MESRecipeAllshow
    
    Call MESRecipeChange_NG                      '양승조추가해
    frmMESRecipe.BHBRecipeChoise.Enabled = False '양승조추가해
    frmMESRecipe.BHBRecipeDelete.Enabled = False '양승조추가해
    
End Sub

Private Sub MSFlexGrid1_Click()
On Error Resume Next
Dim temp As Integer
Dim i As Integer
Dim j As Integer
    iMESGridClickIdx = Me.MSFlexGrid1.Row
    
    For i = 0 To 2
        For j = 1 To iRecipeIDcount - 1
            Me.MSFlexGrid1.Col = i
            Me.MSFlexGrid1.Row = j
            Me.MSFlexGrid1.CellBackColor = vbWhite
        Next j
        Me.MSFlexGrid1.Col = i
        Me.MSFlexGrid1.Row = iMESGridClickIdx
        Me.MSFlexGrid1.CellBackColor = vbYellow
    Next i
    iMESGridClick = Me.MSFlexGrid1.TextMatrix(iMESGridClickIdx, 0)
    Me.BHBRecipeChoise.Enabled = True
    Me.BHBRecipeDelete.Enabled = True
    Me.lblNowRecipeName.Caption = "Grid_" & iMESGridClickIdx & "_&_" & "IDCount_" & iRecipeIDcount & "_&_" & "RecipeID_" & iMESGridClick
End Sub

Private Sub MSFlexGrid1_DblClick()
Dim temp As Integer
    temp = Me.MSFlexGrid1.Row
    
    Me.MSFlexGrid1.TextMatrix(temp, 2) = InputBox((iRecipeIDcount - temp + 1) & " 번 Recipe 의 COMMENT 를 입력하세요", "COMMENT 입력")
    
End Sub

Private Sub TmrRecipe_Timer()
    iTmrRecipe = iTmrRecipe + 1
    If iTmrRecipe = 3 Then
        If bMESReply = False Then
            MsgBox "MES로 부터 응답이 없습니다.", vbCritical, "타임아웃 에러"
            TmrRecipe.Enabled = False
            iTmrRecipe = 0
        Else
            TmrRecipe.Enabled = False
            iTmrRecipe = 0
        End If
    End If
End Sub
