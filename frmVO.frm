VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmVO 
   Appearance      =   0  '평면
   BackColor       =   &H00EAF4D9&
   BorderStyle     =   1  '단일 고정
   Caption         =   "VO Manager"
   ClientHeight    =   13965
   ClientLeft      =   29475
   ClientTop       =   2025
   ClientWidth     =   24630
   Icon            =   "frmVO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   13965
   ScaleWidth      =   24630
   Begin VB.TextBox txt4 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IMEMode         =   10  '한글 
      Left            =   21120
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   1140
      Width           =   3075
   End
   Begin VB.TextBox txt3 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IMEMode         =   10  '한글 
      Left            =   18000
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   1140
      Width           =   1755
   End
   Begin VB.TextBox txt2 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IMEMode         =   10  '한글 
      Left            =   12420
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   1140
      Width           =   4155
   End
   Begin VB.TextBox txt1 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IMEMode         =   10  '한글 
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   1140
      Width           =   3915
   End
   Begin VB.TextBox txtVoname 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IMEMode         =   10  '한글 
      Left            =   6660
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   780
      Width           =   5895
   End
   Begin VB.TextBox txtVoPath 
      Appearance      =   0  '평면
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IMEMode         =   10  '한글 
      Left            =   12720
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   780
      Width           =   10695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   12375
      Left            =   6600
      TabIndex        =   5
      Top             =   1500
      Width           =   17835
      _ExtentX        =   31459
      _ExtentY        =   21828
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "VO Layout"
      TabPicture(0)   =   "frmVO.frx":048A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fpSpread1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fpWebSquare"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fpVOLayout"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkAll"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdInOutVo(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdInOutVo(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "getter setter"
      TabPicture(1)   =   "frmVO.frx":04A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtSetter"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtSetterVar"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtGetter"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtGetterVar"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtGetSet"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "source"
      TabPicture(2)   =   "frmVO.frx":04C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtVOSrc"
      Tab(2).ControlCount=   1
      Begin VB.CommandButton cmdInOutVo 
         Caption         =   "Output"
         Height          =   375
         Index           =   1
         Left            =   7800
         TabIndex        =   28
         Top             =   6550
         Width           =   7575
      End
      Begin VB.CommandButton cmdInOutVo 
         Caption         =   "Input"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   6550
         Width           =   7575
      End
      Begin VB.TextBox txtGetSet 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   16
         Top             =   6420
         Width           =   17235
      End
      Begin VB.TextBox txtGetterVar 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   8  '영문
         Left            =   -66060
         TabIndex        =   15
         Text            =   "inVO"
         Top             =   900
         Width           =   8295
      End
      Begin VB.TextBox txtGetter 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Left            =   -66120
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   14
         Top             =   1260
         Width           =   8595
      End
      Begin VB.TextBox txtSetterVar 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   8  '영문
         Left            =   -74820
         TabIndex        =   13
         Text            =   "outVO"
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox txtSetter 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   12
         Top             =   1260
         Width           =   8595
      End
      Begin VB.TextBox txtVOSrc 
         BeginProperty Font 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   10815
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   9
         Top             =   600
         Width           =   17475
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "Check1"
         Height          =   195
         Left            =   780
         TabIndex        =   7
         Top             =   840
         Width           =   195
      End
      Begin FPSpreadADO.fpSpread fpVOLayout 
         Height          =   6135
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   13815
         _Version        =   393216
         _ExtentX        =   24368
         _ExtentY        =   10821
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ColHeaderDisplay=   0
         ColsFrozen      =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         MaxCols         =   16
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ShadowColor     =   12648447
         ShadowDark      =   12632256
         SpreadDesigner  =   "frmVO.frx":04DE
         ScrollBarTrack  =   1
      End
      Begin FPSpreadADO.fpSpread fpWebSquare 
         Height          =   6135
         Left            =   13980
         TabIndex        =   25
         Top             =   360
         Width           =   3735
         _Version        =   393216
         _ExtentX        =   6588
         _ExtentY        =   10821
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ColHeaderDisplay=   0
         ColsFrozen      =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         MaxCols         =   3
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ShadowColor     =   12648447
         ShadowDark      =   12632256
         SpreadDesigner  =   "frmVO.frx":0BD6
         ScrollBarTrack  =   1
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   4575
         Left            =   120
         TabIndex        =   26
         Top             =   6960
         Width           =   7575
         _Version        =   393216
         _ExtentX        =   13361
         _ExtentY        =   8070
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ColHeaderDisplay=   0
         ColsFrozen      =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         MaxCols         =   16
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ShadowColor     =   12648447
         ShadowDark      =   12632256
         SpreadDesigner  =   "frmVO.frx":0FA6
         ScrollBarTrack  =   1
      End
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "검색"
      Height          =   435
      Left            =   4320
      TabIndex        =   3
      Top             =   1020
      Width           =   1815
   End
   Begin VB.TextBox txtSrchCond 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IMEMode         =   8  '영문
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   4095
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Scan"
      Height          =   495
      Left            =   9540
      TabIndex        =   1
      Top             =   0
      Width           =   1995
   End
   Begin VB.TextBox txtFolder 
      Appearance      =   0  '평면
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IMEMode         =   10  '한글 
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9315
   End
   Begin FPSpreadADO.fpSpread fpFileList 
      Height          =   12135
      Left            =   120
      TabIndex        =   4
      Top             =   1620
      Width           =   6315
      _Version        =   393216
      _ExtentX        =   11139
      _ExtentY        =   21405
      _StockProps     =   64
      ColsFrozen      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14737632
      MaxCols         =   2
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   12648447
      ShadowDark      =   12632256
      SpreadDesigner  =   "frmVO.frx":169E
      ScrollBarTrack  =   1
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '투명
      Caption         =   "resource path"
      Height          =   255
      Left            =   19860
      TabIndex        =   23
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '투명
      Caption         =   "creater/owner"
      Height          =   255
      Left            =   16680
      TabIndex        =   21
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '투명
      Caption         =   "resource ID"
      Height          =   255
      Left            =   11280
      TabIndex        =   19
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '투명
      Caption         =   "logicial"
      Height          =   255
      Left            =   6660
      TabIndex        =   17
      Top             =   1200
      Width           =   915
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   540
      Width           =   21915
   End
End
Attribute VB_Name = "frmVO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private gVoFileName() As String
Private gVoFiledir() As String

'그리드 전체 체크/해제
Private Sub chkAll_Click()
    Dim i As Integer
    
    With fpVOLayout
        For i = 1 To .MaxRows
            If chkAll.Value = 1 Then
                .Row = i
                .Col = 1
                .Value = True
            Else
                .Row = i
                .Col = 1
                .Value = False
            End If
        
        
        Next i
    End With
    
    Call subSetter
End Sub

'---------------------------------------------
'검색하기
'---------------------------------------------
Private Sub cmdFind_Click()
    Dim i As Integer
    
On Error GoTo Err_Handler
    Screen.MousePointer = vbHourglass
    DoEvents
    
    fpFileList.MaxRows = 0
    For i = LBound(gVoFileName) To UBound(gVoFileName)
        
        If InStr(1, UCase(gVoFileName(i)), UCase(Trim(txtSrchCond.Text))) > 0 Then
            fpFileList.MaxRows = fpFileList.MaxRows + 1
            fpFileList.Row = fpFileList.MaxRows
            fpFileList.Col = 1: fpFileList.Value = gVoFileName(i)
            fpFileList.Col = 2: fpFileList.Value = gVoFiledir(i)
        End If
    
    Next i
    
    Screen.MousePointer = vbDefault
    Exit Sub
Err_Handler:
    gfMsgWrn ("Load 후 검색하세요" & vbCrLf & Err.Description)
    Screen.MousePointer = vbDefault

End Sub

'---------------------------------------------
'Disk Load
'---------------------------------------------
Private Sub cmdLoad_Click()
On Error GoTo Err_Handler
    Dim i As Integer
    
    
    Dim iCnt As Integer
           
        
    If Trim(txtFolder.Text) = "" Then
        gfMsgInf ("폴더를 입력하세요!!")
        txtFolder.SetFocus
        Exit Sub
    End If
        
    If Len(Trim(txtFolder.Text)) <= 3 Then
        gfMsgInf ("드라이브 단위는 검색 불가합니다.")
        txtFolder.SetFocus
        Exit Sub
    End If
        
    '레지트스트리에 저장
    Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", "SORURCE_CONVERT", "VO_FOLDER_NAME", Trim(txtFolder.Text))
        
        
        
    Screen.MousePointer = vbHourglass
    
    
    
    '파일명을 검색한다.
    Call GetMetaInfo(txtFolder.Text, Me.lblStatus)
    
    lblStatus.Caption = ""
    
    fpFileList.MaxRows = 0
    
    ReDim gVoFileName(0)
    ReDim gVoFiledir(0)
    

    Dim ffile As Integer
    Dim sTemp  As String
    Dim sTemp2() As String
    
    ffile = FreeFile

    Open App.Path & "\" & "volist.txt" For Output As #ffile   ' 파일을 엽니다.
    
    For i = LBound(gBas_FileDir) + 1 To UBound(gBas_FileDir)
    
        
        If LCase(Right(gBas_FileName(i), 4)) = ".dto" Then
            fpFileList.MaxRows = fpFileList.MaxRows + 1
            fpFileList.Row = i + 1
            fpFileList.Col = 1
            fpFileList.Value = gBas_FileName(i)
            
            fpFileList.Col = 2
            fpFileList.Value = gBas_FileDir(i)
            
            '검색을 위한 로컬 전역 변수에 저장한다.
            iCnt = iCnt + 1
            ReDim Preserve gVoFileName(0 To iCnt)
            ReDim Preserve gVoFiledir(0 To iCnt)
            
            
            gVoFileName(iCnt) = gBas_FileName(i)
            gVoFiledir(iCnt) = gBas_FileDir(i)
            
            
            Print #ffile, gVoFileName(iCnt) & ";" & gVoFiledir(iCnt)
        End If
    
    Next i
    
    Close #ffile
    
    
    fpFileList.Row = -1
    fpFileList.Col = -1
    fpFileList.Lock = True
    
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
Err_Handler:
    Screen.MousePointer = vbDefault
    Close #ffile
    gfMsgWrn (Err.Description)
    
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    
    
    txtFolder.Text = GetSetting("ORAIB_TAB_MAN_BY_HAK", "SORURCE_CONVERT", "VO_FOLDER_NAME", "C:\HF2_DEV\apprun\po7\workspace")
    
    SSTab1.Tab = 0
    
    'volist file을 읽어 그리드에 뿌린다.
    Call ReadFile
End Sub


'----------------------------------------------------------------
'vo list file을 읽어 그리드에 뿌린다.
'----------------------------------------------------------------

Private Sub ReadFile()
On Error GoTo ErrHandler
    Dim ffile As Integer
    Dim sTemp  As String
    Dim sTemp2() As String
    Dim iCnt As Integer
    Dim i As Integer
    
    
    ReDim gVoFileName(0)
    ReDim gVoFiledir(0)


    ffile = FreeFile

    Open App.Path & "\" & "volist.txt" For Input As #ffile   ' 파일을 엽니다.
    
    fpFileList.MaxRows = 0
    Do
    
        Input #ffile, sTemp
        
        sTemp2 = Split(sTemp, ";")
        
        fpFileList.MaxRows = fpFileList.MaxRows + 1
        For i = LBound(sTemp2) To UBound(sTemp2)
            
            fpFileList.Row = fpFileList.MaxRows
            
            fpFileList.Col = i + 1
            fpFileList.Value = sTemp2(i)
        Next i
        
        iCnt = iCnt + 1
        ReDim Preserve gVoFileName(0 To iCnt)
        ReDim Preserve gVoFiledir(0 To iCnt)
        
        
        gVoFileName(iCnt) = sTemp2(0)
        gVoFiledir(iCnt) = sTemp2(1)
        
        
    Loop Until EOF(1)
    
    Close #ffile
    Exit Sub
ErrHandler:
    If Err.Number <> 53 And Err.Number <> 62 Then    '파일이 없는 경우
        gfMsgInf (Err.Description & Err.Number)
    End If
    
    Close #ffile
End Sub

'----------------------------------------------------------------
'setter 문장을 생성한다.
'----------------------------------------------------------------

Private Sub subSetter()
    Dim i As Integer
    
    txtSetter.Text = ""
    Dim sTemp As String
    
    
    
    
    
    Dim iMaxLen As Integer
    

    
    iMaxLen = 0
    '최대 용어 사이즈 계산
    For i = 1 To fpVOLayout.MaxRows
        fpVOLayout.Row = i
        fpVOLayout.Col = 3
        
        If Len(Trim(fpVOLayout.Value)) > iMaxLen Then
            iMaxLen = Len(Trim(fpVOLayout.Value))
        End If
        
    Next i
        
    
    
    With fpVOLayout
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            sTemp = ""
            If .Value = True Then
                sTemp = txtSetterVar.Text & ".set"
                .Col = 3
                sTemp = sTemp & UCase(Left(.Value, 1)) & Mid(.Value, 2) & "();"
                sTemp = sTemp & gfMakeString(" ", iMaxLen - Len(.Value) + 10)
                .Col = 2
                sTemp = sTemp & "//" & .Value
                txtSetter.Text = txtSetter.Text & sTemp & vbCrLf
            End If
    
    
    
        Next i
    End With


    txtGetter.Text = ""
    With fpVOLayout
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            sTemp = ""
            If .Value = True Then
                sTemp = txtGetterVar.Text & ".get"
                .Col = 3
                sTemp = sTemp & UCase(Left(.Value, 1)) & Mid(.Value, 2) & "();"
                sTemp = sTemp & gfMakeString(" ", iMaxLen - Len(.Value) + 10)
                .Col = 2
                sTemp = sTemp & "//" & .Value
                txtGetter.Text = txtGetter.Text & sTemp & vbCrLf
            End If
    
    
    
        Next i
    End With

    
    iMaxLen = 0
    '최대 용어 사이즈 계산
    For i = 1 To fpVOLayout.MaxRows
        fpVOLayout.Row = i
        fpVOLayout.Col = 3
        
        If Len(Trim(fpVOLayout.Value)) * 2 > iMaxLen Then
            iMaxLen = Len(Trim(fpVOLayout.Value)) * 2
        End If
        
    Next i
    
    txtGetSet.Text = ""
    With fpVOLayout
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            sTemp = ""
            If .Value = True Then
                sTemp = txtSetterVar.Text & ".set"
                .Col = 3
                sTemp = sTemp & UCase(Left(.Value, 1)) & Mid(.Value, 2) & "("
                sTemp = sTemp & txtGetterVar.Text & ".get"
                sTemp = sTemp & UCase(Left(.Value, 1)) & Mid(.Value, 2) & "());"
                
                sTemp = sTemp & gfMakeString(" ", iMaxLen - Len(.Value) * 2 + 10)
                .Col = 2
                sTemp = sTemp & "//" & .Value
                txtGetSet.Text = txtGetSet.Text & sTemp & vbCrLf
            End If
    
    
    
        Next i
    End With

End Sub


'----------------------------------------------------------------
'VO 파일을 읽어 text box 및 그리드에 표시한다. UTF-8 METHOD
'----------------------------------------------------------------
Private Function TextStreamRead(strPathName As String) As String

    Dim objStream As Object
    
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Open
    objStream.Type = 2
    objStream.Charset = "UTF-8"
    objStream.LoadFromFile strPathName
    TextStreamRead = objStream.ReadText
    Set objStream = Nothing


End Function

'----------------------------------------------------------------
'VO 파일을 읽어 text box 및 그리드에 표시한다.
'----------------------------------------------------------------
Private Sub fpFileList_DblClick(ByVal Col As Long, ByVal Row As Long)

'On Error GoTo Err_Handler

    Dim ffile As Integer
    Dim iCnt As Integer
    Dim sTemp As String
    Dim sTemp2() As String
    Dim j As Integer
    'Dim sResult() As String
    
    
    Dim sPathName As String
    Dim sFilename As String
    
    fpFileList.Row = Row
    fpFileList.Col = 1
    sFilename = fpFileList.Value
    
    fpFileList.Row = Row
    fpFileList.Col = 2
    sPathName = fpFileList.Value
        
    txtVoname.Text = sFilename
    txtVoPath.Text = sPathName
    
    
    
    sTemp = TextStreamRead(sPathName & "\" & sFilename)
    
    
    Dim sArr() As String
    Dim i As Integer
    
    sArr = Split(sTemp, Chr(10))
    
    
    Dim bFirst As Boolean
    bFirst = False
    txtVOSrc.Text = ""
    fpVOLayout.MaxRows = 0
    
    
    For i = LBound(sArr) To UBound(sArr)
        
        
        '    <ns5:dtoField
       ' ="" ="" ="" ="NONE" ="" =""
       ' ="-1" keyword="" ="" ="" ="e724b61dac13de1235c90520c050558e" ="true" ="false"/>
        
        
        If (InStr(1, UCase(sArr(i)), UCase("ns5:dto")) > 0 Or InStr(1, UCase(sArr(i)), UCase("ns4:dto")) > 0) And bFirst = False Then

            txt1.Text = gfGetXMLValue(sArr(i), "logicalName")
            txt2.Text = gfGetXMLValue(sArr(i), "resourceId")
            txt3.Text = gfGetXMLValue(sArr(i), "creator") & " / " & gfGetXMLValue(sArr(i), "owner")
            txt4.Text = gfGetXMLValue(sArr(i), "resourcePath")
            bFirst = True
        End If
        
        
        If (InStr(1, UCase(sArr(i)), UCase("ns5:dtoField")) > 0 Or InStr(1, UCase(sArr(i)), UCase("ns4:dtoField")) > 0) Then
        


            fpVOLayout.MaxRows = fpVOLayout.MaxRows + 1
            fpVOLayout.Row = fpVOLayout.MaxRows
            
            
            fpVOLayout.Col = 2: fpVOLayout.Value = gfGetXMLValue(sArr(i), "logicalName")    '한글명
            fpVOLayout.Col = 3: fpVOLayout.Value = gfGetXMLValue(sArr(i), "physicalName")    '물리명
            fpVOLayout.Col = 4: fpVOLayout.Value = gfGetXMLValue(sArr(i), "fieldType")  '타입
            fpVOLayout.Col = 5: fpVOLayout.Value = gfGetXMLValue(sArr(i), "includeDtoPath")
            fpVOLayout.Col = 6: fpVOLayout.Value = gfGetXMLValue(sArr(i), "includeDtoName")  'include 5
            fpVOLayout.Col = 7: fpVOLayout.Value = gfGetXMLValue(sArr(i), "Length") '길이6
            fpVOLayout.Col = 8: fpVOLayout.Value = gfGetXMLValue(sArr(i), "array")    '배열크기7
             'decimal8
            fpVOLayout.Col = 19: fpVOLayout.Value = gfGetXMLValue(sArr(i), "nullable")    'nullable9
            fpVOLayout.Col = 11: fpVOLayout.Value = gfGetXMLValue(sArr(i), "DefaultValue")    'default10
            fpVOLayout.Col = 12: fpVOLayout.Value = gfGetXMLValue(sArr(i), "Comments")    '설명11
            fpVOLayout.Col = 13: fpVOLayout.Value = gfGetXMLValue(sArr(i), "schemaName")  'scema12
            fpVOLayout.Col = 14: fpVOLayout.Value = gfGetXMLValue(sArr(i), "tableName")     'table13
            fpVOLayout.Col = 15: fpVOLayout.Value = gfGetXMLValue(sArr(i), "columnName")      'column14
            fpVOLayout.Col = 16: fpVOLayout.Value = gfGetXMLValue(sArr(i), "Key")       'is pk15
            'mask16
            'mask range17
            fpVOLayout.Col = 19: fpVOLayout.Value = gfGetXMLValue(sArr(i), "encrypt")    'encryption18
            fpVOLayout.Col = 20: fpVOLayout.Value = gfGetXMLValue(sArr(i), "metaId")    'metaId
        
        End If
        
        txtVOSrc.Text = txtVOSrc.Text & sArr(i) & vbCrLf
        
        
    Next i
    
    
    chkAll.Value = 1
    Call chkAll_Click
    
    
    Exit Sub
Err_Handler:
    If Err.Number = 53 Then
        gfMsgInf " 파일을 찾을 수 없습니다."
    Else
        gfMsgInf (Err.Description & " " & Err.Number)
    End If

End Sub

Private Sub fpVOLayout_Change(ByVal Col As Long, ByVal Row As Long)
    If Col = 1 Then
        Call subSetter
    End If
End Sub




Private Sub fpVOLayout_Click(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    With fpVOLayout
        .Row = Row
        .Col = 1
        If .Value = True Then
            .Value = False
        Else
            .Value = True
        End If
    End With
    Call subSetter
End Sub

Private Sub txtSetterVar_Change()
    Call subSetter
End Sub

Private Sub txtGetterVar_Change()
    Call subSetter
End Sub

Private Sub txtSrchCond_Change()
    cmdFind_Click
End Sub

Private Sub txtSrchCond_GotFocus()
    txtSrchCond.SelStart = 0
    txtSrchCond.SelLength = Len(txtSrchCond.Text)
End Sub

Private Sub txtSrchCond_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call cmdFind_Click
    End If
End Sub
