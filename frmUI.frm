VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmUI 
   BackColor       =   &H00E6E6FF&
   Caption         =   "UI manager"
   ClientHeight    =   13515
   ClientLeft      =   1845
   ClientTop       =   840
   ClientWidth     =   26145
   Icon            =   "frmUI.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13515
   ScaleWidth      =   26145
   Begin TabDlg.SSTab SSTab1 
      Height          =   12255
      Left            =   5640
      TabIndex        =   6
      Top             =   1020
      Width           =   20355
      _ExtentX        =   35904
      _ExtentY        =   21616
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmUI.frx":048A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Source"
      TabPicture(2)   =   "frmUI.frx":04A6
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).ControlCount=   0
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
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   2895
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "검색"
      Height          =   435
      Left            =   3600
      TabIndex        =   4
      Top             =   1020
      Width           =   1815
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
      Left            =   240
      TabIndex        =   1
      Top             =   180
      Width           =   9315
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Scan"
      Height          =   495
      Left            =   9660
      TabIndex        =   0
      Top             =   60
      Width           =   1995
   End
   Begin FPSpreadADO.fpSpread fpFileList 
      Height          =   11655
      Left            =   240
      TabIndex        =   3
      Top             =   1500
      Width           =   5235
      _Version        =   393216
      _ExtentX        =   9234
      _ExtentY        =   20558
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
      SpreadDesigner  =   "frmUI.frx":04C2
      ScrollBarTrack  =   1
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   19155
   End
End
Attribute VB_Name = "frmUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private gVoFileName() As String
Private gVoFiledir() As String



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
    Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", "UI_MANAGER", "VO_FOLDER_NAME", Trim(txtFolder.Text))
        
        
        
    Screen.MousePointer = vbHourglass
    
    fpFileList.MaxRows = 0
    
    
    '파일명을 검색한다.
    Call GetMetaInfo(txtFolder.Text, Me.lblStatus)
    
    lblStatus.Caption = ""
    
    
    
    ReDim gVoFileName(0)
    ReDim gVoFiledir(0)
    

    Dim ffile As Integer
    Dim sTemp  As String
    Dim sTemp2() As String
    
    ffile = FreeFile

    Open App.Path & "\" & "uilist.txt" For Output As #ffile   ' 파일을 엽니다.
    
    For i = LBound(gBas_FileDir) + 1 To UBound(gBas_FileDir)
    
        
        If LCase(Right(gBas_FileName(i), 4)) = ".xml" Then
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
    gfMsgWrn (Err.Description)
    
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.Height = 14100
    Me.Width = 26385
    
    txtFolder.Text = GetSetting("ORAIB_TAB_MAN_BY_HAK", "UI_MANAGER", "VO_FOLDER_NAME", "C:\HF2_DEV\apprun\ws5\workspace\hf2\khfc\COR\HG")
    
'    SSTab1.Tab = 0
    
    'volist file을 읽어 그리드에 뿌린다.
    Call ReadFile
End Sub




'----------------------------------------------------------------
'ui list file을 읽어 그리드에 뿌린다.
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

    Open App.Path & "\" & "uilist.txt" For Input As #ffile   ' 파일을 엽니다.
    
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
    If Err.Number <> 53 Then    '파일이 없는 경우
        gfMsgInf (Err.Description & Err.Number)
    End If

End Sub

Private Sub txtSrchCond_Change()

    cmdFind_Click

End Sub
