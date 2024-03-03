VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  '평면
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   Caption         =   "로그인"
   ClientHeight    =   5730
   ClientLeft      =   7035
   ClientTop       =   4890
   ClientWidth     =   8040
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":030A
   ScaleHeight     =   5730
   ScaleWidth      =   8040
   Begin VB.Frame Frame3 
      BackColor       =   &H80000005&
      Height          =   1395
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   8055
      Begin VB.Label Label11 
         BackStyle       =   0  '투명
         Caption         =   "TableMan for HF"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   27.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   795
         Left            =   540
         TabIndex        =   36
         Top             =   360
         Width           =   7035
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3855
      Left            =   840
      TabIndex        =   24
      Top             =   5880
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox txtUserId 
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
         Left            =   1140
         TabIndex        =   32
         Top             =   1620
         Width           =   1575
      End
      Begin VB.TextBox txtUserPass 
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
         Left            =   1140
         TabIndex        =   31
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtFTP_Address 
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
         IMEMode         =   3  '사용 못함
         Left            =   1245
         TabIndex        =   27
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtFTP_ID 
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
         IMEMode         =   3  '사용 못함
         Left            =   1245
         TabIndex        =   26
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtFTP_PASS 
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
         IMEMode         =   3  '사용 못함
         Left            =   1245
         TabIndex        =   25
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   900
         TabIndex        =   34
         Top             =   1680
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "PASS"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   33
         Top             =   1980
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "FTP"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   735
         TabIndex        =   30
         Top             =   540
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "FTP id"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   510
         TabIndex        =   29
         Top             =   900
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "FTP Pass"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   28
         Top             =   1260
         Width           =   915
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3420
      ScaleHeight     =   825
      ScaleWidth      =   2145
      TabIndex        =   22
      Top             =   7260
      Width           =   2175
   End
   Begin VB.ComboBox cboSystemList 
      BackColor       =   &H0080C0FF&
      Height          =   300
      ItemData        =   "frmLogin.frx":1884
      Left            =   1260
      List            =   "frmLogin.frx":1886
      TabIndex        =   9
      Top             =   1560
      Width           =   3315
   End
   Begin VB.OptionButton optTest 
      BackColor       =   &H80000009&
      Caption         =   "TEST"
      Height          =   195
      Left            =   1080
      TabIndex        =   8
      Top             =   5340
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optReal 
      BackColor       =   &H80000009&
      Caption         =   "REAL"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   5340
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3195
      Left            =   60
      TabIndex        =   1
      Top             =   2460
      Width           =   7755
      Begin VB.OptionButton optASIS 
         BackColor       =   &H00FFFF00&
         Caption         =   "AS-IS"
         Height          =   180
         Left            =   4140
         TabIndex        =   23
         Top             =   2100
         Width           =   795
      End
      Begin VB.OptionButton optInfoSystem 
         BackColor       =   &H00FF00FF&
         Caption         =   "정보계"
         Height          =   180
         Left            =   3840
         TabIndex        =   21
         Top             =   2940
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton optDevSystem 
         BackColor       =   &H0000FFFF&
         Caption         =   "개발"
         Height          =   180
         Left            =   3300
         TabIndex        =   20
         Top             =   2100
         Width           =   795
      End
      Begin VB.OptionButton optTestSystem 
         BackColor       =   &H0000FF00&
         Caption         =   "테스트"
         Height          =   180
         Left            =   2400
         TabIndex        =   19
         Top             =   2100
         Width           =   915
      End
      Begin VB.OptionButton optRealSystem 
         BackColor       =   &H000080FF&
         Caption         =   "운영"
         Height          =   180
         Left            =   1620
         TabIndex        =   18
         Top             =   2100
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "삭제"
         Height          =   315
         Left            =   1680
         Style           =   1  '그래픽
         TabIndex        =   12
         Top             =   2460
         Width           =   795
      End
      Begin VB.TextBox txtSystem 
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
         Left            =   2700
         TabIndex        =   11
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "닫기"
         Height          =   315
         Left            =   4020
         Style           =   1  '그래픽
         TabIndex        =   6
         Top             =   2460
         Width           =   795
      End
      Begin VB.CommandButton cmdReg 
         Caption         =   "등록"
         Height          =   315
         Left            =   3180
         Style           =   1  '그래픽
         TabIndex        =   17
         Top             =   2460
         Width           =   795
      End
      Begin VB.TextBox txtProvider 
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
         Left            =   2700
         TabIndex        =   13
         Top             =   660
         Width           =   2175
      End
      Begin VB.TextBox txtDSN 
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
         Left            =   2700
         TabIndex        =   14
         Top             =   1020
         Width           =   2175
      End
      Begin VB.TextBox txtID 
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
         Left            =   2700
         TabIndex        =   15
         Top             =   1380
         Width           =   2175
      End
      Begin VB.TextBox txtPWD 
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
         IMEMode         =   3  '사용 못함
         Left            =   2700
         PasswordChar    =   "*"
         TabIndex        =   16
         Top             =   1740
         Width           =   2175
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "시스템명"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1800
         TabIndex        =   10
         Top             =   300
         Width           =   780
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Provider"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1800
         TabIndex        =   5
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "DSN"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2175
         TabIndex        =   4
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   3
         Top             =   1440
         Width           =   180
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "PWD"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2130
         TabIndex        =   2
         Top             =   1800
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdConfig 
      Caption         =   "설정"
      Height          =   375
      Left            =   6600
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   1980
      Width           =   1095
   End
   Begin VB.Image imgLogin 
      Height          =   390
      Left            =   4740
      MouseIcon       =   "frmLogin.frx":1888
      MousePointer    =   99  '사용자 정의
      Picture         =   "frmLogin.frx":19DA
      Top             =   1500
      Width           =   1035
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFile() As String
Dim strVer() As String
Dim iVerCnt As Integer



Private Sub cboSystemList_Click()

    
    If Trim(cboSystemList.Text) <> "" Then
        txtSystem.Text = cboSystemList.Text
        Call getRegi
        txtSystem.Locked = True
        txtSystem.BackColor = &H80000000
    Else
        txtSystem.Locked = False
        txtSystem.BackColor = &HFFFFFF
        
        txtSystem.Text = ""
        txtProvider.Text = ""
        txtDSN.Text = ""
        txtID.Text = ""
        txtPWD.Text = ""

        'optDevSystem.Value = True
        optDevSystem.Value = False
        optRealSystem.Value = False
        optTestSystem.Value = False
        optInfoSystem.Value = False
    End If
End Sub

'DB Config Close
Private Sub cmdClose_Click()
    Me.Height = C_LOGIN_HEIGHT
End Sub

'DB CONFIG 설정
Private Sub cmdConfig_Click()


    Me.Height = C_LOGIN_CONFIG_HEIGHT
End Sub

Private Sub cmdDelete_Click()
    cboSystemList.RemoveItem (cboSystemList.ListIndex)
    Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", "SYSTEM_LIST", "MAX_LIST_COUNT", cboSystemList.ListCount - 1)
    Dim i As Integer
    For i = 1 To cboSystemList.ListCount - 1

    Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", "SYSTEM_LIST", "LIST_" & i, cboSystemList.List(i))


    Next i
End Sub

'DB 설정 정보 저장
Private Sub cmdReg_Click()
    Dim strRegValue As String
    
    'If optReal.Value = True Then
    '    strRegValue = "ORAIB_CONNECT_REAL"
    'Else
    '    strRegValue = "ORAIB_CONNECT_TEST"
    'End If
    
    If Trim(txtSystem.Text) = "" Then
        Call MsgBox("시스템명을 입력하십시오.", vbInformation)
        txtSystem.SetFocus
        Exit Sub
    End If
    
    If (txtSystem.Locked = False) Then
        
        Dim K As Integer
        For K = 1 To cboSystemList.ListCount
            
            If Trim(cboSystemList.List(K)) = Trim(txtSystem.Text) Then
                Call MsgBox("중복된 시스템명이 입력되었습니다.", vbInformation)
                txtSystem.SetFocus
                Exit Sub
            End If
        Next K
    End If
    
       
    
    
    strRegValue = "ORAIB_CONNECT_" & Trim(txtSystem.Text)
    
    'DB설정 정보 레지스트리 save
    Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", strRegValue, "PROVIDER", txtProvider.Text)
    Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", strRegValue, "DSN", txtDSN.Text)
    Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", strRegValue, "USER", txtID.Text)
    Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", strRegValue, "PASS", txtPWD.Text)
    
    If optRealSystem.Value = True Then
        Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", strRegValue, "REAL", "R")
    ElseIf optTestSystem.Value = True Then
        Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", strRegValue, "REAL", "T")
    ElseIf optDevSystem.Value = True Then
        Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", strRegValue, "REAL", "D")
    ElseIf optInfoSystem.Value = True Then
        Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", strRegValue, "REAL", "I")
    ElseIf optASIS.Value = True Then
        Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", strRegValue, "REAL", "A")
    Else
        gfMsg ("운영계, 테스트계, 개발계 구분을 선택하세요")
        Exit Sub
    End If
    
    
    Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", "ORAIB_CONNECT", "FTP", txtFTP_Address.Text)
    Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", "ORAIB_CONNECT", "FTP_ID", txtFTP_ID.Text)
    Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", "ORAIB_CONNECT", "FTP_PASS", txtFTP_PASS.Text)
    
    
    If txtSystem.Locked = False Then
        cboSystemList.AddItem (txtSystem.Text)
    End If
    
    Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", "SYSTEM_LIST", "MAX_LIST_COUNT", cboSystemList.ListCount - 1)
    Dim i As Integer
    For i = 1 To cboSystemList.ListCount - 1

        Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", "SYSTEM_LIST", "LIST_" & i, cboSystemList.List(i))


    Next i

    
    
    
    
    strProvider = txtProvider.Text
    strDSN = txtDSN.Text
    strUser = txtID.Text
    strPass = txtPWD.Text
    
    If C_SPECIAL <> "N" Then
        strProvider = "MSDAORA.1"
        strDSN = "DEOSDB"
        strUser = "EOS_APUSR"
        strPass = "EOSUSR!00"
    End If
    
    
    strFTPAddress = txtFTP_Address.Text
    strFTP_ID = txtFTP_ID.Text
    strFTP_PASS = txtFTP_PASS.Text
    
    MsgBox ("변경완료")
End Sub

Private Sub GET_FTP()
'    Screen.MousePointer = vbHourglass
'    EZFTP.RemoteAddress = strFTPAddress
'    EZFTP.UserName = strFTP_ID
'    EZFTP.Password = strFTP_PASS
'
'     On Error Resume Next
'    '연결
'    EZFTP.Connect
'    If Err.Number <> 0 Then
'        Call gfMsgCrt("FTP Server 연결에 문제가 발생하였습니다. 관리자에게 문의 바랍니다." & vbCrLf & _
'                      "에러 메세지 : " & Err.Description, "FTP 연결 에러")
'        GoTo ExitHandler
'    End If
'
'    On Error GoTo ErrorHandler
'
''TabManOra.exe.1.ver
'
'    Dim fs As Object
'    Set fs = CreateObject("Scripting.FileSystemObject")
'    If fs.FileExists(App.Path & "\" & "*.ver") Then
'        fs.DeleteFile App.Path & "\" & "*.ver"
'    End If
'
'
'
'    'Get Version List File
'    iVerCnt = 0
'    ReDim strFile(0)
'    ReDim strVer(0)
'    EZFTP.GetDirectory ("*.*")
'
'
'    lblStatus.Caption = "파일버전 체크중입니다."
'    '버전 비교하여 File DownLoad
'    Dim i As Integer
'    Dim strNowVer As String
'    For i = 1 To UBound(strFile)
'        strNowVer = GetSetting("ORAIB_TAB_MAN_BY_HAK", "ORAIB_VERSION", UCase(Trim(strFile(i))), "")
'
'        'file DownLoad
'        If strNowVer = "" Or strNowVer <> strVer(i) Then
'
'            EZFTP.RemoteFile = strFile(i)
'
'
'
'            Dim blnReExe As Boolean
'            blnReExe = False
'
'            '실행화일일 경우 파일명 변경
'            If UCase(strFile(i)) = UCase(App.EXEName & ".exe") Then
'                EZFTP.LocalFile = App.Path & "/" & strFile(i) & ".tmp"
'                blnReExe = True
'            Else
'                EZFTP.LocalFile = App.Path & "/" & strFile(i)
'            End If
'
'            '파일을 다운로드한다.
'            EZFTP.Binary = True
'            EZFTP.GetFile
'
'            '클라이언트 버전 정보 레지스트리 갱신
'            SaveSetting "ORAIB_TAB_MAN_BY_HAK", "ORAIB_VERSION", UCase(Trim(strFile(i))), strVer(i)
'
'            DoEvents
'            'prg.Value = prg.Value + 1
'
'        End If
'
'    Next i
'
'    'FTP 연결 해제
'    EZFTP.Disconnect
'
'
'    If blnReExe Then
'        Screen.MousePointer = vbHourglass
'        SaveSetting "ORAIB_TAB_MAN_BY_HAK", "ORAIB_VERSION", "app_name", App.Path & "\" & App.EXEName & ".exe"
'        Shell "ChangeMan.EXE", vbNormalFocus
'        Screen.MousePointer = vbDefault
'        End
'    End If
'
'
'
'
'ExitHandler:
'    lblStatus.Caption = "Login"
'    Screen.MousePointer = vbDefault
'    Exit Sub
'ErrorHandler:
'    lblStatus.Caption = "Login"
'    Screen.MousePointer = vbDefault
'    Call gfMsgWrn(Err.Description)
End Sub


Private Sub Command1_Click()

'    Call GET_FTP
End Sub

Private Sub Command2_Click()
    frmJSP.Show
End Sub

Private Sub EZFTP_NextDirectoryEntry(ByVal FileName As String, ByVal Attributes As Long, ByVal Length As Double)

Dim i As Integer
    If Right(Trim(FileName), 3) = "ver" Then
        iVerCnt = iVerCnt + 1
        
        For i = 1 To InStr(1, Trim(FileName), " ") - 2
        
            FileName = Mid(Trim(FileName), InStr(1, Trim(FileName), " "), Len(Trim(FileName)))
        
        Next i
         
        
        FileName = Trim(FileName)
        ReDim Preserve strFile(0 To iVerCnt)
        ReDim Preserve strVer(0 To iVerCnt)
        
        strFile(iVerCnt) = Mid(FileName, 1, InStr(1, FileName, ".") + 3)
        FileName = Mid(FileName, InStr(1, FileName, ".") + 5)
        strVer(iVerCnt) = Mid(FileName, 1, InStr(1, FileName, ".") - 1)
    
    End If
End Sub

'-------------------------------------------
'Form Load
'-------------------------------------------
Private Sub Form_Load()
    Me.Top = 500 * 12
    Me.Left = 500 * 12
    
    
    
'//중복 실행 체크
    'If App.PrevInstance = True Then
    '    MsgBox "프로그램이 이미 실행되고 있습니다", vbCritical, "ERROR"
    '    Unload Me
    '    End
    'End If

    
        

    '화면 사이즈 설정
    Me.Height = C_LOGIN_HEIGHT
                
    Dim sReal As String
    'sReal = GetSetting("ORAIB_TAB_MAN_BY_HAK", "CONFIG", "REAL_TEST", "T")
    'If sReal = "T" Then
    '    optTest.Value = True
    'ElseIf sReal = "R" Then
    '    optReal.Value = True
    'End If
    
        
    '레지스트리에서 등록된 시스템 정보를 읽어와 콤보박스에 셋팅한다.
    Dim i As Integer
    Dim iMaxListCount As Integer
    iMaxListCount = GetSetting("ORAIB_TAB_MAN_BY_HAK", "SYSTEM_LIST", "MAX_LIST_COUNT", "0")
    
    Dim sSystemList As String
    cboSystemList.AddItem (" ")
    For i = 1 To CInt(iMaxListCount)
    
        sSystemList = GetSetting("ORAIB_TAB_MAN_BY_HAK", "SYSTEM_LIST", "LIST_" & i, "")
        cboSystemList.AddItem (sSystemList)
        
    Next i
    
    
    
    '레지스트리에서 최종 등록된 시스템을 디폴트로 설정한다.
    Dim sLastConnectSystem As String
    sLastConnectSystem = GetSetting("ORAIB_TAB_MAN_BY_HAK", "SYSTEM_LIST", "LAST_CONNECT_SYSTEM", "")
    cboSystemList.Text = sLastConnectSystem

    
    Call cboSystemList_Click
    
            
    'Get DB Info
    Call getRegi
    

    'lblIp.Caption = GetTheIP
    

    
    
    



End Sub

Private Sub getRegi()


    Dim strRegValue As String
    
    
    'If optReal.Value = True Then
    '    strRegValue = "ORAIB_CONNECT_REAL"
    'Else
    '    strRegValue = "ORAIB_CONNECT_TEST"
    'End If
    
    
    strRegValue = "ORAIB_CONNECT_" & cboSystemList.Text

    
    'DB설정 정보 레지스트레에서 읽어오기
    strProvider = GetSetting("ORAIB_TAB_MAN_BY_HAK", strRegValue, "PROVIDER", "")
    strDSN = GetSetting("ORAIB_TAB_MAN_BY_HAK", strRegValue, "DSN", "")
    strUser = GetSetting("ORAIB_TAB_MAN_BY_HAK", strRegValue, "USER", "")
    strPass = GetSetting("ORAIB_TAB_MAN_BY_HAK", strRegValue, "PASS", "")
    strReal = GetSetting("ORAIB_TAB_MAN_BY_HAK", strRegValue, "REAL", "")
    If strReal = "1" Then
        strReal = "R"
    ElseIf strReal = "0" Then
        strReal = "D"
    End If

    
    strFTPAddress = GetSetting("ORAIB_TAB_MAN_BY_HAK", "ORAIB_CONNECT", "FTP", "")
    strFTP_ID = GetSetting("ORAIB_TAB_MAN_BY_HAK", "ORAIB_CONNECT", "FTP_ID", "")
    strFTP_PASS = GetSetting("ORAIB_TAB_MAN_BY_HAK", "ORAIB_CONNECT", "FTP_PASS", "")

    txtProvider.Text = strProvider
    txtDSN.Text = strDSN
    txtID.Text = strUser
    txtPWD.Text = strPass
    
    If strReal = "R" Then
        optRealSystem.Value = True
    ElseIf strReal = "D" Then
        optDevSystem.Value = True
    ElseIf strReal = "T" Then
        optTestSystem.Value = True
    ElseIf strReal = "I" Then
        optInfoSystem.Value = True
    ElseIf strReal = "A" Then
        optASIS.Value = True
    End If
    

    
    txtFTP_Address.Text = strFTPAddress
    txtFTP_ID.Text = strFTP_ID
    txtFTP_PASS.Text = strFTP_PASS
    
    'If Trim(txtFTP_Address.Text) = "" Then
    '    txtFTP_Address.Text = "134.100.1.148"
    '    txtFTP_ID.Text = "simsa"
    '    txtFTP_PASS.Text = "simsa123"
    'End If
End Sub

'--------------------------------------
'Login
'--------------------------------------
Private Sub imgLogin_Click()
    
    If C_SPECIAL <> "N" Then
        If Trim(txtProvider.Text) = "" Then
            MsgBox "DATA BASE 설정정보를 등록하세요"
            cmdConfig_Click
            Exit Sub
        End If
    End If
    
    
    Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", "SYSTEM_LIST", "LAST_CONNECT_SYSTEM", cboSystemList.Text)
    
    
    
'    If optReal.Value = True Then
'        Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", "CONFIG", "REAL_TEST", "R")
'    ElseIf optTest.Value = True Then
'        Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", "CONFIG", "REAL_TEST", "T")
'    Else
'        gfMsg ("Choice Real or Test")
'    End If

    '접속한 시스템 정보를 레지스트리에 저장
    
    If C_SPECIAL <> "N" Then
    
        'FTP down load
        If App.LogMode <> 0 Then
            If Trim(txtFTP_Address.Text) <> "" Then
                GET_FTP
            End If
        End If
    
    End If




    '**************************************
    '업무화면 로드
    '**************************************
    'lblStatus.Caption = "업무화면 Loading 중....."
    DoEvents
    frmOraMain.Caption = "테이블맨-" & cboSystemList.Text
    
     If C_SPECIAL <> "N" Then
        strProvider = "OraOLEDB.Oracle.1"
        strDSN = "EXIMIIS_TEST"
        strUser = "santa"
        strPass = "santaexim"
    End If
    
    'gIPAddress = Trim(lblIp.Caption)
    
'    If Left(lblIp.Caption, 7) = "134.100" Then
'        frmBasel.Tag = "YES"
'        frmBasel.Show
'    Else
'        frmOraMain.Show
'    End If
'
        
    '운영계이면 메인 백컬러 변경
    If strReal = "R" Then
        frmOraMain.BackColor = &H80FF&
    ElseIf strReal = "D" Then
        frmOraMain.BackColor = &HFFFF&
    ElseIf strReal = "T" Then
        frmOraMain.BackColor = &HFF00&
    ElseIf strReal = "I" Then
        frmOraMain.BackColor = &HFF00FF
    ElseIf strReal = "A" Then
        frmOraMain.BackColor = &HFFFF00
    End If
    
    
    
    frmOraMain.cboSystemList.Text = cboSystemList.Text
    frmOraMain.Show
    

    
    Unload Me
    Screen.MousePointer = vbDefault

End Sub




Private Sub optReal_Click()
    Call getRegi
End Sub


Private Sub optTest_Click()
    Call getRegi
End Sub

Private Sub txtUserId_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call imgLogin_Click
    End If

End Sub
