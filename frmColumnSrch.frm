VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Begin VB.Form frmColumnSrch 
   Caption         =   "컬럼단위 검색"
   ClientHeight    =   7710
   ClientLeft      =   5130
   ClientTop       =   5490
   ClientWidth     =   17505
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   17505
   Begin VB.Frame frameCode 
      BackColor       =   &H00FFFFFF&
      Height          =   5355
      Left            =   2760
      TabIndex        =   5
      Top             =   1740
      Visible         =   0   'False
      Width           =   10455
      Begin VB.TextBox txtCodeWhere 
         Appearance      =   0  '평면
         BackColor       =   &H00E1C7F5&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   8  '영문
         Index           =   0
         Left            =   60
         TabIndex        =   10
         Top             =   4380
         Width           =   10275
      End
      Begin VB.TextBox txtCodeWhere 
         Appearance      =   0  '평면
         BackColor       =   &H00E1C7F5&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   8  '영문
         Index           =   1
         Left            =   60
         TabIndex        =   9
         Top             =   4680
         Width           =   10275
      End
      Begin VB.TextBox txtCodeWhere 
         Appearance      =   0  '평면
         BackColor       =   &H00E1C7F5&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   8  '영문
         Index           =   2
         Left            =   60
         TabIndex        =   8
         Top             =   4980
         Width           =   10275
      End
      Begin VB.TextBox txtCodeSrch 
         Appearance      =   0  '평면
         BackColor       =   &H00C0E0FF&
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
         Left            =   1020
         TabIndex        =   7
         Top             =   120
         Width           =   2715
      End
      Begin VB.CommandButton cmdBookReset 
         BackColor       =   &H80000016&
         Caption         =   "BookMark Reset"
         Height          =   315
         Left            =   3840
         Style           =   1  '그래픽
         TabIndex        =   6
         Top             =   120
         Width           =   1875
      End
      Begin FPSpreadADO.fpSpread fpcode 
         Height          =   3855
         Left            =   60
         TabIndex        =   11
         Top             =   480
         Width           =   10275
         _Version        =   393216
         _ExtentX        =   18124
         _ExtentY        =   6800
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         GridColor       =   0
         MaxCols         =   7
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmColumnSrch.frx":0000
         ScrollBarTrack  =   1
      End
      Begin VB.Image Image1 
         Height          =   405
         Left            =   9660
         Picture         =   "frmColumnSrch.frx":054C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   510
      End
   End
   Begin VB.CheckBox chkLike 
      Caption         =   "Like %검색어%"
      Height          =   195
      Left            =   2160
      TabIndex        =   4
      Top             =   120
      Width           =   2115
   End
   Begin VB.TextBox txtTableSrchEng 
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
      Left            =   2100
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton cmdTableSrch 
      Caption         =   "조회"
      Height          =   435
      Left            =   4620
      Style           =   1  '그래픽
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin FPSpreadADO.fpSpread fpColSrch 
      Height          =   6735
      Left            =   180
      TabIndex        =   2
      Top             =   780
      Width           =   17175
      _Version        =   393216
      _ExtentX        =   30295
      _ExtentY        =   11880
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림체"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14737632
      MaxCols         =   9
      MaxRows         =   50
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      ShadowColor     =   12648447
      SpreadDesigner  =   "frmColumnSrch.frx":2F14
      ScrollBarTrack  =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "컬럼"
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
      Left            =   1500
      TabIndex        =   3
      Top             =   360
      Width           =   390
   End
End
Attribute VB_Name = "frmColumnSrch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdTableSrch_Click()
    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    If Trim(txtTableSrchEng.Text) = "" Then
        gfMsgInf ("검색어를 입력 하세요")
        Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
    
    If GCONN_USE = False Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = GetConnString
    
        conn.Open
    End If
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
    sSQL = sSQL & vbCrLf & " SELECT A.OWNER                                                                                     "
    sSQL = sSQL & vbCrLf & "       ,A.TABLE_NAME                                                                                "
    sSQL = sSQL & vbCrLf & "       ,(SELECT COMMENTS                                                                            "
    sSQL = sSQL & vbCrLf & "           FROM ALL_TAB_COMMENTS X                                                                  "
    sSQL = sSQL & vbCrLf & "          WHERE 1=1                                                                                 "
    sSQL = sSQL & vbCrLf & "            AND A.OWNER = X.OWNER                                                                   "
    sSQL = sSQL & vbCrLf & "            AND A.TABLE_NAME = X.TABLE_NAME                                                         "
    sSQL = sSQL & vbCrLf & "          )     AS TABLE_COM                                                                        "
    sSQL = sSQL & vbCrLf & "       ,A.COLUMN_NAME                                                                               "
    sSQL = sSQL & vbCrLf & "       ,B.COMMENTS                                                                                  "
    sSQL = sSQL & vbCrLf & "       ,A.DATA_TYPE                                                                                 "
    sSQL = sSQL & vbCrLf & "       ,CASE A.DATA_TYPE WHEN 'NUMBER' THEN A.DATA_PRECISION || ',' || A.DATA_SCALE                 "
    sSQL = sSQL & vbCrLf & "                         ELSE TO_CHAR(A.DATA_LENGTH)                                                "
    sSQL = sSQL & vbCrLf & "        END  DATA_LENGTH                                                                            "
    sSQL = sSQL & vbCrLf & "       ,C.CODE_DOM                                                                                  "
    sSQL = sSQL & vbCrLf & "       ,C.BIGO1                                                                                     "
    sSQL = sSQL & vbCrLf & "   FROM ALL_TAB_COLUMNS A                                                                           "
    sSQL = sSQL & vbCrLf & "        ,ALL_COL_COMMENTS B                                                                         "
    sSQL = sSQL & vbCrLf & "        ,ln.TB_DIC_COL C                                                                            "
    sSQL = sSQL & vbCrLf & "                                                                                                    "
    sSQL = sSQL & vbCrLf & "  WHERE 1=1                                                                                         "
    sSQL = sSQL & vbCrLf & "   AND A.OWNER = B.OWNER(+)                                                                         "
    sSQL = sSQL & vbCrLf & "   AND A.TABLE_NAME = B.TABLE_NAME(+)                                                               "
    sSQL = sSQL & vbCrLf & "   AND A.COLUMN_NAME = B.COLUMN_NAME (+)                                                            "
    sSQL = sSQL & vbCrLf & "   AND A.OWNER       = C.OWNER(+)                                                                   "
    sSQL = sSQL & vbCrLf & "   AND A.TABLE_NAME  = C.TABLE_NAME(+)                                                              "
    sSQL = sSQL & vbCrLf & "   AND A.COLUMN_NAME = C.COLUMN_NAME(+)                                                             "
    sSQL = sSQL & vbCrLf & "                                                                                                    "
        
    If chkLike.Value = "1" Then
        
        sSQL = sSQL & vbCrLf & "    AND (upper(A.COLUMN_NAME) LIKE upper('%" & Trim(txtTableSrchEng.Text) & "%')                                                      "
        sSQL = sSQL & vbCrLf & "         OR B.COMMENTS like '%" & Trim(txtTableSrchEng.Text) & "%'                                                                         "
        sSQL = sSQL & vbCrLf & "        )                                                                                           "
        
    Else
    
        sSQL = sSQL & vbCrLf & "    AND (upper(A.COLUMN_NAME) = upper('" & Trim(txtTableSrchEng.Text) & "')                                                      "
        sSQL = sSQL & vbCrLf & "         OR B.COMMENTS = '" & Trim(txtTableSrchEng.Text) & "'                                                                         "
        sSQL = sSQL & vbCrLf & "        )                                                                                           "
    
    End If
    
    sSQL = sSQL & vbCrLf & "    AND A.OWNER NOT IN ('SYS')                                                                      "
    sSQL = sSQL & vbCrLf & " ORDER BY 4, 1,2,3,4                                                                                       "
    
   
    frmOraMain.txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If

    fpColSrch.MaxRows = 0
    If rsTemp.RecordCount < 1 Then
        fpColSrch.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpColSrch, rsTemp)
    End If
    
   
    
        
    Screen.MousePointer = vbDefault

    Set conn = Nothing
    Set rsTemp = Nothing
    

    Exit Sub
ErrorHandler:
    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description
End Sub

Private Sub Form_Activate()
    txtTableSrchEng.SetFocus
    
    Me.Left = frmOraMain.Left + 5000
    Me.Top = frmOraMain.Top + 5000
End Sub

'------------------------------------------------------------
'공통코드 조회
'------------------------------------------------------------
Private Sub subSelectCodeList(ByVal sCode As String)
    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    
    If GCONN_USE = False Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = GetConnString
    
        conn.Open
    End If
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient


  
    fpColSrch.Row = fpColSrch.ActiveRow
    fpColSrch.Col = 1
    Dim sEng_Nm As String
    sEng_Nm = fpColSrch.Value


    sSQL = ""
    sSQL = sSQL & vbCrLf & "  SELECT  CD_DMN_NO                       "
    sSQL = sSQL & vbCrLf & "         ,CD_DMN_KRN_NM                   "
    sSQL = sSQL & vbCrLf & " , '" & sEng_Nm & "'"
    sSQL = sSQL & vbCrLf & "         ,CD_INST_VL                        "
    sSQL = sSQL & vbCrLf & "         ,CD_INST_NM                        "
    sSQL = sSQL & vbCrLf & "         ,'' USE_YN                            "
    sSQL = sSQL & vbCrLf & "         ,''  REMARK                              "
    'sSQL = sSQL & vbCrLf & "    FROM LN.VW_TABMAN_CODE                 "
    
    
sSQL = sSQL & vbCrLf & "         FROM (    SELECT B.C_CLAS_NO   AS CD_DMN_NO"
sSQL = sSQL & vbCrLf & "               ,B.C_CLAS_NM   AS CD_DMN_KRN_NM"
sSQL = sSQL & vbCrLf & "               ,B.TBL_NM      AS 테이블명"
sSQL = sSQL & vbCrLf & "               ,B.C_CLAS_DESC AS 공통코드분류설명"
sSQL = sSQL & vbCrLf & "               ,A.C_NO        AS CD_INST_VL"
sSQL = sSQL & vbCrLf & "               ,A.STND_C_NM   AS CD_INST_NM"
sSQL = sSQL & vbCrLf & "               ,A.SUMM_C_NM   AS 축약코드"
sSQL = sSQL & vbCrLf & "               ,A.C_NO || '-' || A.STND_C_NM AS 코드결합"
sSQL = sSQL & vbCrLf & "           FROM CO.TBGM_COC_TBL  A"
sSQL = sSQL & vbCrLf & "               ,CO.TBGM_COC_CLAS B"
sSQL = sSQL & vbCrLf & "          Where 1 = 1"
sSQL = sSQL & vbCrLf & "            AND A.C_CLAS_NO = B.C_CLAS_NO"
sSQL = sSQL & vbCrLf & "            )"
   
'    sSQL = sSQL & vbCrLf & "  FROM     "
'            sSQL = sSQL & vbCrLf & "(SELECT  SMPL_LCCD       AS CD_DMN_NO     --코드도메인     "
'            sSQL = sSQL & vbCrLf & "       ,SMPL_LCCD_NM    AS CD_DMN_KRN_NM  --코드도메인명 "
'            sSQL = sSQL & vbCrLf & "       ,SMPL_SCCD       AS CD_INST_VL    --코드밸류 "
'            sSQL = sSQL & vbCrLf & "       ,SMPL_SCCD_NM    AS CD_INST_NM      --코드명"
'            sSQL = sSQL & vbCrLf & "       ,SMPL_CD_STS_YN  AS USE_YN          --사용여부"
'            sSQL = sSQL & vbCrLf & "       ,CD_OUTL_CNTN1       AS REMARK          --비고"
'            'sSQL = sSQL & vbCrLf & "       ,CD_ARY_SQNC_NO AS ORD_NO          --정렬순서"
'            sSQL = sSQL & vbCrLf & "  From  TCMSMPLCODEC ) "
    sSQL = sSQL & vbCrLf & "   WHERE 1=1   "
    
    
    
    If sCode = "메타" Then
        
        fpColSrch.Row = fpColSrch.ActiveRow
        fpColSrch.Col = 2
        sSQL = sSQL & vbCrLf & "     AND TRIM(CD_DMN_KRN_NM) = '" & fpColSrch.Value & "'              "
    Else
        sSQL = sSQL & vbCrLf & "     AND upper(CD_DMN_NO) = '" & UCase(sCode) & "'              "
    End If
    sSQL = sSQL & vbCrLf & "  ORDER BY CD_DMN_NO, CD_INST_VL                       "
    

    frmOraMain.txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If

    fpcode.MaxRows = 0
    If rsTemp.RecordCount < 1 Then
        fpcode.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpcode, rsTemp)
    End If

    
    Screen.MousePointer = vbDefault

    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    Exit Sub
ErrorHandler:
    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description
'


End Sub




Private Sub fpColSrch_Click(ByVal Col As Long, ByVal Row As Long)
        If Col <> 8 Then Exit Sub
    fpColSrch.Row = Row
    fpColSrch.Col = Col
    
    
    Dim sCode As String
    
    If Trim(fpColSrch.Value) = "" Then Exit Sub
    
    'RSM 공통코드 조회
    'If UCase(Left(fpCol.Value, 2)) = "CD" Then
    '    sCode = "메타"
    'Else
        sCode = fpColSrch.Value
    'End If
    
    subSelectCodeList (sCode)
    txtCodeWhere(0).Text = ""
    txtCodeWhere(1).Text = ""
    txtCodeWhere(2).Text = ""
    
    frameCode.Visible = True
    txtCodeSrch.SetFocus
End Sub

Private Sub fpColSrch_DblClick(ByVal Col As Long, ByVal Row As Long)
'    fpColSrch.Row = Row
'    fpColSrch.Col = 1
'    frmOraMain.cboShema.Text = fpColSrch.Value
'    fpColSrch.Col = 2
'    frmOraMain.txtTableSrchEng = fpColSrch.Value
'
'     'frmOraMain.cmdTableSrch_Click
End Sub

Private Sub Image1_Click()
    frameCode.Visible = False
End Sub

Private Sub txtTableSrchEng_KeyPress(KeyAscii As Integer)
      If KeyAscii = 13 Then
       Call cmdTableSrch_Click
    End If
End Sub
