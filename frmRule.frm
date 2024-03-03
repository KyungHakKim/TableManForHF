VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRule 
   Caption         =   "Rule Converter"
   ClientHeight    =   13605
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   25155
   Icon            =   "frmRule.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13605
   ScaleWidth      =   25155
   Begin VB.Frame Frame4 
      BackColor       =   &H0080C0FF&
      Height          =   795
      Left            =   9360
      TabIndex        =   15
      Top             =   12720
      Width           =   15735
      Begin VB.Frame Frame3 
         BackColor       =   &H0080C0FF&
         Caption         =   "조건문 구분"
         Height          =   555
         Left            =   2760
         TabIndex        =   23
         Top             =   180
         Width           =   2655
         Begin VB.OptionButton optIfType 
            BackColor       =   &H0080C0FF&
            Caption         =   "Space"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   25
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optIfType 
            BackColor       =   &H0080C0FF&
            Caption         =   "Enter"
            Height          =   255
            Index           =   1
            Left            =   1380
            TabIndex        =   24
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0080C0FF&
         Caption         =   "숫자 비교 형식"
         Height          =   555
         Left            =   60
         TabIndex        =   20
         Top             =   180
         Width           =   2655
         Begin VB.OptionButton optNumTy 
            BackColor       =   &H0080C0FF&
            Caption         =   ">"
            Height          =   255
            Index           =   1
            Left            =   1620
            TabIndex        =   22
            Top             =   240
            Width           =   915
         End
         Begin VB.OptionButton optNumTy 
            BackColor       =   &H0080C0FF&
            Caption         =   "BigDecimal"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H0080C0FF&
         Caption         =   "조건형식"
         Height          =   555
         Left            =   5520
         TabIndex        =   17
         Top             =   180
         Width           =   2655
         Begin VB.OptionButton optCondType 
            BackColor       =   &H0080C0FF&
            Caption         =   "if ~ if"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   19
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optCondType 
            BackColor       =   &H0080C0FF&
            Caption         =   "if ~ else if "
            Height          =   255
            Index           =   0
            Left            =   120
            MaskColor       =   &H0080C0FF&
            TabIndex        =   18
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdConvert 
         BackColor       =   &H00C0FFFF&
         Caption         =   "변환"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8400
         MaskColor       =   &H00C0FFFF&
         TabIndex        =   16
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   7215
      End
   End
   Begin VB.CommandButton cmdClear 
      Appearance      =   0  '평면
      BackColor       =   &H008080FF&
      Caption         =   "초기화"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   240
      TabIndex        =   14
      Top             =   12720
      Width           =   7455
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   12495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   24855
      _ExtentX        =   43841
      _ExtentY        =   22040
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   8421631
      TabCaption(0)   =   "조건"
      TabPicture(0)   =   "frmRule.frx":048A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "labelsubrule"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fpSubRule"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fpVariable"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fpCondition"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdClean"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdExtract"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "value"
      TabPicture(1)   =   "frmRule.frx":04A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdGetReturn"
      Tab(1).Control(1)=   "cmdClean2"
      Tab(1).Control(2)=   "fpValue"
      Tab(1).Control(3)=   "fpReturn"
      Tab(1).Control(4)=   "fpComment"
      Tab(1).Control(5)=   "fpSubRule2"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Result"
      TabPicture(2)   =   "frmRule.frx":04C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fpConvert"
      Tab(2).ControlCount=   1
      Begin VB.CommandButton cmdExtract 
         Caption         =   "2) 추출"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   120
         TabIndex        =   4
         Top             =   9480
         Width           =   1455
      End
      Begin VB.CommandButton cmdGetReturn 
         Caption         =   "2) 추출"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -74760
         TabIndex        =   3
         Top             =   9480
         Width           =   1635
      End
      Begin VB.CommandButton cmdClean 
         BackColor       =   &H0080C0FF&
         Caption         =   "1) 문자열 정리 "
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         MaskColor       =   &H008080FF&
         TabIndex        =   2
         Top             =   8760
         Width           =   1455
      End
      Begin VB.CommandButton cmdClean2 
         Caption         =   "1) 문자열 정리 "
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   -74760
         TabIndex        =   1
         Top             =   8760
         Width           =   1635
      End
      Begin FPSpreadADO.fpSpread fpCondition 
         Height          =   8175
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   24555
         _Version        =   393216
         _ExtentX        =   43312
         _ExtentY        =   14420
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
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
         GridColor       =   16777215
         MaxCols         =   1
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ShadowColor     =   12648447
         ShadowDark      =   12632256
         SpreadDesigner  =   "frmRule.frx":04DE
         ScrollBarTrack  =   1
      End
      Begin FPSpreadADO.fpSpread fpConvert 
         Height          =   11775
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   24435
         _Version        =   393216
         _ExtentX        =   43101
         _ExtentY        =   20770
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
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
         MaxCols         =   1
         MaxRows         =   1
         Protect         =   0   'False
         ScrollBarExtMode=   -1  'True
         ShadowColor     =   16744576
         ShadowDark      =   12632256
         SpreadDesigner  =   "frmRule.frx":0856
         ScrollBarTrack  =   1
      End
      Begin FPSpreadADO.fpSpread fpValue 
         Height          =   8175
         Left            =   -74880
         TabIndex        =   7
         Top             =   480
         Width           =   16635
         _Version        =   393216
         _ExtentX        =   29342
         _ExtentY        =   14420
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
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
         MaxCols         =   1
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ShadowColor     =   16777088
         ShadowDark      =   12632256
         SpreadDesigner  =   "frmRule.frx":0BB6
         ScrollBarTrack  =   1
      End
      Begin FPSpreadADO.fpSpread fpVariable 
         Height          =   3375
         Left            =   1620
         TabIndex        =   8
         Top             =   9000
         Width           =   13815
         _Version        =   393216
         _ExtentX        =   24368
         _ExtentY        =   5953
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
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
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ShadowColor     =   12648447
         ShadowDark      =   12632256
         SpreadDesigner  =   "frmRule.frx":0F31
         ScrollBarTrack  =   1
      End
      Begin FPSpreadADO.fpSpread fpSubRule 
         Height          =   3375
         Left            =   15540
         TabIndex        =   9
         Top             =   9000
         Width           =   8955
         _Version        =   393216
         _ExtentX        =   15796
         _ExtentY        =   5953
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
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
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ShadowColor     =   12648447
         ShadowDark      =   12632256
         SpreadDesigner  =   "frmRule.frx":12D0
         ScrollBarTrack  =   1
      End
      Begin FPSpreadADO.fpSpread fpReturn 
         Height          =   3315
         Left            =   -72840
         TabIndex        =   10
         Top             =   8880
         Width           =   9195
         _Version        =   393216
         _ExtentX        =   16219
         _ExtentY        =   5847
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
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
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ShadowColor     =   12648447
         ShadowDark      =   12632256
         SpreadDesigner  =   "frmRule.frx":166F
         ScrollBarTrack  =   1
      End
      Begin FPSpreadADO.fpSpread fpComment 
         Height          =   8175
         Left            =   -58080
         TabIndex        =   11
         Top             =   480
         Width           =   7695
         _Version        =   393216
         _ExtentX        =   13573
         _ExtentY        =   14420
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
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
         MaxCols         =   1
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ShadowColor     =   16777088
         ShadowDark      =   12632256
         SpreadDesigner  =   "frmRule.frx":1A0E
         ScrollBarTrack  =   1
      End
      Begin FPSpreadADO.fpSpread fpSubRule2 
         Height          =   3315
         Left            =   -63480
         TabIndex        =   26
         Top             =   8880
         Width           =   12555
         _Version        =   393216
         _ExtentX        =   22146
         _ExtentY        =   5847
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
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
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ShadowColor     =   12648447
         ShadowDark      =   12632256
         SpreadDesigner  =   "frmRule.frx":1D89
         ScrollBarTrack  =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Param"
         Height          =   315
         Left            =   1680
         TabIndex        =   13
         Top             =   8760
         Width           =   2775
      End
      Begin VB.Label labelsubrule 
         Caption         =   "Sub Rule"
         Height          =   255
         Left            =   15720
         TabIndex        =   12
         Top             =   8760
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const C_ROW_HEIGHT = 40



'비교 조건에 대한 컨버전을 수행
Private Function fCompareCondChange(pParam As String) As String
    
    Dim sCond As String
    Dim sValue As String
    
    sValue = Replace(pParam, " ", "")
    
    '비교 연산자를 찾는다.
    If InStr(1, sValue, "<=") Then
        sCond = "<="
    ElseIf InStr(1, sValue, "=<") Then
        sCond = "<="
    ElseIf InStr(1, sValue, ">=") Then
        sCond = ">="
    ElseIf InStr(1, sValue, "=>") Then
        sCond = ">="
    ElseIf InStr(1, sValue, ">") Then
        sCond = ">"
    ElseIf InStr(1, sValue, "<") Then
        sCond = "<"
    ElseIf InStr(1, sValue, "=") Then
        sCond = "="
    Else
        sCond = ""
    End If
        
    Dim sArr() As String
    Dim sRet As String
    
    If sCond <> "" Then
        sArr = Split(sValue, sCond)
    End If
                    
'Find 구문변환
If InStr(1, UCase(sValue), "FIND") > 0 Then GoTo FIND_HANDLER
    
    '연산문 만들瘦 위하藪 불필요㈖ 괄호磯 제째
    If sCond <> "" Then
        sArr(0) = Replace(sArr(0), "(", "")
        sArr(0) = Replace(sArr(0), ")", "")
        sArr(1) = Replace(sArr(1), "(", "")
        sArr(1) = Replace(sArr(1), ")", "")
    End If
    
    Select Case sCond
        Case "="
            '문자 비교인 경우
            If InStr(1, sArr(1), """") > 0 Or InStr(1, sArr(1), "[") > 0 Then
                sRet = sArr(1) & ".Equals(" & sArr(0) & ")"
            Else
                If optNumTy(0).Value = True Then
                    sRet = "new BigDecimal(" & sArr(1) & ").Equals(" & sArr(0) & ") ) > 0)"
                Else
                    sRet = sArr(0) & " = " & sArr(1)
                End If
            End If
            
        Case "<="
            '문자 비교인 경우
            If InStr(1, sArr(1), """") > 0 Or InStr(1, sArr(1), "[") > 0 Then
                sRet = sArr(1) & ".compareTo(" & sArr(0) & ") >= 0"
            Else
                If optNumTy(0).Value = True Then
                    sRet = "new BigDecimal(" & sArr(1) & ").compareTo(" & sArr(0) & ") ) >= 0"
                Else
                    sRet = sArr(0) & " <= " & sArr(1)
                End If
            End If
            
        Case ">="
            '문자 비교인 경우
            If InStr(1, sArr(1), """") > 0 Or InStr(1, sArr(1), "[") > 0 Then
                sRet = sArr(1) & ".compareTo(" & sArr(0) & ") <= 0"
            Else
                If optNumTy(0).Value = True Then
                    sRet = "new BigDecimal(" & sArr(1) & ").compareTo(" & sArr(0) & ") ) <= 0"
                Else
                    sRet = sArr(0) & " >= " & sArr(1)
                End If
            End If
            
        Case "<"
            '문자 비교인 경우
            If InStr(1, sArr(1), """") > 0 Or InStr(1, sArr(1), "[") > 0 Then
                sRet = sArr(1) & ".compareTo(" & sArr(0) & ") > 0"
            Else
                If optNumTy(0).Value = True Then
                    sRet = "new BigDecimal(" & sArr(1) & ").compareTo(" & sArr(0) & ") )  > 0"
                Else
                    sRet = sArr(0) & " < " & sArr(1)
                End If
            End If
        
        Case ">"
            '문자 비교인 경우
            If InStr(1, sArr(1), """") > 0 Or InStr(1, sArr(1), "[") > 0 Then
                sRet = sArr(1) & ".compareTo(" & sArr(0) & ") < 0"
            Else
                If optNumTy(0).Value = True Then
                    sRet = "new BigDecimal(" & sArr(1) & ").compareTo(" & sArr(0) & ") ) < 0"
                Else
                    sRet = sArr(0) & " > " & sArr(1)
                End If
            End If
        Case Else
                sRet = sValue
    End Select
            
    '원래 문자 앞뒤에 괄호가 있는 경우 괄호를 붙여준다.
    sValue = Trim(sValue)
    
    If Left(sValue, 1) = "(" Then
        sRet = "(" & sRet
    ElseIf Right(sValue, 1) = ")" Then
        sRet = sRet & ")"
    End If
    
            
FIND_HANDLER:
            
    Dim sPre, sCenter, sPost As String
    Dim i As Integer
    Dim bComp, bZero As Boolean
    Dim sToken As String
    
    Dim sFindARR() As String
    
    bComp = False
    bZero = False
    
    'sValue = Trim(sValue)
    
    'Find 구문변환
    If InStr(1, UCase(sValue), "FIND") > 0 Then
    
        'Find를 대문자로 치환
        sValue = Replace(UCase(sValue), "FIND", "FIND")
        
        sPre = Mid(sValue, 1, InStr(1, sValue, "FIND") - 1)
        
        
        i = InStr(1, sValue, "FIND")
        Do
            If bComp And bZero Then Exit Do
        
            'Find 구문을 찾는다 > 0 이 나오면 끝
            sToken = Mid(sValue, i, 1)
            
            If sToken = ">" Then bComp = True
            If sToken = "0" Then bZero = True
            
            If i > Len(sValue) Then Exit Do
            
            sCenter = sCenter & sToken
            i = i + 1
        Loop
        
        sPost = Mid(sValue, i)
                
        'Center 변수에서 Find룰 치환한다.
        'Find 구문에서 ,로 배열 분리하고 첫번째가 비교문 두번째가 변수
        
        sFindARR = Split(sCenter, ",")
        
        sFindARR(0) = Replace(sFindARR(0), "FIND", "")
        sFindARR(0) = Replace(sFindARR(0), "(", "")
        
        sCenter = sFindARR(1) & ".contains (" & sFindARR(0) & ")"
        
        sRet = sPre & sCenter & sPost
        
    End If
                
    
            
    fCompareCondChange = sRet

End Function

Private Function chkPilsu() As Boolean
    
    Dim i As Integer
    
    For i = 1 To fpVariable.MaxRows
        fpVariable.Row = i
        fpVariable.Col = 2
        
        If fpVariable.Value = "" Then
            gfMsgInf ("TO-BE 변수명을 입력하세요!!")
            chkPilsu = False
            Exit Function
        End If
    Next i
    
    For i = 1 To fpSubRule.MaxRows
        fpSubRule.Row = i
        fpSubRule.Col = 2
        
        If fpSubRule.Value = "" Then
            gfMsgInf ("TO-BE Method명을 입력하세요!!")
            chkPilsu = False
            Exit Function
        End If
    Next i
    
    
    For i = 1 To fpReturn.MaxRows
        fpReturn.Row = i
        fpReturn.Col = 2
        
        If fpReturn.Value = "" Then
            gfMsgInf ("TO-BE return 변수명을 입력하세요!!")
            chkPilsu = False
            Exit Function
        End If
    Next i
    
    
    
    
    chkPilsu = True
End Function

'---------------------------------
'문자열 정리
'---------------------------------
Private Sub cmdClean_Click()


    Call subTextClear(fpCondition)
End Sub

Private Sub cmdClean2_Click()
    Call subTextClear(fpValue)
End Sub


Private Sub subTextClear(fpObject As Object)

    Dim i As Integer
    Dim j As Integer
    
    Dim sTemp As String

    With fpObject
    
        For i = 1 To fpObject.MaxRows
            For j = 1 To fpObject.MaxCols
                .Row = i
                .Col = j
                
                
                sTemp = .Value
                
                
                '앞에 3바이트가 """ 이면 원래 문자열이므로 """를 "로 치환
                If Left(sTemp, 3) = """" & """" & """" Then
                
                    sTemp = Replace(sTemp, """" & """" & """", """")
                    
                Else
                    
                    '문자열 안에 CrLF가 있으면
                    If InStr(1, sTemp, Chr(10)) > 0 Or InStr(1, sTemp, Chr(10)) Then
                        
                        '1. 앞뒤 " 제거
                        sTemp = Mid(sTemp, 2)
                        sTemp = Mid(sTemp, 1, Len(sTemp) - 1)
                        
                        '2. "" => "로 치화
                        sTemp = Replace(sTemp, """" & """", """")
                    
                        'crlf 제거
                        sTemp = Replace(sTemp, vbCrLf, " ")
                        sTemp = Replace(sTemp, Chr(10), " ")
                        sTemp = Replace(sTemp, Chr(13), " ")
                    
                    End If
                    
                
                End If
            
                .Value = sTemp
            Next j
        Next i
    
    
    End With
End Sub



'---------------------------------
' 화면 clear 이벤트
'---------------------------------
Private Sub cmdClear_Click()
    Call subGridClear
End Sub

'---------------------------------
' 변환 버튼 클릭 이벤트
'---------------------------------

Private Sub cmdConvert_Click()
    Dim i As Integer
    Dim j As Integer
    Dim sValue As String
    
    Dim sCond(100) As String
    
On Error GoTo Err_Handler
    
    If chkPilsu = False Then
        Exit Sub
    End If
    
    fpConvert.Row = 0
    fpConvert.MaxRows = fpCondition.MaxRows
    
    
    Dim sAndStr As String
    Dim sAndArr() As String
    Dim K As Integer
    
    For i = 1 To fpCondition.MaxRows
        
        sValue = ""
        For j = 1 To fpCondition.MaxCols
            fpCondition.Row = i
            fpCondition.Col = j
            sValue = sValue & Trim(fpCondition.Value)
            
            fpValue.Row = i
            fpValue.Col = j
            sValue = sValue & Trim(fpValue.Value)
            
        Next j
        
        '마지막 문장이면  end
        If Trim(sValue) = "" Then Exit For
        
        sValue = ""
        
    
    
        For j = 1 To fpCondition.MaxCols
            
            fpCondition.Row = i
            fpCondition.Col = j
        
            sValue = fpCondition.Value
            
            
            
            sAndStr = ""
            
            
            
            sValue = Replace(sValue, "AND", "AND")
            sValue = Replace(sValue, "ANd", "AND")
            sValue = Replace(sValue, "AnD", "AND")
            sValue = Replace(sValue, "And", "AND")
            sValue = Replace(sValue, "and", "AND")
            sValue = Replace(sValue, "aND", "AND")
            sValue = Replace(sValue, "aNd", "AND")
            sValue = Replace(sValue, "anD", "AND")
            
            sValue = Replace(sValue, "OR", "OR")
            sValue = Replace(sValue, "Or", "OR")
            sValue = Replace(sValue, "or", "OR")
            sValue = Replace(sValue, "oR", "OR")
            
            
'
'            '1. And 조건이 있는 경우 And 구문으로 분리한다.
'            If InStr(1, sValue, "AND") > 0 Then
'                sAndStr = "AND"
'            ElseIf InStr(1, sValue, "And") > 0 Then
'                sAndStr = "And"
'            ElseIf InStr(1, sValue, "and") > 0 Then
'                sAndStr = "and"
'            ElseIf InStr(1, sValue, "OR") > 0 Then
'                sAndStr = "OR"
'            ElseIf InStr(1, sValue, "or") > 0 Then
'                sAndStr = "or"
'            ElseIf InStr(1, sValue, "Or") > 0 Then
'                sAndStr = "Or"
'            End If
                                        
                                        
            'AND와 OR의 순서를 정의한다.
            
                                        
                                        
            'and or가 섞여있는 경우 모두 ANDOR로 변경한다.
            sValue = Replace(sValue, "AND", "VANT AND")
            sValue = Replace(sValue, "OR", "VANT OR")
                                                        
            'ANDOR 구분으로 배열 split
            If InStr(1, sValue, "VANT") > 0 Then
                sAndStr = "VANT"
            Else
                sAndStr = ""
            End If
            'And, or 구문이 있는 경우
            If sAndStr <> "" Then
                sAndArr = Split(sValue, sAndStr)
                
                'and 구문만큼 반복
                For K = 0 To UBound(sAndArr)
                    
                    If K = 0 Then
                        sCond(j) = "( " & fCompareCondChange(sAndArr(K))
                    Else
                                                
                        'ANDOR AND 로 변경되었으므
                        If Left(LTrim(sAndArr(K)), 3) = "AND" Then
                            sAndArr(K) = Replace(sAndArr(K), "AND", "")
                            sCond(j) = sCond(j) & " && " & fCompareCondChange(sAndArr(K))
                        ElseIf Left(LTrim(sAndArr(K)), 2) = "OR" Then
                            sAndArr(K) = Replace(sAndArr(K), "OR", "")
                            sCond(j) = sCond(j) & " || " & fCompareCondChange(sAndArr(K))
                        End If
                        
                        
                        '조건문을 Enter로 분리하는 경우
                        If optIfType(1).Value = True Then
                            
                            sCond(j) = sCond(j) & vbCrLf
                        
                        End If
                    End If
                    
                Next K
                
                sCond(j) = sCond(j) & " )"
                
                
            '단일 문장인 경우
            Else
            
            
                '단일 문장에 대한 변환
                sCond(j) = fCompareCondChange(sValue)
                
                
                '단일 조건을 괄호로 묶는다.
                If sCond(j) <> "" Then
                    sCond(j) = " ( " & sCond(j) & " ) "
                End If
                
                
                
            End If
            
            
            
        Next j
        
        fpConvert.Row = i
        fpConvert.Col = 1
        
        For j = 1 To UBound(sCond)
            
            If j = 1 Then
            
                If i = 1 Then
                    fpConvert.Value = "if ( " & sCond(j)
                Else
                    If optCondType(0).Value = True Then
                
                        fpConvert.Value = "} else if ( " & sCond(j)
                    Else
                        fpConvert.Value = "} " & vbCrLf & "if ( " & sCond(j)
                    End If
                End If
            Else
                If Trim(sCond(j)) <> "" Then
                    
                    
                    '조건문을 Enter로 분리하는 경우
                    If optIfType(1).Value = True Then
                        
                        fpConvert.Value = fpConvert.Value & "  &&  " & vbCrLf & sCond(j)
                    
                    Else
                        fpConvert.Value = fpConvert.Value & "  &&  " & sCond(j)
                    
                    End If

                End If
                
                
            End If
            
            
        Next j
        
        
        'cont 초기화
        For j = 1 To UBound(sCond)
            sCond(j) = ""
        Next j
        
        '전체 끝나는 괄호
        fpConvert.Value = fpConvert.Value & ")"
        
        
        'if 시작 block
        fpConvert.Value = fpConvert.Value & " {"
        
        '주석 표기
        fpComment.Row = i
        fpComment.Col = 1
         
        If Trim(fpComment.Value) <> "" Then
            fpConvert.Value = fpConvert.Value & "/* " & fpComment.Value & " */"
        End If
        
        
        Dim sValueVariable As String
        Dim sValueValue As String
        
        'value 절 치환
        For j = 1 To fpReturn.MaxRows
        
            fpConvert.Value = fpConvert.Value & vbCrLf
                
            fpReturn.Row = j
            fpReturn.Col = 2
            sValueVariable = fpReturn.Value
            
            
            fpValue.Row = i
            fpValue.Col = j
            
            
            fpConvert.Value = fpConvert.Value & Chr(9) & sValueVariable & "=" & fpValue.Value
        
        Next j
        
        
        
    Next i
    
    '제일 마지막에 닫는 괄호 하나 추가
    fpConvert.Row = i
    fpConvert.Col = 1
    fpConvert.Value = "}"
    
    Dim sOrg As String
    Dim sTrg As String
    
    
    
    '----------------------------------------
    '변수명 치환
    '----------------------------------------
    For K = 1 To fpVariable.MaxRows
            
            
            
        fpVariable.Row = K
        fpVariable.Col = 1
        sOrg = "[" & fpVariable.Value & "]"
        
        fpVariable.Col = 2
        sTrg = fpVariable.Value
        
        
        For i = 1 To fpConvert.MaxRows
            
            fpConvert.Row = i
            fpConvert.Col = 1
            
            fpConvert.Value = Replace(fpConvert.Value, sOrg, sTrg)
    
        Next i
    Next K
    
    
    '----------------------------------------
    'method명 치환
    '----------------------------------------
    For K = 1 To fpSubRule.MaxRows
            
            
            
        fpSubRule.Row = K
        fpSubRule.Col = 1
        sOrg = "{" & fpSubRule.Value & "}"
        
        fpSubRule.Col = 2
        sTrg = fpSubRule.Value
        
        
        For i = 1 To fpConvert.MaxRows
            
            fpConvert.Row = i
            fpConvert.Col = 1
            
            fpConvert.Value = Replace(fpConvert.Value, sOrg, sTrg)
    
        Next i
    Next K
    
    
    
    For K = 1 To fpSubRule2.MaxRows
            
            
            
        fpSubRule2.Row = K
        fpSubRule2.Col = 1
        sOrg = "{" & fpSubRule2.Value & "}"
        
        fpSubRule2.Col = 2
        sTrg = fpSubRule2.Value
        
        
        For i = 1 To fpConvert.MaxRows
            
            fpConvert.Row = i
            fpConvert.Col = 1
            
            fpConvert.Value = Replace(fpConvert.Value, sOrg, sTrg)
    
        Next i
    Next K
    
    
    
    
    SSTab2.Tab = 2
    
    Exit Sub
Err_Handler:
    gfMsgInf (Err.Description)
    
End Sub

'---------------------------------
' 변수 추출
' 조건문에서 변수만 추출한다.
' 변수는 [] 안에 있는 것만 변수
'---------------------------------
Private Sub cmdExtract_Click()
On Error GoTo Err_Handler
    Dim i As Integer
    Dim j As Integer
    Dim sValue As String
    
    Dim iPos, iCnt As Integer
    Dim sArr() As String
    Dim iArrCnt As Integer
    Dim sRet As String
    
    iArrCnt = 0
    For i = 1 To fpCondition.MaxRows
        For j = 1 To fpCondition.MaxCols
            
            fpCondition.Row = i
            fpCondition.Col = j
            sValue = fpCondition.Value
            
            
            
            Do
                If sValue = "" Then Exit Do
                
                '변수 찾기
                sRet = gfGetVariable(sValue)
                
                
                
                
                '변수하나 찾으면 앞에 문자열 잘라내기
                If sRet = "" Then
                    sValue = ""
                Else
                    iArrCnt = iArrCnt + 1
                    ReDim Preserve sArr(1 To iArrCnt)
                    sArr(iArrCnt) = sRet
                
                    sValue = Mid(sValue, InStr(1, sValue, "]") + 1)
                End If
            Loop
        Next j
    Next i
    
    Dim sLastArr() As String
    Dim bFind As Boolean
    
    fpVariable.MaxRows = 0
    If iArrCnt > 0 Then
        For i = 1 To UBound(sArr)
                
            If i = 1 Then
                fpVariable.MaxRows = fpVariable.MaxRows + 1
                fpVariable.Row = fpVariable.MaxRows
                fpVariable.Col = 1
                fpVariable.Value = sArr(i)
            Else
                bFind = False
                For j = 1 To fpVariable.MaxRows
                    fpVariable.Row = j
                    fpVariable.Col = 1
                    
                    If fpVariable.Value = sArr(i) Then
                        bFind = True
                    End If
            
                Next j
                
                If bFind = False Then
                    fpVariable.MaxRows = fpVariable.MaxRows + 1
                    fpVariable.Row = fpVariable.MaxRows
                    fpVariable.Col = 1
                    fpVariable.Value = sArr(i)
                End If
                
            End If
        Next i
    End If
    
    'sub rule 찾기
    Call FindSubRule
    
    Exit Sub
Err_Handler:
    gfMsgInf (Err.Description)
End Sub

'---------------------------------------------
'변수만 추출한다.
'---------------------------------------------
Private Function gfGetVariable(sValue As String) As String
    Dim iPos As String
    Dim sTemp As String
    
    sTemp = sValue
    
    '변수에서 [ 시작위치 찾기
    iPos = InStr(1, sTemp, "[")
    If iPos > 0 Then
        sTemp = Mid(sTemp, iPos + 1)
        
        iPos = InStr(1, sTemp, "]")
        sTemp = Left(sTemp, iPos - 1)
    End If
    
    '변수 리턴
    If iPos > 0 Then
        gfGetVariable = sTemp
    Else
        gfGetVariable = ""
    End If


End Function

'-------------------------------------------------
' Sub Rule 찾기
'--------------------------------------------------
Private Sub FindSubRule()
    Dim i As Integer
    Dim j As Integer
    Dim sValue As String
    
    Dim iPos, iCnt As Integer
    Dim sArr() As String
    Dim iArrCnt As Integer
    Dim sRet As String
    
    iArrCnt = 0
    For i = 1 To fpCondition.MaxRows
        For j = 1 To fpCondition.MaxCols
            
            fpCondition.Row = i
            fpCondition.Col = j
            sValue = fpCondition.Value
            
            
            
            Do
                If sValue = "" Then Exit Do
                
                'sub rule명 찾기
                sRet = gfGetSubRule(sValue)
                
                
                
                
                '변수하나 찾으면 앞에 문자열 잘라내기
                If sRet = "" Then
                    sValue = ""
                Else
                    iArrCnt = iArrCnt + 1
                    ReDim Preserve sArr(1 To iArrCnt)
                    sArr(iArrCnt) = sRet
                
                    sValue = Mid(sValue, InStr(1, sValue, "}") + 1)
                End If
            Loop
        Next j
    Next i
    
    Dim sLastArr() As String
    Dim bFind As Boolean
    
    fpSubRule.MaxRows = 0
    
    If iArrCnt > 0 Then
        For i = 1 To UBound(sArr)
                
            If i = 1 Then
                fpSubRule.MaxRows = fpSubRule.MaxRows + 1
                fpSubRule.Row = fpSubRule.MaxRows
                fpSubRule.Col = 1
                fpSubRule.Value = sArr(i)
            Else
                bFind = False
                For j = 1 To fpSubRule.MaxRows
                    fpSubRule.Row = j
                    fpSubRule.Col = 1
                    
                    If fpSubRule.Value = sArr(i) Then
                        bFind = True
                    End If
            
                Next j
                
                If bFind = False Then
                    fpSubRule.MaxRows = fpSubRule.MaxRows + 1
                    fpSubRule.Row = fpSubRule.MaxRows
                    fpSubRule.Col = 1
                    fpSubRule.Value = sArr(i)
                End If
                
            End If
        Next i
    End If
End Sub


'--------------------------------------
'sub rule만 추출한다.
'--------------------------------------
Private Function gfGetSubRule(sValue As String) As String
    Dim iPos As String
    Dim sTemp As String
    
    sTemp = sValue
    
    '변수에서 [ 시작위치 찾기
    iPos = InStr(1, sTemp, "{")
    If iPos > 0 Then
        sTemp = Mid(sTemp, iPos + 1)
        
        iPos = InStr(1, sTemp, "}")
        sTemp = Left(sTemp, iPos - 1)
    End If
    
    '변수 리턴
    If iPos > 0 Then
        gfGetSubRule = sTemp
    Else
        gfGetSubRule = ""
    End If


End Function


'---------------------------------
' Value에서 Return 변수를 찾는다.
'---------------------------------
Private Sub cmdGetReturn_Click()


    Dim i As Integer
    Dim j As Integer
    Dim sValue As String
    
    Dim iPos, iCnt As Integer
    Dim sArr(100) As String
    Dim iArrCnt As Integer
    Dim sRet As String
    
    For i = 1 To fpValue.MaxRows
        For j = 1 To fpValue.MaxCols
            
            fpValue.Row = i
            fpValue.Col = j
            sValue = fpValue.Value
            
            If Trim(sValue) <> "" Then
                sArr(j) = "Y"
            End If
            
        Next j
    Next i
    
    fpReturn.MaxRows = 0
    
    For i = 1 To UBound(sArr)
                    
        If sArr(i) = "Y" Then
            fpReturn.MaxRows = fpReturn.MaxRows + 1
            fpReturn.Row = fpReturn.MaxRows
            fpReturn.Col = 1
            fpReturn.Value = "Return" & i
        End If
    Next i

    
    'Return value 에서 서브룰 찾기
    FindSubRule2
    
    
End Sub


'-------------------------------------------------
' Sub Rule 찾기
'--------------------------------------------------
Private Sub FindSubRule2()
    Dim i As Integer
    Dim j As Integer
    Dim sValue As String
    
    Dim iPos, iCnt As Integer
    Dim sArr() As String
    Dim iArrCnt As Integer
    Dim sRet As String
    
    iArrCnt = 0
    For i = 1 To fpValue.MaxRows
        For j = 1 To fpValue.MaxCols
            
            fpValue.Row = i
            fpValue.Col = j
            sValue = fpValue.Value
            
            
            
            Do
                If sValue = "" Then Exit Do
                
                'sub rule명 찾기
                sRet = gfGetSubRule(sValue)
                
                
                
                
                '변수하나 찾으면 앞에 문자열 잘라내기
                If sRet = "" Then
                    sValue = ""
                Else
                    iArrCnt = iArrCnt + 1
                    ReDim Preserve sArr(1 To iArrCnt)
                    sArr(iArrCnt) = sRet
                
                    sValue = Mid(sValue, InStr(1, sValue, "}") + 1)
                End If
            Loop
        Next j
    Next i
    
    Dim sLastArr() As String
    Dim bFind As Boolean
    
    fpSubRule2.MaxRows = 0
    
    If iArrCnt > 0 Then
        For i = 1 To UBound(sArr)
                
            If i = 1 Then
                fpSubRule2.MaxRows = fpSubRule2.MaxRows + 1
                fpSubRule2.Row = fpSubRule2.MaxRows
                fpSubRule2.Col = 1
                fpSubRule2.Value = sArr(i)
            Else
                bFind = False
                For j = 1 To fpSubRule2.MaxRows
                    fpSubRule2.Row = j
                    fpSubRule2.Col = 1
                    
                    If fpSubRule2.Value = sArr(i) Then
                        bFind = True
                    End If
            
                Next j
                
                If bFind = False Then
                    fpSubRule2.MaxRows = fpSubRule2.MaxRows + 1
                    fpSubRule2.Row = fpSubRule2.MaxRows
                    fpSubRule2.Col = 1
                    fpSubRule2.Value = sArr(i)
                End If
                
            End If
        Next i
    End If
    
    '그리드를 sort한다.
    Call gfGrid_Sort(fpSubRule2, 1, 0)
    
'
'    '서브룰을 order by 한다.
'    Dim sOrderByOrg() As String
'    Dim sOrderBy() As String
'    Dim iOrderCnt As Integer
'
'
'
'    Dim sWord1 As String
'    Dim sSplit1() As String
'    Dim sCompare1 As String
'
'    Dim sWord2 As String
'    Dim sSplit2() As String
'    Dim sCompare2 As String
'
'    iOrderCnt = 0
'    With fpSubRule2
'        For i = 1 To .MaxRows
'            .Row = i
'            .Col = 1
'            sWord1 = Trim(.Value)
'
'
'            iOrderCnt = iOrderCnt + 1
'            ReDim Preserve sOrderBy(0 To iOrderCnt)
'            sOrderBy(iOrderCnt) = sWord1
'
'            '::가 포함되어 있으면 Split 한다.
'            If InStr(1, sWord1, "::") > 0 Then
'                sSplit1 = Split(sWord1, "::")
'                sCompare1 = sSplit(0)
'            Else
'                sCompare1 = sWord1
'            End If
'
'
'
'            For j = 2 To .MaxRows
'                .Row = i
'                .Col = 1
'
'                sWord2 = Trim(.Value)
'
'                '::가 포함되어 있으면 Split 한다.
'                If InStr(1, sWord2, "::") > 0 Then
'                    sSplit2 = Split(sWord2, "::")
'                    sCompare2 = sSplit(0)
'                Else
'                    sCompare2 = sWord2
'                End If
'
'
'            Next j
'        Next i
'    End With
'
    
End Sub


'---------------------------------
' Form Load Event
'---------------------------------
Private Sub Form_Load()
    
    Me.Top = 0
    Me.Left = 0
    
    Me.Height = 14265
    Me.Width = 25800
    
    
    Call subGridClear

    SSTab2.Tab = 0
    
    
    fpVariable.MaxRows = 0
    fpSubRule.MaxRows = 0
    fpReturn.MaxRows = 0
'    fpCondition.Col = 1
'    fpCondition.Row = 1
'    fpCondition.Value = "[조회일]>='20100908' and [경매시작일]='2222'"
    
End Sub


'Grid Clear
Private Sub subGridClear()
   Dim i As Integer
    
    '그리드 초기화
    With fpCondition
        .MaxRows = 0
        .MaxCols = 50
        .MaxRows = 200
        For i = 1 To 100
            .RowHeight(i) = C_ROW_HEIGHT
        Next i
        For i = 1 To 10
            .ColWidth(i) = 35
        Next i
        
        .Col = -1
        .Row = -1
        .CellType = CellTypeEdit
        .TypeEditMultiLine = True
        .TypeMaxEditLen = 9999
        
        .Col = 0
        .Row = 0
        
    End With
    
    '그리드 초기화
    With fpValue
        .MaxRows = 0
        .MaxCols = 50
        .MaxRows = 200
        For i = 1 To 100
            .RowHeight(i) = C_ROW_HEIGHT
        Next i
        For i = 1 To 10
            .ColWidth(i) = 25
        Next i
    
        .Col = -1
        .Row = -1
        .CellType = CellTypeEdit
        .TypeEditMultiLine = True
        .TypeMaxEditLen = 9999
    
    
    End With
    
    '그리드 초기화
    With fpConvert
        .MaxRows = 0
        .MaxCols = 1
        .MaxRows = 200
        For i = 1 To 100
            .RowHeight(i) = C_ROW_HEIGHT
        Next i
    
        .Col = -1
        .Row = -1
        .CellType = CellTypeEdit
        .TypeEditMultiLine = True
        .TypeMaxEditLen = 9999
    
    End With
    
    '그리드 초기화
    With fpComment
        .MaxRows = 0
        .MaxCols = 1
        .MaxRows = 200
        For i = 1 To 100
            .RowHeight(i) = C_ROW_HEIGHT
        Next i
    
        .Col = -1
        .Row = -1
        .CellType = CellTypeEdit
        .TypeEditMultiLine = True
        .TypeMaxEditLen = 9999
    
    End With
    
    
    With fpConvert
        .RowHeight(1) = C_ROW_HEIGHT
    End With
    
    fpSubRule.MaxRows = 0
    fpSubRule2.MaxRows = 0
    fpReturn.MaxRows = 0
    fpVariable.MaxRows = 0
 
End Sub

