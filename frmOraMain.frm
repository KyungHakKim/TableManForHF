VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmOraMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "TabManOra"
   ClientHeight    =   15090
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   24075
   Icon            =   "frmOraMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   15090
   ScaleWidth      =   24075
   Begin VB.ComboBox cboSystemList 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "frmOraMain.frx":030A
      Left            =   21000
      List            =   "frmOraMain.frx":030C
      TabIndex        =   0
      Top             =   60
      Width           =   3015
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   14955
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   23940
      _ExtentX        =   42228
      _ExtentY        =   26379
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   7
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Home"
      TabPicture(0)   =   "frmOraMain.frx":030E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cboSubject"
      Tab(0).Control(1)=   "txtTableSrchKor"
      Tab(0).Control(2)=   "txtTabMapping"
      Tab(0).Control(3)=   "Command5"
      Tab(0).Control(4)=   "Command3"
      Tab(0).Control(5)=   "txtMapping"
      Tab(0).Control(6)=   "Frame4"
      Tab(0).Control(7)=   "Frame3"
      Tab(0).Control(8)=   "Label1"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Table Manager"
      TabPicture(1)   =   "frmOraMain.frx":032A
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "SSTab3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "SSTab8"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdColumn"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdNeed"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "SSTab2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "SQL Format"
      TabPicture(2)   =   "frmOraMain.frx":0346
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Stored 검색"
      TabPicture(3)   =   "frmOraMain.frx":0362
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "SSTab6"
      Tab(3).Control(1)=   "SSTab7"
      Tab(3).ControlCount=   2
      Begin TabDlg.SSTab SSTab7 
         Height          =   14265
         Left            =   -74880
         TabIndex        =   196
         Top             =   480
         Width           =   23355
         _ExtentX        =   41196
         _ExtentY        =   25162
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   7
         TabHeight       =   520
         TabCaption(0)   =   "SP조회"
         TabPicture(0)   =   "frmOraMain.frx":037E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label13"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label12"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "txtSRC2"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "fpSP"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtSPconect"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cmdSPSearch"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "txtSPname"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "미사용"
         TabPicture(1)   =   "frmOraMain.frx":039A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame6"
         Tab(1).Control(1)=   "cmdViewSrc"
         Tab(1).Control(2)=   "cmdReload"
         Tab(1).Control(3)=   "cmbPgmIDLog"
         Tab(1).Control(4)=   "txtEdDay"
         Tab(1).Control(5)=   "txtStDay"
         Tab(1).Control(6)=   "cmdDebugLog"
         Tab(1).Control(7)=   "frameSrc"
         Tab(1).Control(8)=   "fpDebugLog"
         Tab(1).Control(9)=   "fpDebugLogHist"
         Tab(1).Control(10)=   "Label8"
         Tab(1).ControlCount=   11
         Begin VB.Frame Frame6 
            Caption         =   "Frame6"
            Height          =   12495
            Left            =   -74880
            TabIndex        =   217
            Top             =   480
            Width           =   19095
         End
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   12
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   11175
            IMEMode         =   8  '영문
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   208
            Text            =   "frmOraMain.frx":03B6
            Top             =   420
            Width           =   16875
         End
         Begin VB.CommandButton cmdViewSrc 
            BackColor       =   &H0080FFFF&
            Caption         =   "소스보기"
            Height          =   555
            Left            =   -64020
            Style           =   1  '그래픽
            TabIndex        =   207
            Top             =   480
            Width           =   2475
         End
         Begin VB.CommandButton cmdReload 
            Caption         =   "re-load →"
            Height          =   315
            Left            =   -74520
            Style           =   1  '그래픽
            TabIndex        =   206
            Top             =   720
            Width           =   1035
         End
         Begin VB.ComboBox cmbPgmIDLog 
            Height          =   300
            ItemData        =   "frmOraMain.frx":04F3
            Left            =   -73380
            List            =   "frmOraMain.frx":04F5
            TabIndex        =   205
            Top             =   720
            Width           =   4875
         End
         Begin VB.TextBox txtEdDay 
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
            Left            =   -65640
            TabIndex        =   204
            Top             =   720
            Width           =   1515
         End
         Begin VB.TextBox txtStDay 
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
            Left            =   -67260
            TabIndex        =   203
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton cmdDebugLog 
            Caption         =   "조회"
            Height          =   555
            Left            =   -68400
            Style           =   1  '그래픽
            TabIndex        =   202
            Top             =   480
            Width           =   975
         End
         Begin VB.Frame frameSrc 
            BackColor       =   &H00FFFFFF&
            Height          =   10335
            Left            =   -74100
            TabIndex        =   200
            Top             =   1440
            Width           =   16215
            Begin VB.TextBox txtSrc 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   12
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   9555
               IMEMode         =   8  '영문
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   3  '양방향
               TabIndex        =   201
               Top             =   600
               Width           =   15915
            End
            Begin VB.Image Image2 
               Height          =   405
               Left            =   15480
               Picture         =   "frmOraMain.frx":04F7
               Stretch         =   -1  'True
               Top             =   120
               Width           =   510
            End
         End
         Begin VB.TextBox txtSPname 
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
            Left            =   1020
            TabIndex        =   199
            Top             =   660
            Width           =   3375
         End
         Begin VB.CommandButton cmdSPSearch 
            Caption         =   "조회"
            Height          =   555
            Left            =   4440
            Style           =   1  '그래픽
            TabIndex        =   198
            Top             =   660
            Width           =   1335
         End
         Begin VB.TextBox txtSPconect 
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
            Left            =   1020
            TabIndex        =   197
            Top             =   1020
            Width           =   3375
         End
         Begin FPSpreadADO.fpSpread fpSpread1 
            Height          =   2355
            Left            =   -74880
            TabIndex        =   209
            Top             =   405
            Width           =   17175
            _Version        =   393216
            _ExtentX        =   30295
            _ExtentY        =   4154
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   200
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            SelectBlockOptions=   0
            SpreadDesigner  =   "frmOraMain.frx":2EBF
            ScrollBarTrack  =   1
         End
         Begin FPSpreadADO.fpSpread fpDebugLog 
            Height          =   6615
            Left            =   -74820
            TabIndex        =   210
            Top             =   1080
            Width           =   16995
            _Version        =   393216
            _ExtentX        =   29977
            _ExtentY        =   11668
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
            MaxCols         =   6
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            ShadowColor     =   12648447
            SpreadDesigner  =   "frmOraMain.frx":51A3
            ScrollBarTrack  =   1
         End
         Begin FPSpreadADO.fpSpread fpDebugLogHist 
            Height          =   4155
            Left            =   -74820
            TabIndex        =   211
            Top             =   7800
            Width           =   14415
            _Version        =   393216
            _ExtentX        =   25426
            _ExtentY        =   7329
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
            MaxCols         =   4
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            ShadowColor     =   12648447
            SpreadDesigner  =   "frmOraMain.frx":569F
            ScrollBarTrack  =   1
         End
         Begin FPSpreadADO.fpSpread fpSP 
            Height          =   11895
            Left            =   60
            TabIndex        =   212
            Top             =   1380
            Width           =   5835
            _Version        =   393216
            _ExtentX        =   10292
            _ExtentY        =   20981
            _StockProps     =   64
            ColsFrozen      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   0
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
            SpreadDesigner  =   "frmOraMain.frx":5B18
            ScrollBarTrack  =   1
         End
         Begin RichTextLib.RichTextBox txtSRC2 
            Height          =   13395
            Left            =   6000
            TabIndex        =   216
            Top             =   600
            Width           =   16995
            _ExtentX        =   29977
            _ExtentY        =   23627
            _Version        =   393217
            ScrollBars      =   2
            TextRTF         =   $"frmOraMain.frx":5EE4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Fixedsys"
               Size            =   12
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "SP명을 직접입력하거나 콤보박스에서 선택하세요"
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
            Left            =   -73380
            TabIndex        =   215
            Top             =   480
            Width           =   4350
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "SP명"
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
            Left            =   120
            TabIndex        =   214
            Top             =   720
            Width           =   465
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "SP내용"
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
            Left            =   120
            TabIndex        =   213
            Top             =   1080
            Width           =   660
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   3855
         Left            =   60
         TabIndex        =   55
         Top             =   11040
         Width           =   23655
         _ExtentX        =   41725
         _ExtentY        =   6800
         _Version        =   393216
         Tabs            =   5
         TabsPerRow      =   8
         TabHeight       =   520
         BackColor       =   -2147483644
         TabCaption(0)   =   "sample Data"
         TabPicture(0)   =   "frmOraMain.frx":5F79
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblHint"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "fpSample"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdColRow"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "chkCode"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cmdMinMax(0)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cmdMinMax(1)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "cmdCopyRow"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "cmdInsertUpdate"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "cmdDelLine"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Command7"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).ControlCount=   10
         TabCaption(1)   =   "INDEX"
         TabPicture(1)   =   "frmOraMain.frx":5F95
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fpIndex"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Partition"
         TabPicture(2)   =   "frmOraMain.frx":5FB1
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fpPart"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "TABLE Info."
         TabPicture(3)   =   "frmOraMain.frx":5FCD
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "getClipinsert"
         Tab(3).Control(1)=   "txtInsertData"
         Tab(3).Control(2)=   "cmdMakeInsert"
         Tab(3).Control(3)=   "fpInfo"
         Tab(3).ControlCount=   4
         TabCaption(4)   =   "데이터생성"
         TabPicture(4)   =   "frmOraMain.frx":5FE9
         Tab(4).ControlEnabled=   0   'False
         Tab(4).ControlCount=   0
         Begin VB.CommandButton Command7 
            Caption         =   "sql보기"
            Height          =   315
            Left            =   18600
            Style           =   1  '그래픽
            TabIndex        =   220
            Top             =   0
            Width           =   1335
         End
         Begin VB.CommandButton cmdDelLine 
            Caption         =   "삭제"
            Height          =   315
            Left            =   15840
            Style           =   1  '그래픽
            TabIndex        =   156
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdInsertUpdate 
            Caption         =   "저장"
            Height          =   315
            Left            =   17040
            Style           =   1  '그래픽
            TabIndex        =   155
            Top             =   0
            Width           =   1455
         End
         Begin VB.CommandButton cmdCopyRow 
            Caption         =   "Row복사"
            Height          =   315
            Left            =   14880
            Style           =   1  '그래픽
            TabIndex        =   154
            Top             =   0
            Width           =   915
         End
         Begin VB.CommandButton getClipinsert 
            Caption         =   "getClip"
            Height          =   435
            Left            =   -56280
            TabIndex        =   153
            Top             =   1200
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox txtInsertData 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   12
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            IMEMode         =   8  '영문
            Left            =   -56640
            MultiLine       =   -1  'True
            ScrollBars      =   3  '양방향
            TabIndex        =   152
            Top             =   2820
            Visible         =   0   'False
            Width           =   6075
         End
         Begin VB.CommandButton cmdMakeInsert 
            Caption         =   "쿼리생성"
            Height          =   435
            Left            =   -56280
            TabIndex        =   151
            Top             =   720
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CommandButton cmdMinMax 
            Caption         =   "↓"
            Height          =   315
            Index           =   1
            Left            =   22920
            Style           =   1  '그래픽
            TabIndex        =   150
            Top             =   -60
            Width           =   735
         End
         Begin VB.CommandButton cmdMinMax 
            Caption         =   "↑"
            Height          =   315
            Index           =   0
            Left            =   22140
            Style           =   1  '그래픽
            TabIndex        =   149
            Top             =   -60
            Width           =   735
         End
         Begin VB.CheckBox chkCode 
            Caption         =   "코드보기"
            Height          =   195
            Left            =   19440
            TabIndex        =   148
            Top             =   60
            Width           =   1275
         End
         Begin VB.CommandButton cmdColRow 
            Caption         =   "Col/Row Rotate"
            Height          =   315
            Left            =   20280
            Style           =   1  '그래픽
            TabIndex        =   86
            Top             =   0
            Width           =   1815
         End
         Begin FPSpreadADO.fpSpread fpIndex 
            Height          =   3375
            Left            =   -74940
            TabIndex        =   56
            Top             =   405
            Width           =   17415
            _Version        =   393216
            _ExtentX        =   30718
            _ExtentY        =   5953
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GridColor       =   16777215
            MaxCols         =   5
            MaxRows         =   1
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            SpreadDesigner  =   "frmOraMain.frx":6005
            ScrollBarTrack  =   1
         End
         Begin FPSpreadADO.fpSpread fpSample 
            Height          =   3375
            Left            =   120
            TabIndex        =   57
            Top             =   420
            Width           =   23415
            _Version        =   393216
            _ExtentX        =   41301
            _ExtentY        =   5953
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   200
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            SelectBlockOptions=   0
            SpreadDesigner  =   "frmOraMain.frx":6506
            ScrollBarTrack  =   1
         End
         Begin FPSpreadADO.fpSpread fpPart 
            Height          =   1155
            Left            =   -74940
            TabIndex        =   58
            Top             =   360
            Width           =   17415
            _Version        =   393216
            _ExtentX        =   30718
            _ExtentY        =   2037
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   8
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            SelectBlockOptions=   0
            SpreadDesigner  =   "frmOraMain.frx":87F5
            ScrollBarTrack  =   1
         End
         Begin FPSpreadADO.fpSpread fpInfo 
            Height          =   1695
            Left            =   -74880
            TabIndex        =   179
            Top             =   480
            Width           =   17415
            _Version        =   393216
            _ExtentX        =   30718
            _ExtentY        =   2990
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   9
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            SelectBlockOptions=   0
            SpreadDesigner  =   "frmOraMain.frx":8DF3
            ScrollBarTrack  =   1
         End
         Begin VB.Label lblHint 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "....."
            BeginProperty Font 
               Name            =   "맑은 고딕"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   17040
            TabIndex        =   59
            Top             =   -120
            Width           =   120
         End
      End
      Begin VB.ComboBox cboSubject 
         Height          =   300
         ItemData        =   "frmOraMain.frx":9464
         Left            =   -72600
         List            =   "frmOraMain.frx":9486
         TabIndex        =   144
         Top             =   1260
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtTableSrchKor 
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
         Left            =   -72960
         TabIndex        =   142
         Top             =   5580
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox txtTabMapping 
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
         Left            =   -74400
         Locked          =   -1  'True
         TabIndex        =   141
         Top             =   4560
         Visible         =   0   'False
         Width           =   6555
      End
      Begin VB.CommandButton Command5 
         Caption         =   "copy"
         Height          =   315
         Left            =   -67800
         Style           =   1  '그래픽
         TabIndex        =   140
         Top             =   4560
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "copy"
         Height          =   310
         Left            =   -63180
         Style           =   1  '그래픽
         TabIndex        =   139
         Top             =   3600
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.TextBox txtMapping 
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
         Left            =   -74460
         TabIndex        =   138
         Top             =   3600
         Visible         =   0   'False
         Width           =   11235
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   4575
         Left            =   -72720
         TabIndex        =   109
         Top             =   5640
         Visible         =   0   'False
         Width           =   11535
         Begin VB.TextBox txtLdgrNo 
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
            Left            =   900
            TabIndex        =   111
            Top             =   60
            Width           =   3375
         End
         Begin VB.CommandButton cmdSearchLdgr 
            Caption         =   "조회"
            Height          =   375
            Left            =   4380
            TabIndex        =   110
            Top             =   0
            Width           =   1575
         End
         Begin TabDlg.SSTab SSTab9 
            Height          =   13215
            Left            =   240
            TabIndex        =   112
            Top             =   360
            Width           =   20175
            _ExtentX        =   35586
            _ExtentY        =   23310
            _Version        =   393216
            TabHeight       =   520
            TabCaption(0)   =   "Tab 0"
            TabPicture(0)   =   "frmOraMain.frx":94B0
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label20"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label19"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label11"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Label10"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "fpTLNE_ADAMTCALC"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "fpTLNG_DLY"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "fpTLNE_LDGRADHIST"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "fpTLNI_LNLDGR"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).ControlCount=   8
            TabCaption(1)   =   "Tab 1"
            TabPicture(1)   =   "frmOraMain.frx":94CC
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "fpTLNG_LDGRTRHIST"
            Tab(1).Control(1)=   "fpTLNG_TRAMTCALC"
            Tab(1).Control(2)=   "Label17"
            Tab(1).Control(3)=   "Label18"
            Tab(1).ControlCount=   4
            TabCaption(2)   =   "Tab 2"
            TabPicture(2)   =   "frmOraMain.frx":94E8
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "fpTLNS_RFND"
            Tab(2).Control(1)=   "fpTLNS_RFNDTRHIST"
            Tab(2).Control(2)=   "fpTLNS_RFAMTCALC"
            Tab(2).Control(3)=   "fpTLNS_SPCLLDGR"
            Tab(2).Control(4)=   "fpTLNS_SPCLTRHIST"
            Tab(2).Control(5)=   "fpTLNS_SPAMTCALC"
            Tab(2).Control(6)=   "Label21"
            Tab(2).Control(7)=   "Label22"
            Tab(2).Control(8)=   "Label23"
            Tab(2).Control(9)=   "Label24"
            Tab(2).Control(10)=   "Label25"
            Tab(2).Control(11)=   "Label26"
            Tab(2).ControlCount=   12
            Begin FPSpreadADO.fpSpread fpTLNI_LNLDGR 
               Height          =   915
               Left            =   120
               TabIndex        =   113
               Top             =   660
               Width           =   19815
               _Version        =   393216
               _ExtentX        =   34951
               _ExtentY        =   1614
               _StockProps     =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   200
               MaxRows         =   1
               RetainSelBlock  =   0   'False
               ScrollBarExtMode=   -1  'True
               SelectBlockOptions=   0
               SpreadDesigner  =   "frmOraMain.frx":9504
               ScrollBarTrack  =   1
            End
            Begin FPSpreadADO.fpSpread fpTLNE_LDGRADHIST 
               Height          =   4455
               Left            =   120
               TabIndex        =   114
               Top             =   1920
               Width           =   19815
               _Version        =   393216
               _ExtentX        =   34951
               _ExtentY        =   7858
               _StockProps     =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   200
               MaxRows         =   1
               RetainSelBlock  =   0   'False
               ScrollBarExtMode=   -1  'True
               SelectBlockOptions=   0
               SpreadDesigner  =   "frmOraMain.frx":B7F3
               ScrollBarTrack  =   1
            End
            Begin FPSpreadADO.fpSpread fpTLNG_DLY 
               Height          =   2955
               Left            =   120
               TabIndex        =   115
               Top             =   10140
               Width           =   19815
               _Version        =   393216
               _ExtentX        =   34951
               _ExtentY        =   5212
               _StockProps     =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   200
               MaxRows         =   1
               RetainSelBlock  =   0   'False
               ScrollBarExtMode=   -1  'True
               SelectBlockOptions=   0
               SpreadDesigner  =   "frmOraMain.frx":DAE2
               ScrollBarTrack  =   1
            End
            Begin FPSpreadADO.fpSpread fpTLNG_LDGRTRHIST 
               Height          =   4335
               Left            =   -74820
               TabIndex        =   116
               Top             =   720
               Width           =   19815
               _Version        =   393216
               _ExtentX        =   34951
               _ExtentY        =   7646
               _StockProps     =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   200
               MaxRows         =   1
               RetainSelBlock  =   0   'False
               ScrollBarExtMode=   -1  'True
               SelectBlockOptions=   0
               SpreadDesigner  =   "frmOraMain.frx":FDD1
               ScrollBarTrack  =   1
            End
            Begin FPSpreadADO.fpSpread fpTLNG_TRAMTCALC 
               Height          =   3195
               Left            =   -74820
               TabIndex        =   117
               Top             =   5580
               Width           =   19815
               _Version        =   393216
               _ExtentX        =   34951
               _ExtentY        =   5636
               _StockProps     =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   200
               MaxRows         =   1
               RetainSelBlock  =   0   'False
               ScrollBarExtMode=   -1  'True
               SelectBlockOptions=   0
               SpreadDesigner  =   "frmOraMain.frx":120C0
               ScrollBarTrack  =   1
            End
            Begin FPSpreadADO.fpSpread fpTLNE_ADAMTCALC 
               Height          =   3015
               Left            =   120
               TabIndex        =   118
               Top             =   6720
               Width           =   19815
               _Version        =   393216
               _ExtentX        =   34951
               _ExtentY        =   5318
               _StockProps     =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   200
               MaxRows         =   1
               RetainSelBlock  =   0   'False
               ScrollBarExtMode=   -1  'True
               SelectBlockOptions=   0
               SpreadDesigner  =   "frmOraMain.frx":143AF
               ScrollBarTrack  =   1
            End
            Begin FPSpreadADO.fpSpread fpTLNS_RFND 
               Height          =   1515
               Left            =   -74940
               TabIndex        =   119
               Top             =   600
               Width           =   19815
               _Version        =   393216
               _ExtentX        =   34951
               _ExtentY        =   2672
               _StockProps     =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   200
               MaxRows         =   1
               RetainSelBlock  =   0   'False
               ScrollBarExtMode=   -1  'True
               SelectBlockOptions=   0
               SpreadDesigner  =   "frmOraMain.frx":1669E
               ScrollBarTrack  =   1
            End
            Begin FPSpreadADO.fpSpread fpTLNS_RFNDTRHIST 
               Height          =   1695
               Left            =   -74880
               TabIndex        =   120
               Top             =   4680
               Width           =   19815
               _Version        =   393216
               _ExtentX        =   34951
               _ExtentY        =   2990
               _StockProps     =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   200
               MaxRows         =   1
               RetainSelBlock  =   0   'False
               ScrollBarExtMode=   -1  'True
               SelectBlockOptions=   0
               SpreadDesigner  =   "frmOraMain.frx":1898D
               ScrollBarTrack  =   1
            End
            Begin FPSpreadADO.fpSpread fpTLNS_RFAMTCALC 
               Height          =   1575
               Left            =   -74880
               TabIndex        =   121
               Top             =   2580
               Width           =   19815
               _Version        =   393216
               _ExtentX        =   34951
               _ExtentY        =   2778
               _StockProps     =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   200
               MaxRows         =   1
               RetainSelBlock  =   0   'False
               ScrollBarExtMode=   -1  'True
               SelectBlockOptions=   0
               SpreadDesigner  =   "frmOraMain.frx":1AC7C
               ScrollBarTrack  =   1
            End
            Begin FPSpreadADO.fpSpread fpTLNS_SPCLLDGR 
               Height          =   1275
               Left            =   -74940
               TabIndex        =   122
               Top             =   6780
               Width           =   19815
               _Version        =   393216
               _ExtentX        =   34951
               _ExtentY        =   2249
               _StockProps     =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   200
               MaxRows         =   1
               RetainSelBlock  =   0   'False
               ScrollBarExtMode=   -1  'True
               SelectBlockOptions=   0
               SpreadDesigner  =   "frmOraMain.frx":1CF6B
               ScrollBarTrack  =   1
            End
            Begin FPSpreadADO.fpSpread fpTLNS_SPCLTRHIST 
               Height          =   2175
               Left            =   -74940
               TabIndex        =   123
               Top             =   8460
               Width           =   19815
               _Version        =   393216
               _ExtentX        =   34951
               _ExtentY        =   3836
               _StockProps     =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   200
               MaxRows         =   1
               RetainSelBlock  =   0   'False
               ScrollBarExtMode=   -1  'True
               SelectBlockOptions=   0
               SpreadDesigner  =   "frmOraMain.frx":1F25A
               ScrollBarTrack  =   1
            End
            Begin FPSpreadADO.fpSpread fpTLNS_SPAMTCALC 
               Height          =   2175
               Left            =   -74940
               TabIndex        =   124
               Top             =   10920
               Width           =   19815
               _Version        =   393216
               _ExtentX        =   34951
               _ExtentY        =   3836
               _StockProps     =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   200
               MaxRows         =   1
               RetainSelBlock  =   0   'False
               ScrollBarExtMode=   -1  'True
               SelectBlockOptions=   0
               SpreadDesigner  =   "frmOraMain.frx":21549
               ScrollBarTrack  =   1
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "기일"
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
               Left            =   120
               TabIndex        =   136
               Top             =   1680
               Width           =   390
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "연체"
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
               Left            =   120
               TabIndex        =   135
               Top             =   9900
               Width           =   390
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "여신원장"
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
               Left            =   240
               TabIndex        =   134
               Top             =   360
               Width           =   780
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "거래금액 계산"
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
               Left            =   -74820
               TabIndex        =   133
               Top             =   5340
               Width           =   1230
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "원장거래내역"
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
               Left            =   -74820
               TabIndex        =   132
               Top             =   480
               Width           =   1170
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "기일금액계산"
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
               Left            =   120
               TabIndex        =   131
               Top             =   6480
               Width           =   1170
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "채권상각내역"
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
               Left            =   -74940
               TabIndex        =   130
               Top             =   360
               Width           =   1170
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "채권상각거래내역"
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
               Left            =   -74880
               TabIndex        =   129
               Top             =   4440
               Width           =   1560
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "상각금액계산"
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
               Left            =   -74880
               TabIndex        =   128
               Top             =   2280
               Width           =   1170
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "특수채권"
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
               Left            =   -74940
               TabIndex        =   127
               Top             =   6540
               Width           =   780
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "특수거래내역"
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
               Left            =   -74940
               TabIndex        =   126
               Top             =   8220
               Width           =   1170
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "특수금액계산"
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
               Left            =   -74940
               TabIndex        =   125
               Top             =   10680
               Width           =   1170
            End
         End
         Begin VB.Label 원장번호 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "원장번호"
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
            Left            =   0
            TabIndex        =   137
            Top             =   120
            Width           =   780
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   1155
         Left            =   6660
         TabIndex        =   7
         Top             =   480
         Width           =   17055
         Begin VB.TextBox txtTableDao 
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
            Left            =   9000
            Locked          =   -1  'True
            TabIndex        =   147
            Top             =   120
            Width           =   3195
         End
         Begin VB.TextBox txtOwner 
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
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   145
            Top             =   120
            Width           =   915
         End
         Begin VB.TextBox edtTableBigo 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   12
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            IMEMode         =   8  '영문
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   18
            Top             =   480
            Width           =   13515
         End
         Begin VB.CommandButton cmdValueSelect 
            BackColor       =   &H00F9F9D5&
            Caption         =   "value조회"
            Height          =   975
            Left            =   14580
            Style           =   1  '그래픽
            TabIndex        =   17
            Top             =   120
            Width           =   2415
         End
         Begin VB.CommandButton cmdColMultiple 
            Caption         =   "컬럼 일괄등록"
            Height          =   375
            Left            =   14100
            Style           =   1  '그래픽
            TabIndex        =   16
            Top             =   420
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CommandButton cmdTableUpdate 
            Caption         =   "설명등록"
            Height          =   375
            Left            =   13140
            Style           =   1  '그래픽
            TabIndex        =   15
            Top             =   120
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox edtTableCrdate 
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
            Left            =   12720
            TabIndex        =   14
            Top             =   -60
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox edtTablenm 
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
            Left            =   4620
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   120
            Width           =   4335
         End
         Begin VB.TextBox edtTable 
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
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   120
            Width           =   3495
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "Print"
            Height          =   495
            Left            =   13680
            Style           =   1  '그래픽
            TabIndex        =   11
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdExcel 
            Caption         =   "Excel"
            Height          =   435
            Left            =   13680
            Style           =   1  '그래픽
            TabIndex        =   10
            Top             =   660
            Width           =   855
         End
         Begin VB.CommandButton cmdFavPlus 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   12240
            TabIndex        =   9
            Top             =   120
            Width           =   375
         End
         Begin VB.CommandButton cmdFavMinus 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   12660
            TabIndex        =   8
            Top             =   120
            Width           =   375
         End
         Begin VB.Label lblFavo 
            BackStyle       =   0  '투명
            Caption         =   "☆"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   12
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   12180
            MousePointer    =   99  '사용자 정의
            TabIndex        =   19
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
      End
      Begin VB.Frame Frame5 
         Height          =   14295
         Left            =   -74880
         TabIndex        =   108
         Top             =   360
         Width           =   23535
         Begin VB.ComboBox cboSplliter 
            Height          =   300
            ItemData        =   "frmOraMain.frx":23838
            Left            =   10920
            List            =   "frmOraMain.frx":23842
            TabIndex        =   195
            Text            =   " AS "
            Top             =   120
            Width           =   1935
         End
         Begin VB.CommandButton cmdAlign 
            Caption         =   ">"
            Height          =   1455
            Left            =   11640
            TabIndex        =   191
            Top             =   5040
            Width           =   615
         End
         Begin RichTextLib.RichTextBox txtOrgSQL 
            Height          =   13215
            Left            =   360
            TabIndex        =   189
            Top             =   480
            Width           =   11295
            _ExtentX        =   19923
            _ExtentY        =   23310
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"frmOraMain.frx":23850
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RichTextLib.RichTextBox txtTgtSQL 
            Height          =   13095
            Left            =   12240
            TabIndex        =   190
            Top             =   480
            Width           =   11295
            _ExtentX        =   19923
            _ExtentY        =   23098
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"frmOraMain.frx":238F5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   3855
         Left            =   -61080
         TabIndex        =   87
         Top             =   6840
         Visible         =   0   'False
         Width           =   4455
         Begin VB.PictureBox EZFTP1 
            Height          =   480
            Left            =   2280
            ScaleHeight     =   420
            ScaleWidth      =   1140
            TabIndex        =   91
            Top             =   0
            Width           =   1200
         End
         Begin VB.CommandButton cmdSock 
            Caption         =   "socket"
            Height          =   435
            Left            =   0
            TabIndex        =   90
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   435
            Left            =   4440
            TabIndex        =   89
            Top             =   0
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton Command2 
            Caption         =   "ftptest"
            Height          =   435
            Left            =   5340
            TabIndex        =   88
            Top             =   0
            Visible         =   0   'False
            Width           =   555
         End
         Begin TabDlg.SSTab SSTab4 
            Height          =   12135
            Left            =   14280
            TabIndex        =   92
            Top             =   240
            Width           =   17475
            _ExtentX        =   30824
            _ExtentY        =   21405
            _Version        =   393216
            TabsPerRow      =   7
            TabHeight       =   520
            TabCaption(0)   =   "META"
            TabPicture(0)   =   "frmOraMain.frx":2399A
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label7"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label5"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "WebBrowser1"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "cmdFindEmp"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "cmdFineNetemp"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "cmdIBKsystem"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "cmdIBKSGroup"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "txtIBKBankEmp"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "cmdEHR"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "cmdMail"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "cmdBasel"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).ControlCount=   11
            TabCaption(1)   =   "PMS"
            TabPicture(1)   =   "frmOraMain.frx":239B6
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "WebBrowser2"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "차세대 메타조회"
            TabPicture(2)   =   "frmOraMain.frx":239D2
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "WebBrowser3"
            Tab(2).Control(1)=   "Command4"
            Tab(2).Control(2)=   "Text2"
            Tab(2).ControlCount=   3
            Begin VB.CommandButton cmdBasel 
               Caption         =   "바젤전용"
               Height          =   555
               Left            =   15540
               TabIndex        =   102
               Top             =   540
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.TextBox Text2 
               Height          =   315
               Left            =   -71760
               TabIndex        =   101
               Text            =   "db915b // ibkit012"
               Top             =   540
               Visible         =   0   'False
               Width           =   3735
            End
            Begin VB.CommandButton Command4 
               Caption         =   "Refresh"
               Height          =   555
               Left            =   -74880
               Style           =   1  '그래픽
               TabIndex        =   100
               Top             =   420
               Width           =   2775
            End
            Begin VB.CommandButton cmdMail 
               Caption         =   "Mail System"
               Height          =   375
               Left            =   11880
               Style           =   1  '그래픽
               TabIndex        =   99
               Top             =   480
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.CommandButton cmdEHR 
               Caption         =   "e-HR"
               Height          =   375
               Left            =   13140
               Style           =   1  '그래픽
               TabIndex        =   98
               Top             =   480
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.TextBox txtIBKBankEmp 
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
               Left            =   2760
               TabIndex        =   97
               Top             =   540
               Visible         =   0   'False
               Width           =   2355
            End
            Begin VB.CommandButton cmdIBKSGroup 
               Caption         =   "그룹웨어"
               Height          =   375
               Left            =   10620
               Style           =   1  '그래픽
               TabIndex        =   96
               Top             =   480
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.CommandButton cmdIBKsystem 
               Caption         =   "직원찾기"
               Height          =   375
               Left            =   9240
               Style           =   1  '그래픽
               TabIndex        =   95
               Top             =   480
               Visible         =   0   'False
               Width           =   1275
            End
            Begin VB.CommandButton cmdFineNetemp 
               Caption         =   "직원찾기홈"
               Height          =   375
               Left            =   6360
               Style           =   1  '그래픽
               TabIndex        =   94
               Top             =   480
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.CommandButton cmdFindEmp 
               Caption         =   "직원조회"
               Height          =   375
               Left            =   5160
               Style           =   1  '그래픽
               TabIndex        =   93
               Top             =   480
               Visible         =   0   'False
               Width           =   1095
            End
            Begin SHDocVwCtl.WebBrowser WebBrowser1 
               Height          =   9015
               Left            =   120
               TabIndex        =   103
               Top             =   480
               Width           =   17175
               ExtentX         =   30295
               ExtentY         =   15901
               ViewMode        =   0
               Offline         =   0
               Silent          =   0
               RegisterAsBrowser=   0
               RegisterAsDropTarget=   1
               AutoArrange     =   0   'False
               NoClientEdge    =   0   'False
               AlignLeft       =   0   'False
               NoWebView       =   0   'False
               HideFileNames   =   0   'False
               SingleClick     =   0   'False
               SingleSelection =   0   'False
               NoFolders       =   0   'False
               Transparent     =   0   'False
               ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
               Location        =   "http:///"
            End
            Begin SHDocVwCtl.WebBrowser WebBrowser3 
               Height          =   11055
               Left            =   -74880
               TabIndex        =   104
               Top             =   1020
               Width           =   17235
               ExtentX         =   30401
               ExtentY         =   19500
               ViewMode        =   0
               Offline         =   0
               Silent          =   0
               RegisterAsBrowser=   0
               RegisterAsDropTarget=   1
               AutoArrange     =   0   'False
               NoClientEdge    =   0   'False
               AlignLeft       =   0   'False
               NoWebView       =   0   'False
               HideFileNames   =   0   'False
               SingleClick     =   0   'False
               SingleSelection =   0   'False
               NoFolders       =   0   'False
               Transparent     =   0   'False
               ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
               Location        =   "http:///"
            End
            Begin SHDocVwCtl.WebBrowser WebBrowser2 
               Height          =   11415
               Left            =   -74880
               TabIndex        =   105
               Top             =   660
               Width           =   17235
               ExtentX         =   30401
               ExtentY         =   20135
               ViewMode        =   0
               Offline         =   0
               Silent          =   0
               RegisterAsBrowser=   0
               RegisterAsDropTarget=   1
               AutoArrange     =   0   'False
               NoClientEdge    =   0   'False
               AlignLeft       =   0   'False
               NoWebView       =   0   'False
               HideFileNames   =   0   'False
               SingleClick     =   0   'False
               SingleSelection =   0   'False
               NoFolders       =   0   'False
               Transparent     =   0   'False
               ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
               Location        =   "http:///"
            End
            Begin VB.Label Label5 
               Appearance      =   0  '평면
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  '투명
               Caption         =   "IBK System"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   8040
               TabIndex        =   107
               Top             =   600
               Visible         =   0   'False
               Width           =   1050
            End
            Begin VB.Label Label7 
               Appearance      =   0  '평면
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  '투명
               Caption         =   "기업은행 직원명, 직원번호"
               BeginProperty Font 
                  Name            =   "굴림"
                  Size            =   9.75
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   360
               TabIndex        =   106
               Top             =   600
               Visible         =   0   'False
               Width           =   2325
            End
         End
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   17400
         Top             =   -120
      End
      Begin VB.CommandButton cmdNeed 
         Caption         =   "필수생성"
         Height          =   315
         Left            =   18420
         TabIndex        =   6
         Top             =   -60
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "컬럼단위 검색"
         Height          =   435
         Left            =   4800
         TabIndex        =   2
         Top             =   9600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin TabDlg.SSTab SSTab8 
         Height          =   10635
         Left            =   0
         TabIndex        =   3
         Top             =   420
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   18759
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "테이블 검색"
         TabPicture(0)   =   "frmOraMain.frx":239EE
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fpTable"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "즐겨찾기"
         TabPicture(1)   =   "frmOraMain.frx":23A0A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fpFavorite"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "최근 검색"
         TabPicture(2)   =   "frmOraMain.frx":23A26
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fpNear"
         Tab(2).ControlCount=   1
         Begin VB.Frame Frame1 
            BackColor       =   &H80000004&
            Height          =   735
            Left            =   120
            TabIndex        =   79
            Top             =   360
            Width           =   6375
            Begin VB.CheckBox chkASIS 
               Caption         =   "AS-IS"
               Height          =   255
               Left            =   4440
               TabIndex        =   84
               Top             =   1080
               Visible         =   0   'False
               Width           =   1095
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
               IMEMode         =   10  '한글 
               Left            =   1740
               TabIndex        =   81
               Top             =   240
               Width           =   3375
            End
            Begin VB.CommandButton cmdTableSrch 
               Caption         =   "조회"
               Height          =   555
               Left            =   5220
               Style           =   1  '그래픽
               TabIndex        =   82
               Top             =   180
               Width           =   1035
            End
            Begin VB.TextBox txtTableSrchBigo 
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
               Left            =   960
               TabIndex        =   83
               Top             =   420
               Visible         =   0   'False
               Width           =   3375
            End
            Begin VB.ComboBox cboShema 
               Height          =   300
               ItemData        =   "frmOraMain.frx":23A42
               Left            =   180
               List            =   "frmOraMain.frx":23A44
               TabIndex        =   80
               Top             =   240
               Width           =   1515
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "설명"
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
               Left            =   60
               TabIndex        =   85
               Top             =   1200
               Visible         =   0   'False
               Width           =   390
            End
         End
         Begin FPSpreadADO.fpSpread fpTable 
            Height          =   9375
            Left            =   120
            TabIndex        =   4
            Top             =   1140
            Width           =   6375
            _Version        =   393216
            _ExtentX        =   11245
            _ExtentY        =   16536
            _StockProps     =   64
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
            MaxCols         =   5
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            SelectBlockOptions=   0
            ShadowColor     =   12648447
            ShadowDark      =   12632256
            SpreadDesigner  =   "frmOraMain.frx":23A46
            ScrollBarTrack  =   1
         End
         Begin FPSpreadADO.fpSpread fpFavorite 
            Height          =   8655
            Left            =   -74940
            TabIndex        =   5
            Top             =   1200
            Width           =   6435
            _Version        =   393216
            _ExtentX        =   11351
            _ExtentY        =   15266
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ColsFrozen      =   1
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
            MaxCols         =   5
            MaxRows         =   0
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            ShadowColor     =   12648447
            ShadowDark      =   12632256
            SpreadDesigner  =   "frmOraMain.frx":23EFF
            ScrollBarTrack  =   1
         End
         Begin FPSpreadADO.fpSpread fpNear 
            Height          =   9495
            Left            =   -74940
            TabIndex        =   178
            Top             =   1140
            Width           =   6435
            _Version        =   393216
            _ExtentX        =   11351
            _ExtentY        =   16748
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ColsFrozen      =   1
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
            MaxCols         =   5
            MaxRows         =   0
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            ShadowColor     =   12648447
            ShadowDark      =   12632256
            SpreadDesigner  =   "frmOraMain.frx":243CB
            ScrollBarTrack  =   1
         End
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   9315
         Left            =   6660
         TabIndex        =   20
         Top             =   1680
         Width           =   17115
         _ExtentX        =   30189
         _ExtentY        =   16431
         _Version        =   393216
         Tabs            =   6
         Tab             =   3
         TabsPerRow      =   6
         TabHeight       =   520
         BackColor       =   16777215
         TabCaption(0)   =   "테이블정보"
         TabPicture(0)   =   "frmOraMain.frx":24897
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "frmNeed"
         Tab(0).Control(1)=   "fpCol"
         Tab(0).Control(2)=   "picMatch"
         Tab(0).Control(3)=   "optColor(2)"
         Tab(0).Control(4)=   "optColor(1)"
         Tab(0).Control(5)=   "optColor(0)"
         Tab(0).Control(6)=   "cmdBookMakr"
         Tab(0).Control(7)=   "txtSrchId"
         Tab(0).Control(8)=   "txtColNm"
         Tab(0).Control(9)=   "txtColId"
         Tab(0).Control(10)=   "txtCodeDomain"
         Tab(0).Control(11)=   "cmdColUpdate"
         Tab(0).Control(12)=   "frameCode"
         Tab(0).Control(13)=   "txtColBigo"
         Tab(0).Control(14)=   "txtCodeSrchNm"
         Tab(0).ControlCount=   15
         TabCaption(1)   =   "기본SQL"
         TabPicture(1)   =   "frmOraMain.frx":248B3
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SSTab5"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "EXECUTE SQL "
         TabPicture(2)   =   "frmOraMain.frx":248CF
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label3"
         Tab(2).Control(1)=   "cmdFreeSQL"
         Tab(2).Control(2)=   "txtReqSQL"
         Tab(2).ControlCount=   3
         TabCaption(3)   =   "적용 SQL"
         TabPicture(3)   =   "frmOraMain.frx":248EB
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "txtSQL"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "이행데이터"
         TabPicture(4)   =   "frmOraMain.frx":24907
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "txtMigSQL"
         Tab(4).Control(1)=   "Frame9"
         Tab(4).Control(2)=   "cmdCntSerarch"
         Tab(4).Control(3)=   "Frame8"
         Tab(4).Control(4)=   "Frame7"
         Tab(4).ControlCount=   5
         TabCaption(5)   =   "PO"
         TabPicture(5)   =   "frmOraMain.frx":24923
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "SSTab10"
         Tab(5).ControlCount=   1
         Begin TabDlg.SSTab SSTab10 
            Height          =   8535
            Left            =   -74880
            TabIndex        =   192
            Top             =   480
            Width           =   16695
            _ExtentX        =   29448
            _ExtentY        =   15055
            _Version        =   393216
            Tabs            =   5
            TabsPerRow      =   5
            TabHeight       =   520
            TabCaption(0)   =   "Delete"
            TabPicture(0)   =   "frmOraMain.frx":2493F
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "txtMapper"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "txtPoDelete"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "Tab 1"
            Tab(1).ControlEnabled=   0   'False
            Tab(1).ControlCount=   0
            TabCaption(2)   =   "Tab 2"
            Tab(2).ControlEnabled=   0   'False
            Tab(2).ControlCount=   0
            TabCaption(3)   =   "Tab 3"
            Tab(3).ControlEnabled=   0   'False
            Tab(3).ControlCount=   0
            TabCaption(4)   =   "Tab 4"
            Tab(4).ControlEnabled=   0   'False
            Tab(4).ControlCount=   0
            Begin RichTextLib.RichTextBox txtPoDelete 
               Height          =   3540
               Left            =   120
               TabIndex        =   193
               Top             =   600
               Width           =   16335
               _ExtentX        =   28813
               _ExtentY        =   6244
               _Version        =   393217
               ScrollBars      =   2
               Appearance      =   0
               TextRTF         =   $"frmOraMain.frx":2495B
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin RichTextLib.RichTextBox txtMapper 
               Height          =   4140
               Left            =   120
               TabIndex        =   194
               Top             =   4200
               Width           =   16335
               _ExtentX        =   28813
               _ExtentY        =   7303
               _Version        =   393217
               ScrollBars      =   2
               Appearance      =   0
               TextRTF         =   $"frmOraMain.frx":24A00
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin VB.TextBox txtMigSQL 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   12
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5595
            IMEMode         =   8  '영문
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   180
            Top             =   3540
            Width           =   16815
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00ECD9DD&
            Caption         =   "Mig -> 개발"
            Height          =   2775
            Left            =   -64920
            TabIndex        =   170
            Top             =   540
            Width           =   6855
            Begin VB.CheckBox chkNoExe 
               BackColor       =   &H00ECD9DD&
               Caption         =   "실행 하지 않고  SQL만 생성"
               Height          =   195
               Left            =   3840
               TabIndex        =   188
               Top             =   1440
               Width           =   2835
            End
            Begin VB.CommandButton cmdMig 
               BackColor       =   &H00F9F9D5&
               Caption         =   "이관"
               Height          =   555
               Left            =   4020
               Style           =   1  '그래픽
               TabIndex        =   176
               Top             =   780
               Width           =   2295
            End
            Begin VB.CheckBox chtTruncate 
               BackColor       =   &H00ECD9DD&
               Caption         =   "개발 DB Delete All 후 이관"
               Height          =   195
               Left            =   3840
               TabIndex        =   175
               Top             =   420
               Width           =   2835
            End
            Begin VB.TextBox txtMoveCnt 
               Alignment       =   1  '오른쪽 맞춤
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
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   173
               Top             =   840
               Width           =   1335
            End
            Begin VB.OptionButton optMigSome 
               BackColor       =   &H00ECD9DD&
               Caption         =   "부분 이관"
               Height          =   195
               Left            =   300
               TabIndex        =   172
               Top             =   900
               Width           =   1095
            End
            Begin VB.OptionButton optMigAll 
               BackColor       =   &H00ECD9DD&
               Caption         =   "전체 이관"
               Height          =   195
               Left            =   300
               TabIndex        =   171
               Top             =   480
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "(WHERE ROWNUM <= )"
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
               TabIndex        =   177
               Top             =   1200
               Width           =   2130
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "건"
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
               Left            =   2880
               TabIndex        =   174
               Top             =   900
               Width           =   195
            End
         End
         Begin VB.CommandButton cmdCntSerarch 
            BackColor       =   &H00F9F9D5&
            Caption         =   "건수 조회"
            Height          =   2895
            Left            =   -66360
            Style           =   1  '그래픽
            TabIndex        =   169
            Top             =   540
            Width           =   1335
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00ECD9DD&
            Caption         =   "개발 DB"
            Height          =   1395
            Left            =   -74820
            TabIndex        =   158
            Top             =   2040
            Width           =   8355
            Begin VB.TextBox txtTgtCnt 
               Alignment       =   1  '오른쪽 맞춤
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
               Left            =   6540
               Locked          =   -1  'True
               TabIndex        =   168
               Top             =   360
               Width           =   1575
            End
            Begin VB.TextBox txtTgtTableNm 
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
               Left            =   300
               Locked          =   -1  'True
               TabIndex        =   166
               Top             =   720
               Width           =   5595
            End
            Begin VB.TextBox txtTgtTableId 
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
               Left            =   1260
               Locked          =   -1  'True
               TabIndex        =   165
               Top             =   360
               Width           =   4635
            End
            Begin VB.TextBox txtTgtOwner 
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
               Left            =   300
               Locked          =   -1  'True
               TabIndex        =   164
               Top             =   360
               Width           =   915
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "건수"
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
               Left            =   6060
               TabIndex        =   167
               Top             =   420
               Width           =   390
            End
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00ECD9DD&
            Caption         =   "Mig DB"
            Height          =   1455
            Left            =   -74820
            TabIndex        =   157
            Top             =   480
            Width           =   8355
            Begin VB.TextBox txtMigCnt 
               Alignment       =   1  '오른쪽 맞춤
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
               Left            =   6480
               Locked          =   -1  'True
               TabIndex        =   163
               Top             =   420
               Width           =   1575
            End
            Begin VB.TextBox txtMigTableNm 
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
               Left            =   240
               Locked          =   -1  'True
               TabIndex        =   161
               Top             =   780
               Width           =   5595
            End
            Begin VB.TextBox txtMigTableID 
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
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   160
               Top             =   420
               Width           =   4635
            End
            Begin VB.TextBox txtMigOwner 
               Appearance      =   0  '평면
               BackColor       =   &H00FFFFFF&
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
               Left            =   240
               TabIndex        =   159
               Top             =   420
               Width           =   915
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "건수"
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
               Left            =   6000
               TabIndex        =   162
               Top             =   480
               Width           =   390
            End
         End
         Begin VB.TextBox txtCodeSrchNm 
            Appearance      =   0  '평면
            BackColor       =   &H0080C0FF&
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
            Left            =   -64140
            TabIndex        =   146
            Top             =   420
            Width           =   4335
         End
         Begin VB.TextBox txtColBigo 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   12
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            IMEMode         =   8  '영문
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   47
            Top             =   8520
            Width           =   16755
         End
         Begin VB.Frame frameCode 
            BackColor       =   &H00FFFFFF&
            Height          =   6195
            Left            =   -72300
            TabIndex        =   37
            Top             =   1920
            Width           =   12255
            Begin VB.TextBox txtcolname 
               Height          =   375
               Left            =   11160
               TabIndex        =   219
               Top             =   5520
               Width           =   735
            End
            Begin VB.TextBox txtCodeWhere 
               Appearance      =   0  '평면
               BackColor       =   &H0080FFFF&
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   12
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               IMEMode         =   8  '영문
               Index           =   3
               Left            =   360
               TabIndex        =   218
               Top             =   4680
               Width           =   11715
            End
            Begin VB.TextBox txtCodeWhere 
               Appearance      =   0  '평면
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   12
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               IMEMode         =   8  '영문
               Index           =   0
               Left            =   360
               TabIndex        =   45
               Top             =   5700
               Width           =   10275
            End
            Begin VB.TextBox txtCodeWhere 
               Appearance      =   0  '평면
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   12
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               IMEMode         =   8  '영문
               Index           =   1
               Left            =   360
               TabIndex        =   44
               Top             =   5400
               Width           =   10275
            End
            Begin VB.TextBox txtCodeWhere 
               Appearance      =   0  '평면
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   12
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               IMEMode         =   8  '영문
               Index           =   2
               Left            =   360
               TabIndex        =   43
               Top             =   5100
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
               Height          =   360
               IMEMode         =   10  '한글 
               Left            =   1020
               TabIndex        =   42
               Top             =   300
               Width           =   2715
            End
            Begin VB.CommandButton cmdBookReset 
               BackColor       =   &H80000016&
               Caption         =   "BookMark Reset"
               Height          =   375
               Left            =   3840
               Style           =   1  '그래픽
               TabIndex        =   41
               Top             =   300
               Width           =   1875
            End
            Begin VB.CommandButton cmdADD 
               Caption         =   "추가"
               Height          =   375
               Left            =   6240
               TabIndex        =   40
               Top             =   180
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.CommandButton cmdDel 
               Caption         =   "삭제"
               Height          =   375
               Left            =   7080
               TabIndex        =   39
               Top             =   180
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.CommandButton cmdSave 
               Caption         =   "저장"
               Height          =   555
               Left            =   8100
               TabIndex        =   38
               Top             =   180
               Visible         =   0   'False
               Width           =   1395
            End
            Begin FPSpreadADO.fpSpread fpcode 
               Height          =   3855
               Left            =   300
               TabIndex        =   46
               Top             =   780
               Width           =   11595
               _Version        =   393216
               _ExtentX        =   20452
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
               SpreadDesigner  =   "frmOraMain.frx":24AA5
               ScrollBarTrack  =   1
            End
            Begin VB.Image Image1 
               Height          =   405
               Left            =   11100
               Picture         =   "frmOraMain.frx":25003
               Stretch         =   -1  'True
               Top             =   240
               Width           =   510
            End
         End
         Begin VB.CommandButton cmdColUpdate 
            Caption         =   "수정"
            Height          =   310
            Left            =   -59940
            Style           =   1  '그래픽
            TabIndex        =   36
            Top             =   8190
            Visible         =   0   'False
            Width           =   1755
         End
         Begin VB.TextBox txtCodeDomain 
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
            Left            =   -66540
            TabIndex        =   35
            Top             =   8190
            Width           =   1275
         End
         Begin VB.TextBox txtColId 
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
            Left            =   -74880
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   8190
            Width           =   3315
         End
         Begin VB.TextBox txtColNm 
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
            Left            =   -71520
            TabIndex        =   33
            Top             =   8190
            Width           =   4935
         End
         Begin VB.TextBox txtReqSQL 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   12
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   8355
            IMEMode         =   8  '영문
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   32
            Top             =   900
            Width           =   16875
         End
         Begin VB.CommandButton cmdFreeSQL 
            Caption         =   "조회"
            Height          =   435
            Left            =   -60420
            Style           =   1  '그래픽
            TabIndex        =   31
            Top             =   420
            Width           =   2235
         End
         Begin VB.TextBox txtSrchId 
            Appearance      =   0  '평면
            BackColor       =   &H00F2E6A8&
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
            Left            =   -74460
            TabIndex        =   30
            Top             =   397
            Width           =   3195
         End
         Begin VB.CommandButton cmdBookMakr 
            BackColor       =   &H80000016&
            Caption         =   "BookMark Reset"
            Height          =   375
            Left            =   -71160
            Style           =   1  '그래픽
            TabIndex        =   29
            Top             =   360
            Width           =   1875
         End
         Begin VB.OptionButton optColor 
            Caption         =   "Cyan"
            Height          =   195
            Index           =   0
            Left            =   -69180
            TabIndex        =   28
            Top             =   480
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton optColor 
            Caption         =   "Yellow"
            Height          =   195
            Index           =   1
            Left            =   -67200
            TabIndex        =   27
            Top             =   480
            Width           =   915
         End
         Begin VB.OptionButton optColor 
            Caption         =   "Brown"
            Height          =   195
            Index           =   2
            Left            =   -68220
            TabIndex        =   26
            Top             =   480
            Width           =   915
         End
         Begin VB.PictureBox picMatch 
            Appearance      =   0  '평면
            BackColor       =   &H80000005&
            BorderStyle     =   0  '없음
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   -74400
            ScaleHeight     =   315
            ScaleWidth      =   4155
            TabIndex        =   24
            Top             =   1260
            Visible         =   0   'False
            Width           =   4155
            Begin VB.Label lblMatch 
               BackStyle       =   0  '투명
               Caption         =   "Matched 3 columns....."
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   12
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   195
               Left            =   300
               TabIndex        =   25
               Top             =   60
               Width           =   3195
            End
         End
         Begin TabDlg.SSTab SSTab5 
            Height          =   8835
            Left            =   -74880
            TabIndex        =   48
            Top             =   420
            Width           =   16935
            _ExtentX        =   29871
            _ExtentY        =   15584
            _Version        =   393216
            Tabs            =   6
            TabsPerRow      =   7
            TabHeight       =   520
            TabCaption(0)   =   "SELECT"
            TabPicture(0)   =   "frmOraMain.frx":279CB
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label4"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "txtSelectSQL"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "chkAS"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Command6"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "txtAlias"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).ControlCount=   5
            TabCaption(1)   =   "INSERT"
            TabPicture(1)   =   "frmOraMain.frx":279E7
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "txtInsertSQL"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "UPDATE"
            TabPicture(2)   =   "frmOraMain.frx":27A03
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "txUpdateSQL"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "INSERT/SELECT"
            TabPicture(3)   =   "frmOraMain.frx":27A1F
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "txtInsSelect"
            Tab(3).ControlCount=   1
            TabCaption(4)   =   "SELECT_CODENM"
            TabPicture(4)   =   "frmOraMain.frx":27A3B
            Tab(4).ControlEnabled=   0   'False
            Tab(4).Control(0)=   "txtSQLcdnm"
            Tab(4).ControlCount=   1
            TabCaption(5)   =   "PO insert"
            TabPicture(5)   =   "frmOraMain.frx":27A57
            Tab(5).ControlEnabled=   0   'False
            Tab(5).Control(0)=   "chkComment"
            Tab(5).Control(1)=   "txtIQuery"
            Tab(5).ControlCount=   2
            Begin RichTextLib.RichTextBox txtIQuery 
               Height          =   7875
               Left            =   -74880
               TabIndex        =   182
               Top             =   420
               Width           =   16455
               _ExtentX        =   29025
               _ExtentY        =   13891
               _Version        =   393217
               ScrollBars      =   2
               TextRTF         =   $"frmOraMain.frx":27A73
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.CheckBox chkComment 
               Caption         =   "주석"
               Height          =   195
               Left            =   -74580
               TabIndex        =   181
               Top             =   8400
               Value           =   1  '확인
               Width           =   1215
            End
            Begin VB.TextBox txtAlias 
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
               Left            =   13020
               TabIndex        =   51
               Top             =   8340
               Width           =   495
            End
            Begin VB.CommandButton Command6 
               Caption         =   "SQL 재생성"
               Height          =   375
               Left            =   13680
               TabIndex        =   50
               Top             =   8280
               Width           =   1395
            End
            Begin VB.CheckBox chkAS 
               Caption         =   "AS 논리명"
               Height          =   195
               Left            =   11160
               TabIndex        =   49
               Top             =   8400
               Width           =   1215
            End
            Begin RichTextLib.RichTextBox txtSQLcdnm 
               Height          =   7875
               Left            =   -74820
               TabIndex        =   183
               Top             =   540
               Width           =   16455
               _ExtentX        =   29025
               _ExtentY        =   13891
               _Version        =   393217
               ScrollBars      =   2
               TextRTF         =   $"frmOraMain.frx":27B18
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin RichTextLib.RichTextBox txtInsSelect 
               Height          =   7875
               Left            =   -74880
               TabIndex        =   184
               Top             =   420
               Width           =   16455
               _ExtentX        =   29025
               _ExtentY        =   13891
               _Version        =   393217
               ScrollBars      =   2
               TextRTF         =   $"frmOraMain.frx":27BBD
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin RichTextLib.RichTextBox txUpdateSQL 
               Height          =   7875
               Left            =   -74760
               TabIndex        =   185
               Top             =   480
               Width           =   16455
               _ExtentX        =   29025
               _ExtentY        =   13891
               _Version        =   393217
               ScrollBars      =   2
               TextRTF         =   $"frmOraMain.frx":27C62
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin RichTextLib.RichTextBox txtInsertSQL 
               Height          =   7875
               Left            =   -74820
               TabIndex        =   186
               Top             =   480
               Width           =   16455
               _ExtentX        =   29025
               _ExtentY        =   13891
               _Version        =   393217
               ScrollBars      =   2
               TextRTF         =   $"frmOraMain.frx":27D07
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin RichTextLib.RichTextBox txtSelectSQL 
               Height          =   7815
               Left            =   120
               TabIndex        =   187
               Top             =   420
               Width           =   16575
               _ExtentX        =   29236
               _ExtentY        =   13785
               _Version        =   393217
               ScrollBars      =   2
               TextRTF         =   $"frmOraMain.frx":27DAC
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "굴림체"
                  Size            =   9
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "Alias"
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
               Left            =   12540
               TabIndex        =   52
               Top             =   8400
               Width           =   450
            End
         End
         Begin FPSpreadADO.fpSpread fpCol 
            Height          =   7395
            Left            =   -74880
            TabIndex        =   53
            Top             =   780
            Width           =   16815
            _Version        =   393216
            _ExtentX        =   29660
            _ExtentY        =   13044
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
            MaxCols         =   12
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            ShadowColor     =   12648447
            SpreadDesigner  =   "frmOraMain.frx":27E51
            ScrollBarTrack  =   1
         End
         Begin VB.Frame frmNeed 
            BackColor       =   &H00FFFFFF&
            Height          =   3855
            Left            =   -69420
            TabIndex        =   21
            Top             =   1740
            Visible         =   0   'False
            Width           =   6855
            Begin VB.TextBox txtNeed 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   12
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2835
               IMEMode         =   8  '영문
               Left            =   60
               MultiLine       =   -1  'True
               ScrollBars      =   3  '양방향
               TabIndex        =   23
               Top             =   180
               Width           =   9675
            End
            Begin VB.TextBox txtLogger 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   12
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2535
               IMEMode         =   8  '영문
               Left            =   60
               MultiLine       =   -1  'True
               ScrollBars      =   3  '양방향
               TabIndex        =   22
               Top             =   3060
               Width           =   9675
            End
            Begin VB.Image Image3 
               Height          =   405
               Left            =   9720
               Picture         =   "frmOraMain.frx":28589
               Stretch         =   -1  'True
               Top             =   120
               Width           =   510
            End
         End
         Begin RichTextLib.RichTextBox txtSQL 
            Height          =   8715
            Left            =   120
            TabIndex        =   221
            Top             =   480
            Width           =   16875
            _ExtentX        =   29766
            _ExtentY        =   15372
            _Version        =   393217
            ScrollBars      =   2
            TextRTF         =   $"frmOraMain.frx":2AF51
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Fixedsys"
               Size            =   12
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "실행 ==> F5"
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
            Left            =   -74760
            TabIndex        =   54
            Top             =   600
            Width           =   1035
         End
      End
      Begin TabDlg.SSTab SSTab6 
         Height          =   13365
         Left            =   -74760
         TabIndex        =   60
         Top             =   480
         Width           =   19575
         _ExtentX        =   34528
         _ExtentY        =   23574
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   7
         TabHeight       =   520
         TabCaption(0)   =   "DicTable생성"
         TabPicture(0)   =   "frmOraMain.frx":2AFE6
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "txtDicSQL"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "코드VIEW생성"
         TabPicture(1)   =   "frmOraMain.frx":2B002
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Text1"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "쿼리변환"
         TabPicture(2)   =   "frmOraMain.frx":2B01E
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtNewTabName"
         Tab(2).Control(1)=   "txtAsisAlias"
         Tab(2).Control(2)=   "cmdConvert"
         Tab(2).Control(3)=   "txtOldTabName"
         Tab(2).Control(4)=   "txtNewSQL"
         Tab(2).Control(5)=   "txtOldSQL"
         Tab(2).ControlCount=   6
         TabCaption(3)   =   "Java SQL "
         TabPicture(3)   =   "frmOraMain.frx":2B03A
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "txtQeryAppend"
         Tab(3).Control(1)=   "cmdDel1"
         Tab(3).Control(2)=   "txtConvSQL"
         Tab(3).Control(3)=   "txtDel2"
         Tab(3).Control(4)=   "cmdDel2"
         Tab(3).Control(5)=   "Text3"
         Tab(3).Control(6)=   "cmdDel3"
         Tab(3).Control(7)=   "cmdclear2"
         Tab(3).Control(8)=   "cmdCopy"
         Tab(3).ControlCount=   9
         Begin VB.TextBox txtNewTabName 
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
            Left            =   -66240
            TabIndex        =   77
            Top             =   7320
            Width           =   1815
         End
         Begin VB.TextBox txtAsisAlias 
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
            Left            =   -68160
            TabIndex        =   76
            Top             =   7680
            Width           =   1455
         End
         Begin VB.CommandButton cmdConvert 
            Caption         =   "변환"
            Height          =   555
            Left            =   -65760
            Style           =   1  '그래픽
            TabIndex        =   75
            Top             =   7740
            Width           =   1335
         End
         Begin VB.TextBox txtOldTabName 
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
            Left            =   -68160
            TabIndex        =   74
            Top             =   7320
            Width           =   1815
         End
         Begin VB.TextBox txtNewSQL 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   12
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4515
            IMEMode         =   8  '영문
            Left            =   -74760
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   73
            Top             =   8520
            Width           =   16875
         End
         Begin VB.TextBox txtOldSQL 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   12
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5655
            IMEMode         =   8  '영문
            Left            =   -74760
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   72
            Top             =   1080
            Width           =   16875
         End
         Begin VB.TextBox txtDicSQL 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   12
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   11175
            IMEMode         =   8  '영문
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   71
            Text            =   "frmOraMain.frx":2B056
            Top             =   600
            Width           =   16875
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   12
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   11175
            IMEMode         =   8  '영문
            Left            =   -74880
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   70
            Text            =   "frmOraMain.frx":2B2FA
            Top             =   420
            Width           =   16875
         End
         Begin VB.TextBox txtQeryAppend 
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
            Left            =   -74640
            TabIndex        =   69
            Text            =   "query.append("""
            Top             =   720
            Width           =   1995
         End
         Begin VB.CommandButton cmdDel1 
            Caption         =   "제거"
            Height          =   375
            Left            =   -72600
            TabIndex        =   68
            Top             =   660
            Width           =   975
         End
         Begin VB.TextBox txtConvSQL 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   12
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   10875
            IMEMode         =   8  '영문
            Left            =   -74760
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   67
            Top             =   2220
            Width           =   18615
         End
         Begin VB.TextBox txtDel2 
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
            Left            =   -74640
            TabIndex        =   66
            Text            =   "\n"");"
            Top             =   1140
            Width           =   1995
         End
         Begin VB.CommandButton cmdDel2 
            Caption         =   "제거"
            Height          =   375
            Left            =   -72600
            TabIndex        =   65
            Top             =   1140
            Width           =   975
         End
         Begin VB.TextBox Text3 
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
            Left            =   -74640
            TabIndex        =   64
            Text            =   "//   -->   --"
            Top             =   1620
            Width           =   1995
         End
         Begin VB.CommandButton cmdDel3 
            Caption         =   "변환"
            Height          =   375
            Left            =   -72540
            TabIndex        =   63
            Top             =   1560
            Width           =   975
         End
         Begin VB.CommandButton cmdclear2 
            Caption         =   "CLEAR"
            Height          =   1095
            Left            =   -67560
            Style           =   1  '그래픽
            TabIndex        =   62
            Top             =   840
            Width           =   2595
         End
         Begin VB.CommandButton cmdCopy 
            Caption         =   "copy"
            Height          =   1155
            Left            =   -70920
            Style           =   1  '그래픽
            TabIndex        =   61
            Top             =   780
            Width           =   3015
         End
         Begin FPSpreadADO.fpSpread fpSpread2 
            Height          =   2355
            Left            =   -74880
            TabIndex        =   78
            Top             =   405
            Width           =   17175
            _Version        =   393216
            _ExtentX        =   30295
            _ExtentY        =   4154
            _StockProps     =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   200
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            SelectBlockOptions=   0
            SpreadDesigner  =   "frmOraMain.frx":2B437
            ScrollBarTrack  =   1
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "테이블명"
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
         Left            =   -73740
         TabIndex        =   143
         Top             =   5580
         Visible         =   0   'False
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmOraMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_NOTIN = "('SYS', 'SYSTEM')"

'최근 검색 내용을 파일에 저장
Private Sub saveNearFile()

On Error GoTo err_Handler
    Dim ffile As Integer
    Dim sTemp  As String
    Dim sTemp2() As String
    
    ffile = FreeFile

    Open App.Path & "\" & cboSystemList.Text & "_near.txt" For Output As #ffile   ' 파일을 엽니다.
    
    
    'Print #ffile, edtTable.Text&; ";" & edtTablenm.Text & txtOwner.Text
    
    
    
    
    Dim i As Integer
    For i = 1 To fpNear.MaxRows
        
        sTemp = ""
        fpNear.Row = i
        fpNear.Col = 1
        sTemp = sTemp & fpNear.Value
        
        fpNear.Col = 2
        sTemp = sTemp & ";" & fpNear.Value
        
        fpNear.Col = 5
        sTemp = sTemp & ";" & fpNear.Value
        
        Print #ffile, sTemp
    
    Next i
    
    Close #ffile
    Exit Sub
err_Handler:

    Close #ffile
    gfMsgWrn (Err.Description)
End Sub






Private Sub cboShema_Change()
    
    Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", "DB_OWNER", "OWNER_" & cboSystemList.Text, cboShema.Text)

End Sub

Private Sub cboShema_Click()
    Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", "DB_OWNER", "OWNER_" & cboSystemList.Text, cboShema.Text)
End Sub

Private Sub cboSystemList_Click()
    Dim strRegValue As String
    
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
    
    
     Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", "SYSTEM_LIST", "LAST_CONNECT_SYSTEM", cboSystemList.Text)
     
     
     
     
    If strReal = "T" Then
        'cboShema.Text = "LN"
    ElseIf strReal = "I" Then
        'cboShema.Text = "LN"
    Else
        'cboShema.Text = "LN"
    
    
    End If
    
    'owner list 설정
    Call setOwnerList

    
    '최근 검색 불러오기
    Call ReadFile
    
    '내 ower 지정
    
    cboShema.Text = GetSetting("ORAIB_TAB_MAN_BY_HAK", "DB_OWNER", "OWNER_" & cboSystemList.Text, "")
    
End Sub


Private Sub setOwnerList()
    
    
    cboShema.Clear
    If strReal = "A" Then
        cboShema.AddItem ("")
        cboShema.AddItem ("HG")
        cboShema.AddItem ("RG")
        cboShema.AddItem ("DW")
        cboShema.AddItem ("ER")
        cboShema.AddItem ("HSP")
        cboShema.AddItem ("SG")
        cboShema.AddItem ("BM")
        cboShema.AddItem ("KP")
        cboShema.AddItem ("EP")
        cboShema.AddItem ("AC")
        cboShema.AddItem ("HGMG")
        cboShema.AddItem ("IP")
        cboShema.AddItem ("MSS")
        cboShema.AddItem ("SE")
        cboShema.AddItem ("EW")
        cboShema.AddItem ("RU")
        cboShema.AddItem ("SC")
        cboShema.AddItem ("XDB")
        cboShema.AddItem ("CSS")
        cboShema.AddItem ("CU")
        cboShema.AddItem ("EI")
        cboShema.AddItem ("MS")
        cboShema.AddItem ("SL")
        cboShema.AddItem ("SS")
        cboShema.AddItem ("ZZ")
    'TOBE이면
    Else
        cboShema.AddItem ("")
        cboShema.AddItem ("USHG")
        cboShema.AddItem ("USRG")
        cboShema.AddItem ("USCO")
        cboShema.AddItem ("USCU")
        cboShema.AddItem ("USPD")
        cboShema.AddItem ("USBD")
        cboShema.AddItem ("USBM")
        cboShema.AddItem ("USSE")
        cboShema.AddItem ("USFB")
        cboShema.AddItem ("USEX")
    End If

End Sub



'----------------------------------------------------------------
'최금 검색 파일을
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
    
    fpNear.MaxRows = 0

    ffile = FreeFile

    Open App.Path & "\" & cboSystemList.Text & "_near.txt" For Input As #ffile   ' 파일을 엽니다.
    
    
    Do
    
        Input #ffile, sTemp
        
        sTemp2 = Split(sTemp, ";")
        
        fpNear.MaxRows = fpNear.MaxRows + 1

            
        fpNear.Row = fpNear.MaxRows
        
        fpNear.Col = 1
        fpNear.Value = sTemp2(0)

        fpNear.Col = 2
        fpNear.Value = sTemp2(1)

        fpNear.Col = 5
        fpNear.Value = sTemp2(2)
        
        
    Loop Until EOF(1)
    
    Close #ffile
    Exit Sub
ErrHandler:
    If Err.Number <> 53 And Err.Number <> 62 Then    '파일이 없는 경우
        gfMsgInf (Err.Description & Err.Number)
    End If
    
    Close #ffile
End Sub



'---------------------------------------------
'코드 보기
'---------------------------------------------
Private Sub Check1_Click()

End Sub

Private Sub chkASIS_Click()
    If chkASIS.Value = 1 Then
        Frame1.BackColor = &H80FF&
    Else
        Frame1.BackColor = &HC0C0C0
        
    End If
End Sub

Private Sub TabMaxMin(sTy As String)

    



End Sub
'--------------------------------
'코드보기
'--------------------------------
Private Sub chkCode_Click()
    Dim i, j As Integer
    Dim sWhere As String
    
    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    
    Dim sCodeName As String
    Dim sCodeValue As String
    
    Dim sSplit() As String
    
On Error GoTo Error_Handler



    Screen.MousePointer = vbHourglass
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
    Set conn = New ADODB.Connection
    conn.ConnectionString = GetConnString

    conn.Open


    
    If chkCode.Value = "1" Then
        cmdInsertUpdate.Enabled = False
    Else
        cmdInsertUpdate.Enabled = True
    End If


        
    For i = 1 To fpSample.MaxCols
        fpSample.Row = 0
        fpSample.Col = i
        
        If Right(fpSample.Value, 2) = "코드" Then
            If Trim(sWhere) = "" Then
                sWhere = sWhere & "'" & fpSample.Value & "'"
            Else
                sWhere = sWhere & ", '" & fpSample.Value & "'"
            End If
                    
        
        End If
    Next i
    
    If Trim(sWhere) = "" Then Exit Sub
    
    
    sSQL = ""

    sSQL = sSQL & vbCrLf & "SELECT A.CD_GRP_ID"
    sSQL = sSQL & vbCrLf & "      ,A.CD_GRP_NM"
    sSQL = sSQL & vbCrLf & "      ,''   as eng_nm"
    sSQL = sSQL & vbCrLf & "      ,B.CD_ID"
    sSQL = sSQL & vbCrLf & "      ,B.CD_NM"
    sSQL = sSQL & vbCrLf & "      ,''       AS USE_YN"
    sSQL = sSQL & vbCrLf & "      ,''       AS REMARKS"
    sSQL = sSQL & vbCrLf & "      ,B.CD_ID || '-' || B.CD_NM   AS ADDCODE "
    sSQL = sSQL & vbCrLf & "  FROM USCO.TB_COA310M_CDGRPBSC A"
    sSQL = sSQL & vbCrLf & "      ,USCO.TB_COA311M_CDBSC B"
    sSQL = sSQL & vbCrLf & " WHERE 1=1"
    sSQL = sSQL & vbCrLf & "   AND A.CD_GRP_ID = B.CD_GRP_ID"
    sSQL = sSQL & vbCrLf & "   AND A.CD_GRP_NM IN (" & sWhere & ")"

    rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

    
    
    
    For i = 1 To fpSample.MaxRows
        For j = 1 To fpSample.MaxCols
                
                
            fpSample.Row = 0
            fpSample.Col = j
            
            sCodeName = fpSample.Value
            
            fpSample.Row = i
            sCodeValue = fpSample.Value
                            
            If InStr(1, sCodeValue, "-") > 0 Then
            
                sSplit = Split(sCodeValue, "-")
            
            End If
                            
            If Right(sCodeName, 2) = "코드" Then
                
                If chkCode.Value = "1" Then
                    rsTemp.Filter = adFilterNone
                    rsTemp.Filter = "CD_GRP_NM='" & sCodeName & "' AND CD_ID ='" & sCodeValue & "'"
                Else
                    If InStr(1, sCodeValue, "-") > 0 Then
                        rsTemp.Filter = adFilterNone
                        rsTemp.Filter = "CD_GRP_NM='" & sCodeName & "' AND CD_ID ='" & sSplit(0) & "'"
                
                    Else
                        rsTemp.Filter = adFilterNone
                        rsTemp.Filter = "CD_GRP_NM='" & sCodeName & "' AND CD_ID ='" & sCodeValue & "'"
                    End If
                End If
                            
                
                If rsTemp.RecordCount > 0 Then
                    If chkCode.Value = "1" Then
                        rsTemp.MoveFirst
                        fpSample.Value = rsTemp.Fields("ADDCODE").Value
                    Else
                        rsTemp.MoveFirst
                        fpSample.Value = rsTemp.Fields("CD_ID").Value
                    End If
                End If
            End If
                    
        
        Next j
    Next i
    
    
    
    Screen.MousePointer = vbDefault

    Set conn = Nothing
    Set rsTemp = Nothing
        
    
    Exit Sub
Error_Handler:
    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description

End Sub

Private Sub chkComment_Click()
    subMakeSQL
End Sub

Private Sub cmdAdd_Click()
    fpcode.MaxRows = fpcode.MaxRows + 1
End Sub


'SQL AS 구문 정렬
Private Sub cmdAlign_Click()
    On Error GoTo err_Handler
    
    Dim sOrgArr() As String
    Dim sSplit() As String
    Dim i As Integer
    Dim iMaxSize As Integer
    
    txtOrgSQL.Text = UCase(txtOrgSQL.Text)
    
    sOrgArr = Split(txtOrgSQL.Text, vbCrLf)
    
    txtTgtSQL.Text = ""
    iMaxSize = 0
    For i = LBound(sOrgArr) To UBound(sOrgArr)
                    
                    
        If InStr(1, UCase(sOrgArr(i)), cboSplliter.Text) Then
            sSplit = Split(sOrgArr(i), cboSplliter.Text)
            If iMaxSize < Len(sSplit(0)) Then
                iMaxSize = Len(sSplit(0))
            End If
        
        Else
            
        End If
    Next i
    
    For i = LBound(sOrgArr) To UBound(sOrgArr)
                    
                    
        If InStr(1, UCase(sOrgArr(i)), cboSplliter.Text) Then
            sSplit = Split(sOrgArr(i), cboSplliter.Text)
            
            txtTgtSQL.Text = txtTgtSQL.Text & sSplit(0) & gfMakeString(" ", iMaxSize - Len(sSplit(0)) + 5) & cboSplliter.Text
            txtTgtSQL.Text = txtTgtSQL.Text & sSplit(1) & vbCrLf
        Else
            txtTgtSQL.Text = txtTgtSQL.Text & sOrgArr(i) & vbCrLf
        End If
    Next i
    
    '주석 정렬
    sOrgArr = Split(txtTgtSQL.Text, vbCrLf)
    
    txtTgtSQL.Text = ""
    iMaxSize = 0
    For i = LBound(sOrgArr) To UBound(sOrgArr)
                    
                    
        If InStr(1, UCase(sOrgArr(i)), "--") Then
            sSplit = Split(sOrgArr(i), "--")
            If iMaxSize < Len(sSplit(0)) Then
                iMaxSize = Len(sSplit(0))
            End If
        ElseIf InStr(1, UCase(sOrgArr(i)), "/*") Then
            sSplit = Split(sOrgArr(i), "/*")
            If iMaxSize < Len(sSplit(0)) Then
                iMaxSize = Len(sSplit(0))
            End If
        
        ElseIf InStr(1, UCase(sOrgArr(i)), "//") Then
            sSplit = Split(sOrgArr(i), "//")
            If iMaxSize < Len(sSplit(0)) Then
                iMaxSize = Len(sSplit(0))
            End If

            
        End If
    Next i
    
    For i = LBound(sOrgArr) To UBound(sOrgArr)
                    
                    
        If InStr(1, UCase(sOrgArr(i)), "--") Then
            sSplit = Split(sOrgArr(i), "--")
            
            txtTgtSQL.Text = txtTgtSQL.Text & sSplit(0) & gfMakeString(" ", iMaxSize - Len(sSplit(0)) + 5) & "--"
            txtTgtSQL.Text = txtTgtSQL.Text & sSplit(1) & vbCrLf
        ElseIf InStr(1, UCase(sOrgArr(i)), "/*") Then
            sSplit = Split(sOrgArr(i), "/*")
            
            txtTgtSQL.Text = txtTgtSQL.Text & sSplit(0) & gfMakeString(" ", iMaxSize - Len(sSplit(0)) + 5) & "/*"
            txtTgtSQL.Text = txtTgtSQL.Text & sSplit(1) & vbCrLf
        Else
            txtTgtSQL.Text = txtTgtSQL.Text & sOrgArr(i) & vbCrLf
        End If
    Next i
    
    
    
    Exit Sub
err_Handler:
    gfMsgWrn Err.Description
End Sub

Private Sub cmdBookMakr_Click()
    
    
    Dim i As Integer
    For i = 1 To fpCol.MaxRows
        fpCol.Row = i
        fpCol.Col = 5
        
        If fpCol.Value <> "" Then
            fpCol.Col = -1
            
            fpCol.BackColor = &HFFC0FF
        Else
            fpCol.Col = -1
            fpCol.BackColor = S_WHITE
        End If
        
        
    
    Next i
End Sub

Private Sub cmdBookReset_Click()
    
    
    Dim i As Integer
    For i = 1 To fpcode.MaxRows
        fpcode.Row = i
        fpcode.Col = -1
        
      fpcode.BackColor = S_WHITE
        
    
    Next i
End Sub

Private Sub cmdClear2_Click()
    txtConvSQL.Text = ""
End Sub


Private Sub cmdCntSerarch_Click()
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


    sSQL = ""
    
    
    sSQL = sSQL & "SELECT"
    sSQL = sSQL & "  (SELECT COUNT(*) " & vbCrLf
    sSQL = sSQL & "    FROM " & Trim(txtMigOwner.Text) & "." & Trim(txtMigTableID.Text) & "@LK_STG3) AS MIG_CNT " & vbCrLf
    sSQL = sSQL & " ,(SELECT COUNT(*) " & vbCrLf
    sSQL = sSQL & "    FROM " & Trim(txtTgtOwner.Text) & "." & Trim(txtTgtTableId.Text) & ") AS DEV_CNT " & vbCrLf
    sSQL = sSQL & "  FROM DUAL "
    
   
    txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If

    rsTemp.MoveFirst
    
    txtMigCnt.Text = rsTemp.Fields("MIG_CNT").Value
    txtTgtCnt.Text = rsTemp.Fields("DEV_CNT").Value
    
        
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

'컬럼정보 일괄 등록
Private Sub cmdColMultiple_Click()
    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    
    If GCONN_USE = False Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = GetConnString
    
        conn.Open
        conn.BeginTrans
    Else
        gConn.BeginTrans
        
    End If
    
    Dim i As Integer
    
    
    Dim sColid As String
    Dim sColnm As String
    Dim sCodeDomain As String
    Dim sColBigo As String
    Dim sColBigo2 As String
    
    If strReal = "R" Then
        MsgBox "운영계 처리불가"
        
    End If
    
    For i = 1 To fpCol.MaxRows
        
        fpCol.Row = i
        fpCol.Col = 1: sColid = fpCol.Value
        fpCol.Col = 2: sColnm = fpCol.Value
        fpCol.Col = 6: sCodeDomain = fpCol.Value
        fpCol.Col = 7: sColBigo = fpCol.Value
        fpCol.Col = 10: sColBigo2 = fpCol.Value
    
    
    
        sSQL = ""
        sSQL = sSQL & vbCrLf & " DELETE FROM ln.TB_DIC_COL" & IIf(chkASIS.Value = 1, "@DL_IBKCDEV", IIf(strReal = "R", "@DL_DEOSDB", "")) & " "
        If Trim(cboShema.Text) <> "" Then
            sSQL = sSQL & vbCrLf & "  WHERE OWNER  = '" & IIf(chkASIS.Value = 1, "IBKCDEV", UCase(cboShema.Text)) & "'"
        Else
            sSQL = sSQL & vbCrLf & "  WHERE OWNER      NOT IN  " & C_NOTIN
        End If
        sSQL = sSQL & vbCrLf & "    AND TABLE_NAME  = '" & edtTable.Text & "' "
        sSQL = sSQL & vbCrLf & "    AND COLUMN_NAME  = '" & sColid & "'"
        
        
        txtSQL.Text = sSQL
        If GCONN_USE = True Then
            gConn.Execute sSQL
        Else
            conn.Execute sSQL
        End If
        
    
        
        sSQL = ""
        sSQL = sSQL & vbCrLf & " INSERT INTO ln.TB_DIC_COL (OWNER, TABLE_NAME, COLUMN_NAME, COLUMN_COMM, CODE_DOM, BIGO1, BIGO2) "
        sSQL = sSQL & vbCrLf & "  VALUES ('" & cboShema.Text & "' ,"
        sSQL = sSQL & vbCrLf & "          '" & edtTable.Text & "' ,"
        sSQL = sSQL & vbCrLf & "          '" & Trim(sColid) & "' ,"
        sSQL = sSQL & vbCrLf & "          '" & Trim(sColnm) & "' ,"
        sSQL = sSQL & vbCrLf & "          '" & Trim(sCodeDomain) & "' ,"
        sSQL = sSQL & vbCrLf & "          '" & Trim(sColBigo) & "' ,"
        sSQL = sSQL & vbCrLf & "          '" & Trim(sColBigo2) & "' )"
        
        
        txtSQL.Text = sSQL
        If GCONN_USE = True Then
            gConn.Execute sSQL
        Else
            conn.Execute sSQL
        End If
    
    Next i
    
    
    gfMsg ("등록되었습니다.")

    'commit
    If GCONN_USE = True Then
        gConn.CommitTrans
    Else
        conn.CommitTrans
    End If
    
    '테이블 레이아웃 재조회
    fpTable.MaxRows = fpTable.MaxRows + 1
    fpTable.Row = fpTable.MaxRows
    
    
    fpTable.Col = 1
    fpTable.Value = edtTable.Text
    
    fpTable.Col = 2
    fpTable.Value = edtTablenm
    
    
    Call fpTable_DblClick(1, fpTable.MaxRows)
    
    
    
    
    
    Screen.MousePointer = vbDefault

    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    Exit Sub
ErrorHandler:
    
    If GCONN_USE = True Then
        gConn.RollbackTrans
        
    Else
        conn.RollbackTrans
        
    End If
    
    
    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description

End Sub

Private Sub cmdColRow_Click()
    Dim i As Integer
    Dim j As Integer
On Error Resume Next
    With frmDataView.fpDataView
        .MaxCols = C_DADA_VIEW_BASIC_COL_SIZE
        .MaxCols = C_DADA_VIEW_BASIC_COL_SIZE + fpSample.MaxRows
        .MaxRows = 0
        .MaxRows = UBound(gColId)
        
        For i = 1 To UBound(gColId)
            .Row = i
            .Col = 2: .Value = gColId(i)
            .Col = 3: .Value = gColName(i)
            .Col = 4: .Value = gType(i)
            .Col = 5: .Value = gPK(i)
        Next i
        
        For i = 1 To fpSample.MaxRows
            .Col = C_DADA_VIEW_BASIC_COL_SIZE + i
            .Row = 0
            .Value = i
                
            .ColWidth(C_DADA_VIEW_BASIC_COL_SIZE + i) = 20
         
            fpSample.Row = i
            For j = 1 To UBound(gColId)
                .Row = j
                fpSample.Col = j
                
                    .Value = fpSample.Value
                
            Next j
        Next i
                
        'PK색칠하기
        For i = 1 To .MaxRows
            .Row = i
            .Col = 5
            
            If .Value <> "" Then
                .Col = -1
                
                .BackColor = &HFFC0FF
            Else
                .Col = -1
                .BackColor = S_WHITE
            End If
        
        Next i
        
        
        frmDataView.edtTable.Text = edtTable.Text
        frmDataView.edtTablenm.Text = edtTablenm.Text
        
        frmDataView.Left = Me.Left
        frmDataView.Top = Me.Top
        frmDataView.Show
    End With
End Sub

Private Sub cmdColumn_Click()
    frmColumnSrch.Show
End Sub

'******************************************
'컬럼정보 update
'******************************************
Private Sub cmdColUpdate_Click()
Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    
    If GCONN_USE = False Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = GetConnString
    
        conn.Open
        conn.BeginTrans
    Else
        gConn.BeginTrans
        
    End If
    


    sSQL = ""
    sSQL = sSQL & vbCrLf & " DELETE FROM ln.TB_DIC_COL" & IIf(chkASIS.Value = 1, "@DL_IBKCDEV", "") & " "
    If Trim(cboShema.Text) <> "" Then
        sSQL = sSQL & vbCrLf & "  WHERE OWNER  = '" & IIf(chkASIS.Value = 1, "IBKCDEV", UCase(cboShema.Text)) & "'"
    Else
        sSQL = sSQL & vbCrLf & "  WHERE OWNER      NOT IN  " & C_NOTIN
    End If
    sSQL = sSQL & vbCrLf & "    AND TABLE_NAME  = '" & edtTable.Text & "' "
    sSQL = sSQL & vbCrLf & "    AND COLUMN_NAME  = '" & txtColId.Text & "' "
    
    If GCONN_USE = True Then
        gConn.Execute sSQL
    Else
        conn.Execute sSQL
    End If
    

    
    sSQL = ""
    sSQL = sSQL & vbCrLf & " INSERT INTO ln.TB_DIC_COL (OWNER, TABLE_NAME, COLUMN_NAME, COLUMN_COMM, CODE_DOM, BIGO1, BIGO2) "
    sSQL = sSQL & vbCrLf & "  VALUES ('" & cboShema.Text & "' ,"
    sSQL = sSQL & vbCrLf & "          '" & edtTable.Text & "' ,"
    sSQL = sSQL & vbCrLf & "          '" & txtColId.Text & "' ,"
    sSQL = sSQL & vbCrLf & "          '" & Trim(txtColNm.Text) & "' ,"
    sSQL = sSQL & vbCrLf & "          '" & Trim(txtCodeDomain.Text) & "' ,"
    sSQL = sSQL & vbCrLf & "          '" & Trim(txtColBigo.Text) & "' ,"
    sSQL = sSQL & vbCrLf & "          '" & " " & "' )"
    
    

    If GCONN_USE = True Then
        gConn.Execute sSQL
    Else
        conn.Execute sSQL
    End If
    
    
    gfMsg ("등록되었습니다.")

    'commit
    If GCONN_USE = True Then
        gConn.CommitTrans
    Else
        conn.CommitTrans
    End If
    
    Dim iRow As Integer
    iRow = fpCol.TopRow
    
    
    '테이블 레이아웃 재조회
    fpTable.MaxRows = fpTable.MaxRows + 1
    fpTable.Row = fpTable.MaxRows
    
    
    fpTable.Col = 1
    fpTable.Value = edtTable.Text
    
    fpTable.Col = 2
    fpTable.Value = edtTablenm
    
    
    Call fpTable_DblClick(1, fpTable.MaxRows)
    
    
    
    
    
    
    Call fpCol.ShowCell(1, iRow, PositionUpperCenter)
    
    
    Screen.MousePointer = vbDefault

    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    Exit Sub
ErrorHandler:
    
    If GCONN_USE = True Then
        gConn.RollbackTrans
        
    Else
        conn.RollbackTrans
        
    End If
    
    
    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description
'
End Sub

'부서정보 검색
Private Sub cmdFindDept_Click()
    'http://134.100.1.60/com.ibk.finehr.hrm.comm.EmpnoSearchBranchInfoSelAction.action?brcd=001&sejno=A97011
    'http://134.100.1.60/com.ibk.finehr.hrm.comm.EmpnoSearchBranchListSelAction.action?brnm=영업부&sejno=A97011
    
    Dim sUrl As String
    
    'If IsNumeric(Trim(txtIBKBankEmp.Text)) Then
    '    sUrl = "http://134.100.1.60/com.ibk.finehr.hrm.comm.EmpnoSearchBranchInfoSelAction.action?brcd=" & Trim(txtIBKBankEmp.Text)
    'Else
    '    sUrl = "http://134.100.1.60/com.ibk.finehr.hrm.comm.EmpnoSearchBranchListSelAction.action?brnm=" & Trim(txtIBKBankEmp.Text)
    'End If
    sUrl = "http://150.100.21.183:7005/logon_emer.html" 'IBK캐피탈 메타정보
    WebBrowser1.Navigate (sUrl)
End Sub




' 쿼리변환
Private Sub cmdConvert_Click()
    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

'On Error GoTo ErrorHandler
    
    If Trim(txtOldTabName.Text) = "" Then
        gfMsg ("테이블명을 입력하세요")
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    txtNewSQL.Text = ""
    If GCONN_USE = False Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = GetConnString
    
        conn.Open
    End If
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient

    

    sSQL = sSQL & vbCrLf & "SELECT * FROM META_TABLE_MAP@dl_meosdb  "
    sSQL = sSQL & vbCrLf & "  WHERE 1=1"
    sSQL = sSQL & vbCrLf & " AND TRIM(UPPER(A_ENG_TABLE_NAME)) = '" & UCase(Trim(txtOldTabName.Text)) & "'"
    If Trim(txtNewTabName.Text) <> "" Then
        sSQL = sSQL & vbCrLf & " AND UPPER(T_ENG_TABLE_NAME) = '" & UCase(Trim(txtNewTabName.Text)) & "'"
    End If

    sSQL = sSQL & vbCrLf & " ORDER BY   to_number(map_sort)"
    
   
    txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If


    Dim i As Integer
    Dim sAlias As String
    
    If Trim(txtAsisAlias.Text) <> "" Then
        sAlias = Trim(txtAsisAlias.Text) & "."
    Else
        sAlias = ""
    End If
    
    txtNewSQL.Text = UCase(txtOldSQL.Text)
    
    If rsTemp.EOF Then Exit Sub
    rsTemp.MoveFirst
    
    For i = 1 To rsTemp.RecordCount
        If rsTemp.Fields("A_ENG_COLUMN_NAME") <> "" Then
            If rsTemp.Fields("T_ENG_COLUMN_NAME") <> "" Then
                If rsTemp.Fields("A_ENG_TABLE_NAME") <> "" Then
                    If rsTemp.Fields("T_ENG_TABLE_NAME") <> "" Then
                        txtNewSQL.Text = Replace(txtNewSQL.Text, Trim(rsTemp.Fields("A_ENG_COLUMN_NAME")), Trim(rsTemp.Fields("T_ENG_COLUMN_NAME")))
                        txtNewSQL.Text = Replace(txtNewSQL.Text, Trim(rsTemp.Fields("A_ENG_TABLE_NAME")), Trim(rsTemp.Fields("T_ENG_TABLE_NAME")))
                    End If
                End If
            End If
        End If
        rsTemp.MoveNext
    Next i
    
    
    
   
    
    
        
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

Private Sub cmdCopy_Click()

Clipboard.Clear

    Clipboard.SetText (txtConvSQL.Text)
End Sub

Private Sub cmdCopyRow_Click()
    
    With fpSample
        
        .MaxRows = .MaxRows + 1
        .InsertRows .ActiveRow + 1, 1
        
        .Col = -1
        .Row = .ActiveRow + 1
        .BackColor = C_ADD_COLOR
    
    End With
    
    Dim i As Integer
    Dim sValue As String
    
    
    For i = 1 To fpSample.MaxCols
        fpSample.Col = i
        fpSample.Row = fpSample.ActiveRow
        sValue = fpSample.Value
        
        fpSample.Row = fpSample.ActiveRow + 1
        fpSample.Value = sValue
    
        
        
    Next i
    
    makeCellBorderline
    
End Sub

Private Sub cmdDebugLog_Click()
    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
   
    
    If GCONN_USE = False Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = GetConnString
    
        conn.Open
    End If
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
    If Trim(cmbPgmIDLog.Text) = "" Then
        MsgBox " Program Id를 입력하랑께롱 ", vbCritical
        cmbPgmIDLog.SetFocus
        
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass

    sSQL = ""
    sSQL = sSQL & vbCrLf & "SELECT                                                                                                     "
    sSQL = sSQL & vbCrLf & "       LAG(PRRM_ID,1,PRRM_ID) OVER (ORDER BY LOG_SRNO)  AS PRRM_ID                          "
    sSQL = sSQL & vbCrLf & "      ,MNNO                                                                                                "
    sSQL = sSQL & vbCrLf & "      ,CASE WHEN LOG_SRNO = 99999 THEN LAG(LOG_SRNO,1,LOG_SRNO) OVER (ORDER BY LOG_SRNO)  + 1             "
    sSQL = sSQL & vbCrLf & "             ELSE LOG_SRNO                                                                                 "
    sSQL = sSQL & vbCrLf & "       END  AS  LOG_SRNO                                                                                  "
    sSQL = sSQL & vbCrLf & "      ,BULT_CNTN                                                                                          "
    sSQL = sSQL & vbCrLf & "      ,AF_TIME                                                                                            "
    sSQL = sSQL & vbCrLf & "      ,TO_CHAR(FLOOR(AF_TIME - BF_TIME))  ||  '일 ' ||                                                     "
    sSQL = sSQL & vbCrLf & "       TO_CHAR(                                                                                            "
    sSQL = sSQL & vbCrLf & "                FLOOR( (MOD(AF_TIME - BF_TIME, 1) * 24) )                                                  "
    sSQL = sSQL & vbCrLf & "              , '00') || ':' ||                                                                            "
    sSQL = sSQL & vbCrLf & "       TO_CHAR(                                                                                            "
    sSQL = sSQL & vbCrLf & "                   FLOOR( (MOD( (MOD(AF_TIME - BF_TIME, 1                                                  "
    sSQL = sSQL & vbCrLf & "                                    ) * 24                                                                 "
    sSQL = sSQL & vbCrLf & "                                 ), 1                                                                      "
    sSQL = sSQL & vbCrLf & "                              ) * 60                                                                       "
    sSQL = sSQL & vbCrLf & "                          )                                                                                "
    sSQL = sSQL & vbCrLf & "                        )                                                                                  "
    sSQL = sSQL & vbCrLf & "             , '00' ) || ':' ||                                                                            "
    sSQL = sSQL & vbCrLf & "       TO_CHAR(                                                                                            "
    sSQL = sSQL & vbCrLf & "                   FLOOR( MOD (                                                                            "
    sSQL = sSQL & vbCrLf & "                                   (MOD( (MOD(AF_TIME - BF_TIME, 1                                         "
    sSQL = sSQL & vbCrLf & "                                            ) * 24                                                         "
    sSQL = sSQL & vbCrLf & "                                         ), 1                                                              "
    sSQL = sSQL & vbCrLf & "                                      ) * 60                                                               "
    sSQL = sSQL & vbCrLf & "                                  )                                                                        "
    sSQL = sSQL & vbCrLf & "                                , 1                                                                        "
    sSQL = sSQL & vbCrLf & "                              ) * 60                                                                       "
    sSQL = sSQL & vbCrLf & "                        )                                                                                  "
    sSQL = sSQL & vbCrLf & "              , '00' )                                                                                     "
    sSQL = sSQL & vbCrLf & "  FROM                                                                                                     "
    sSQL = sSQL & vbCrLf & "        (                                                                                                  "
    sSQL = sSQL & vbCrLf & "            SELECT  PRRM_ID                                                                           "
    sSQL = sSQL & vbCrLf & "                   ,LOG_SRNO                                                                               "
    sSQL = sSQL & vbCrLf & "                   ,LAG(LOG_TRYMD,1, LOG_TRYMD) OVER (ORDER BY LOG_SRNO)  AS BF_TIME                       "
    sSQL = sSQL & vbCrLf & "                   ,LOG_TRYMD                                  AS AF_TIME                                  "
    sSQL = sSQL & vbCrLf & "                   ,BULT_CNTN   ,MNNO                                                                           "
    sSQL = sSQL & vbCrLf & "              FROM (                                                                                       "
    sSQL = sSQL & vbCrLf & "                    SELECT  PRRM_ID                                                                   "
    sSQL = sSQL & vbCrLf & "                           ,LOG_SRNO                                                                       "
    sSQL = sSQL & vbCrLf & "                           ,LOG_TRYMD                                                                      "
    sSQL = sSQL & vbCrLf & "                           ,BULT_CNTN   ,MNNO                                                                   "
    sSQL = sSQL & vbCrLf & "                      FROM TB_CRS_CM013_L_A                                                                "
    sSQL = sSQL & vbCrLf & "                     WHERE PRRM_ID = UPPER('" & cmbPgmIDLog.Text & "')                                "
    sSQL = sSQL & vbCrLf & "                       AND TO_NUMBER(MNNO) = (SELECT MAX(to_number(MNNO))          "
    sSQL = sSQL & vbCrLf & "                                     FROM TB_CRS_CM013_L_A                                                 "
    sSQL = sSQL & vbCrLf & "                                    WHERE PRRM_ID = UPPER('" & cmbPgmIDLog.Text & "')                                              "
    sSQL = sSQL & vbCrLf & "                                  )                                                                        "
    sSQL = sSQL & vbCrLf & "                    UNION ALL                                                                              "
    sSQL = sSQL & vbCrLf & "                    SELECT  'ZZZZZZ'                                                                       "
    sSQL = sSQL & vbCrLf & "                           ,99999                                                                          "
    sSQL = sSQL & vbCrLf & "                           ,SYSDATE                                                                        "
    sSQL = sSQL & vbCrLf & "                           ,'현재 시각'     ,''                                                               "
    sSQL = sSQL & vbCrLf & "                      FROM DUAL                                                                            "
    sSQL = sSQL & vbCrLf & "                    ) X                                                                                    "
    sSQL = sSQL & vbCrLf & "         )                                                                                                 "
    sSQL = sSQL & vbCrLf & "   ORDER BY   LOG_SRNO                                                                                     "
    
    

    txtSQL.Text = sSQL

    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If

    fpDebugLog.MaxRows = 0
    If rsTemp.RecordCount < 1 Then
        fpDebugLog.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpDebugLog, rsTemp)
    End If
    
    
    
    
    
    
    
    
    'log hist 조회
 Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
    If Trim(cmbPgmIDLog.Text) = "" Then
        MsgBox "Input Program Id "
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    
    
    Dim sStDay As String
    Dim sEdDay As String
    sSQL = ""
    sSQL = sSQL & vbCrLf & "SELECT  PRRM_ID                                                                                        "
    sSQL = sSQL & vbCrLf & "       ,MNNO                                                                                                "
    sSQL = sSQL & vbCrLf & "       ,MIN(LOG_TRYMD)                                                                                      "
    sSQL = sSQL & vbCrLf & "       ,MAX(LOG_TRYMD)                                                                                      "
    sSQL = sSQL & vbCrLf & "  FROM TB_CRS_CM013_L_A                                                                                     "
    sSQL = sSQL & vbCrLf & " WHERE PRRM_ID = UPPER('" & Trim(cmbPgmIDLog.Text) & "')                                                                         "
    sSQL = sSQL & vbCrLf & "   AND LOG_TRYMD BETWEEN TO_DATE('" & Trim(txtStDay.Text) & "', 'YYYY-MM-DD') AND TO_DATE('" & Trim(txtEdDay.Text) & "', 'YYYY-MM-DD') + 1        "
    sSQL = sSQL & vbCrLf & "GROUP BY PRRM_ID ,MNNO                                                                                 "
    sSQL = sSQL & vbCrLf & "ORDER BY  TO_NUMBER(MNNO)   DESC                                                                          "


    txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If

    fpDebugLogHist.MaxRows = 0
    If rsTemp.RecordCount < 1 Then
        fpDebugLogHist.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpDebugLogHist, rsTemp)
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

Private Sub cmdDel_Click()
    'fpcode.Row = fpcode.ActiveRow
    Call fpcode.DeleteRows(fpcode.ActiveRow, 1)
    fpcode.MaxRows = fpcode.MaxRows - 1
    
End Sub

Private Sub cmdDel1_Click()
    txtConvSQL.Text = Replace(txtConvSQL.Text, Trim(txtQeryAppend.Text), "")
End Sub

Private Sub cmdDel2_Click()
    txtConvSQL.Text = Replace(txtConvSQL.Text, Trim(txtDel2.Text), "")
End Sub

Private Sub cmdDel3_Click()
    txtConvSQL.Text = Replace(txtConvSQL.Text, "//", "--")
End Sub

Private Sub cmdDelLine_Click()
    With fpSample
        
        .Col = -1
        .Row = .ActiveRow
        .BackColor = S_DARK_RED
    
    End With
    
    
    makeCellBorderline
End Sub

Private Sub cmdEHR_Click()
    WebBrowser1.Navigate ("https://ehr.ibksystem.co.kr/JSP/LoginHrms.jsp")

End Sub

Private Sub cmdExcel_Click()
    If Trim(edtTable.Text) = "" Then
        gfMsg "테이블을 선택하십시오"
        Exit Sub
    End If
    
    Call gfFpToExcel(fpCol)
End Sub

Private Sub cmdFavMinus_Click()
    Dim i As Integer
    If Trim(edtTable.Text) = "" Then Exit Sub

    For i = 1 To fpFavorite.MaxRows
        fpFavorite.Row = i
        fpFavorite.Col = 1
        
        If Trim(fpFavorite.Value) = Trim(edtTable.Text) Then
             fpFavorite.DeleteRows i, 1
             fpFavorite.MaxRows = fpFavorite.MaxRows - 1
            Exit Sub
        End If
    Next i


    

    Call setFavorite
End Sub

Private Sub cmdFavPlus_Click()
                
    Dim i As Integer
 
    If Trim(edtTable.Text) = "" Then Exit Sub
        
    For i = 1 To fpFavorite.MaxRows
        fpFavorite.Row = i
        fpFavorite.Col = 1
        
        If Trim(fpFavorite.Value) = Trim(edtTable.Text) Then Exit Sub
    Next i
        
    fpFavorite.MaxRows = fpFavorite.MaxRows + 1
    fpFavorite.Row = fpFavorite.MaxRows
    fpFavorite.Col = 1
    fpFavorite.Value = edtTable.Text

    fpFavorite.Col = 2
    fpFavorite.Value = edtTablenm.Text
    
    fpFavorite.Col = 5
    fpFavorite.Value = txtOwner.Text
    
    
    SSTab8.Tab = 1
        
    Call setFavorite
End Sub

'지캠프 직원찾기
Private Sub cmdFineNetemp_Click()
    WebBrowser1.Navigate ("http://134.100.1.60/hrm/comm/empno_search_sel_frm.jsp")
    
End Sub

'
Private Sub cmdGetClipboard_Click()
    If SSTab5.Tab = 0 Then Clipboard.SetText (txtSelectSQL.Text)
    If SSTab5.Tab = 1 Then Clipboard.SetText (txtInsertSQL.Text)
    If SSTab5.Tab = 2 Then Clipboard.SetText (txUpdateSQL.Text)
    If SSTab5.Tab = 3 Then Clipboard.SetText (txtInsSelect.Text)
    If SSTab5.Tab = 5 Then Clipboard.SetText (txtIQuery.Text)
End Sub

'회사그룹웨어
Private Sub cmdIBKSGroup_Click()
    WebBrowser1.Navigate ("http://iworld.ibksystem.co.kr/iFrame/out/main.xml")
    
End Sub

'IBKS직원찾기
Private Sub cmdIBKsystem_Click()
    WebBrowser1.Navigate ("http://iworld.ibksystem.co.kr/iFrame?Class=IDSearch&직원조회=1")
    
End Sub

'---------------------------------------------------------
'데이터를 신규 등록 하거나 변경 등록 함.
'---------------------------------------------------------
Private Sub cmdInsertUpdate_Click()
    subExeSql True
End Sub

Private Sub Command7_Click()
    subExeSql False
End Sub


Private Sub subExeSql(bExeYn As Boolean)
    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    Dim i As Integer
    

On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    
    If GCONN_USE = False Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = GetConnString
    
        conn.Open
        conn.BeginTrans
    Else
        gConn.BeginTrans
        
    End If
    
    If bExeYn Then
        If gfMsgQst("저장 하시겠습니까?") = vbNo Then Exit Sub
    End If
    
Dim iSample As Integer

txtSQL.Text = ""

For iSample = 1 To fpSample.MaxRows
    
'00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
    
    Dim sCode As String
    Dim sName As String
    
    Dim iMaxLen As Integer
    iMaxLen = 0
    Dim iSpace As Integer
    iSpace = 10
    Dim sTy As String
    
    '최대 컬럼 사이즈 계산
    For i = 1 To fpCol.MaxRows
        fpCol.Row = i
        fpCol.Col = 1
        
        If Len(Trim(fpCol.Value)) > iMaxLen Then
            iMaxLen = Len(Trim(fpCol.Value))
        End If
        
    Next i
    
    
    Dim sISQL As String
    Dim sISQLValue As String
    sISQLValue = ""
    
    
    Dim sIQuerySQL As String
    Dim sIQuerySQLValue As String
    sIQuerySQLValue = ""
    
    Dim sUSQL As String
    Dim sUSQLWhere As String
    sUSQLWhere = ""
    
    Dim sSSQL As String
    Dim sSSQL2 As String
    
    
    Dim sGSQL As String
    
    sISQL = "INSERT INTO " & edtTable.Text & " ( " & "        /* " & edtTablenm.Text & "*/" & vbCrLf
    sISQLValue = "( "
    
    
    sIQuerySQL = "INSERT INTO " & edtTable.Text & " ( " & vbCrLf
    sIQuerySQLValue = "( "
    
    sUSQL = "UPDATE " & edtTable.Text & "        /* " & edtTablenm.Text & "*/" & vbCrLf & "   SET " & vbCrLf
    sUSQLWhere = " WHERE 1=1" & vbCrLf
    
    
    sSSQL = "SELECT "
    sSSQL2 = "SELECT "
    
    sGSQL = " "
    

    
    For i = 1 To fpCol.MaxRows
    
        fpCol.Row = i
        fpCol.Col = 1
        sCode = Trim(fpCol.Value)
        fpCol.Col = 2
        sName = Trim(fpCol.Value)
        fpCol.Col = 3
        sTy = Left(fpCol.Value, 1)
        

        
        'insert field
        If i <> 1 Then
            sISQL = sISQL & "         ,"
            sIQuerySQL = sIQuerySQL & "         ,"
        Else
            sISQL = sISQL & "          "
            sIQuerySQL = sIQuerySQL & "          "
        End If
        
        
        sISQL = sISQL & "" & sCode
        sISQL = sISQL & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
        
        
        sIQuerySQL = sIQuerySQL & "" & sCode
        sIQuerySQL = sIQuerySQL & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)

        
        If i <> fpCol.MaxRows Then
            sISQL = sISQL & "   -- " & sName
            
            sIQuerySQL = sIQuerySQL & "   /* " & sName & " */"
            
        Else
            sISQL = sISQL & "     -- " & sName
            
            sIQuerySQL = sIQuerySQL & "   /* " & sName & " */"
        End If
        sISQL = sISQL & vbCrLf
        sIQuerySQL = sIQuerySQL & vbCrLf
        
        
        
        'insert values
        If i <> 1 Then
            sISQLValue = sISQLValue & "       ,"
            
            sIQuerySQLValue = sIQuerySQLValue & "       ,"
            
        Else
            sISQLValue = sISQLValue & "      "
            
            sIQuerySQLValue = sIQuerySQLValue & "      "
        End If
        
        
        fpSample.Row = iSample
        fpSample.Col = i
        
        
        If sTy = "N" Or sTy = "I" Then
            sISQLValue = sISQLValue & Trim(fpSample.Value)

        Else
            sISQLValue = sISQLValue & "'" & Replace(Trim(fpSample.Value), "'", "''") & "'"

        End If
        






        If sTy = "N" Or sTy = "I" Then
            
            
            sIQuerySQLValue = sIQuerySQLValue & gfMakeString(" ", iMaxLen - Len(sCode) - 7 + iSpace)
            sIQuerySQLValue = sIQuerySQLValue & "   /* " & sName & " */"
            
        Else
            
            sIQuerySQLValue = sIQuerySQLValue & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
            sIQuerySQLValue = sIQuerySQLValue & "   /* " & sName & " */"
            
        End If
        
        
        
        sISQLValue = sISQLValue & vbCrLf
        
        sIQuerySQLValue = sIQuerySQLValue & vbCrLf
        
        
        'update field
        If i <> 1 Then
            sUSQL = sUSQL & "       ,"
        Else
            sUSQL = sUSQL & "        "
        End If
        
        sUSQL = sUSQL & "" & sCode
        sUSQL = sUSQL & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
        
        
        
        fpSample.Row = iSample
        fpSample.Col = i
        
        If sTy = "N" Or sTy = "I" Then
            sUSQL = sUSQL & "    = " & Trim(fpSample.Value)
        Else
            sUSQL = sUSQL & "    = '" & Replace(Trim(fpSample.Value), "'", "''") & "'"
        End If
        
        sUSQL = sUSQL & vbCrLf
        
        
'        'sUSQLWhere = sUSQLWhere & "   AND "
'       '
'        sUSQLWhere = sUSQLWhere & IIf(Trim(txtAlias.Text) <> "", Trim(txtAlias.Text) & ".", "") & sCode
'        sUSQLWhere = sUSQLWhere & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
'        If sTy = "N" Or sTy = "I" Then
'            sUSQLWhere = sUSQLWhere & "    = XXX"
'            sUSQLWhere = sUSQLWhere & "            -- " & sName
'        Else
'            sUSQLWhere = sUSQLWhere & "    = 'XXX'"
'            sUSQLWhere = sUSQLWhere & "          -- " & sName
'        End If
'
'        sUSQLWhere = sUSQLWhere & vbCrLf
'
        
        'SELECT FIELD
        
        If i <> 1 Then
            sSSQL = sSSQL & "       ,"
            sSSQL2 = sSSQL2 & "       ,"
            
            fpCol.Col = 5
            If Trim(fpCol.Value) <> "" Then
                sGSQL = sGSQL & "       ,"
            End If
        End If
        
'-------------------------------------
        
        
        If i = 1 Then
            sSSQL = sSSQL & " " & IIf(Trim(txtAlias.Text) <> "", Trim(txtAlias.Text) & ".", "") & sCode
            sSSQL2 = sSSQL2 & " " & IIf(Trim(txtAlias.Text) <> "", Trim(txtAlias.Text) & ".", "") & sCode
            
            fpCol.Col = 5
            If Trim(fpCol.Value) <> "" Then
                sGSQL = sGSQL & " " & sCode
            End If
                
            
        Else
            sSSQL = sSSQL & "" & IIf(Trim(txtAlias.Text) <> "", Trim(txtAlias.Text) & ".", "") & sCode
            sSSQL2 = sSSQL2 & "" & IIf(Trim(txtAlias.Text) <> "", Trim(txtAlias.Text) & ".", "") & sCode
        
            fpCol.Col = 5
            If Trim(fpCol.Value) <> "" Then
                sGSQL = sGSQL & " " & sCode
            End If
        

        
        
        End If
        
        
        
        sSSQL = sSSQL & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
        sSSQL2 = sSSQL2 & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
        
        
        
        
        'as 논리명 처리
        If chkAS.Value = 1 Then
            sSSQL = sSSQL & "   AS """ & sName & """ "
            sSSQL2 = sSSQL2 & "   AS """ & sName & """ "
        Else
            sSSQL = sSSQL & "   -- " & sName
            sSSQL2 = sSSQL2 & "   AS " & sCode
            sSSQL2 = sSSQL2 & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
            sSSQL2 = sSSQL2 & "   -- " & sName
        End If
        
                
        
        If Right(sName, 2) = "코드" Or Trim(sName) = "고객관리번호" Or Right(sName, 4) = "부점코드" Or Right(sName, 4) = "직원번호" Then
            
            sSSQL2 = sSSQL2 & vbCrLf
            sSSQL2 = sSSQL2 & "       ,"
            sSSQL2 = sSSQL2 & "" & IIf(Trim(txtAlias.Text) <> "", Trim(txtAlias.Text) & ".", "") & sCode
        
            sSSQL2 = sSSQL2 & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
            
            sSSQL2 = sSSQL2 & "   AS " & sCode & "_NM"
            sSSQL2 = sSSQL2 & gfMakeString(" ", iMaxLen - Len(sCode) - 3 + iSpace)
            sSSQL2 = sSSQL2 & "   -- " & sName & "명"
        End If

        
        
        
        
        sSSQL = sSSQL & vbCrLf
        sSSQL2 = sSSQL2 & vbCrLf
        
        fpCol.Col = 5
        If Trim(fpCol.Value) <> "" Then
            sGSQL = sGSQL & vbCrLf
        End If
        
        
    Next i

    
    sIQuerySQL = sIQuerySQL & ")"
    sIQuerySQLValue = sIQuerySQLValue & ")"
    'txtIQuery.Text = txtIQuery.Text & vbCrLf & sIQuerySQL & " VALUES " & vbCrLf & sIQuerySQLValue
    
    
    sISQL = sISQL & ")"
    sISQLValue = sISQLValue & ")"
    
    'txtInsertSQL.Text = sISQL & " VALUES " & vbCrLf & sISQLValue
    
    
    fpSample.Row = iSample
    fpSample.Col = fpSample.MaxCols
    sUSQLWhere = " WHERE ROWID = '" & fpSample.Value & "'"
    sSQL = sUSQL & sUSQLWhere
      
    'txtSelectSQL.Text = sSSQL & "  FROM " & edtTable.Text & IIf(Trim(txtAlias.Text) <> "", "  " & Trim(txtAlias.Text), "") & _
                                "        /* " & edtTablenm.Text & "*/" & vbCrLf & sUSQLWhere
                                
                                
    txtSelectSQL.Text = sSSQL & "  FROM " & edtTable.Text & IIf(strReal = "A", "@LK_STG3", "") & IIf(Trim(txtAlias.Text) <> "", "  " & Trim(txtAlias.Text), "") & _
                                "        /* " & edtTablenm.Text & "*/" & vbCrLf & sUSQLWhere
                                
    txtSQLcdnm.Text = sSSQL2 & "  FROM " & edtTable.Text & IIf(strReal = "A", "@LK_STG3", "") & IIf(Trim(txtAlias.Text) <> "", "  " & Trim(txtAlias.Text), "") & _
                                "        /* " & edtTablenm.Text & "*/" & vbCrLf & sUSQLWhere
    
    'txtInsSelect.Text = sISQL & vbCrLf & sSSQL & "  FROM " & edtTable.Text & vbCrLf & sUSQLWhere
    txtInsSelect.Text = sISQL & vbCrLf & sSSQL & "  FROM " & edtTable.Text & IIf(strReal = "A", "@LK_STG3", "") & vbCrLf & sUSQLWhere
    
'00000000000000000000000000000000000000000000000000000000000000000000000000000000000

    
    'Update한는 경우
    fpSample.Row = iSample
    fpSample.Col = 1
    If fpSample.BackColor = S_DARK_CYAN Then
    
        If bExeYn Then
            If GCONN_USE = True Then
                gConn.Execute sUSQL & sUSQLWhere
            Else
                conn.Execute sUSQL & sUSQLWhere
            End If
        Else
            txtSQL.Text = txtSQL.Text & sUSQL & sUSQLWhere & ";" & vbCrLf
        End If
    'insert한는 경우
    ElseIf fpSample.BackColor = C_ADD_COLOR Then
        If bExeYn Then
            If GCONN_USE = True Then
                gConn.Execute sISQL & " VALUES " & vbCrLf & sISQLValue
            Else
                conn.Execute sISQL & " VALUES " & vbCrLf & sISQLValue
            End If
        Else
            txtSQL.Text = txtSQL.Text & sISQL & " VALUES " & vbCrLf & sISQLValue & ";" & vbCrLf
        End If
    
    'DELETE하는 경우
    ElseIf fpSample.BackColor = S_DARK_RED Then
        If bExeYn Then
            If GCONN_USE = True Then
                gConn.Execute "DELETE FROM  " & edtTable.Text & "" & sUSQLWhere
            Else
                conn.Execute "DELETE FROM  " & edtTable.Text & "" & sUSQLWhere
            End If
        Else
            txtSQL.Text = txtSQL.Text & "DELETE FROM  " & edtTable.Text & "" & sUSQLWhere & ";" & vbCrLf
        End If
    End If

    
Next iSample
    
    
    


    If bExeYn Then
        gfMsg ("저장되었습니다.")
    End If
    
    
    If GCONN_USE = True Then
        gConn.CommitTrans
        
    Else
        conn.CommitTrans
        
    End If
        
    If bExeYn Then
        '재조회 하기
        Call cmdValueSelect_Click
    End If
    
    Screen.MousePointer = vbDefault



    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    Exit Sub
ErrorHandler:
    
    If GCONN_USE = True Then
        gConn.RollbackTrans
        
    Else
        conn.RollbackTrans
        
    End If
    
    
    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description



End Sub

Private Sub cmdMail_Click()
    WebBrowser1.Navigate ("https://mail.ibksystem.co.kr/owa/")
    
End Sub

Private Sub cmdMakeInsert_Click()



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


    Dim sType As String
    
        
    sSQL = "SELECT * FROM " & edtTable.Text
 
    txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If
    
    Dim sColumn As String
    Dim sValue As String
    Dim i As Long
    Dim j As Integer
    txtInsertData.Text = ""
    
    
    
    
    
    
    
    Dim ffile As Integer
    
    ffile = FreeFile
    
    Open "C:\" & edtTable.Text & ".sql" For Output As #ffile   ' 파일을 엽니다.
 

    
    
    
    
    
    txtInsertData.Text = "--------start--------------" & vbCrLf
    Print #ffile, "--------start--------------"
    For i = 1 To rsTemp.RecordCount
        
        sColumn = ""
        sValue = ""
        For j = 1 To rsTemp.Fields.Count
            sColumn = sColumn & rsTemp.Fields(j - 1).Name & IIf(j = rsTemp.Fields.Count, "", " ,")
                            
                            
            sValue = sValue & IIf(131 = rsTemp.Fields(j - 1).Type, "", "'")
            sValue = sValue & rsTemp.Fields(j - 1).Value
            
            If IsNull(rsTemp.Fields(j - 1).Value) Then
                sValue = sValue & "null"
                
            End If
            
            sValue = sValue & IIf(131 = rsTemp.Fields(j - 1).Type, "", "'")
            sValue = sValue & IIf(j = rsTemp.Fields.Count, "", " ,")
            
            'Debug.Print rsTemp.Fields(J - 1).Attributes & " " & rsTemp.Fields(J - 1).Type & " " & rsTemp.Fields(J - 1).Name & " " & rsTemp.Fields(J - 1).Value
            
            
        Next j
        txtInsertData.Text = txtInsertData.Text & "INSERT INTO  " & edtTable.Text & "(" & sColumn & ") values (" & sValue & ")" & ";" & vbCrLf
    
        Debug.Print "INSERT INTO  " & edtTable.Text & "(" & sColumn & ") values (" & sValue & ")" & ";" & vbCrLf
        
        Print #ffile, "INSERT INTO  " & edtTable.Text & "(" & sColumn & ") values (" & sValue & ")" & ";"
        
        rsTemp.MoveNext
    
    
    Next i
            
    
    txtInsertData.Text = txtInsertData.Text & "--------end--------------" & vbCrLf
    
    Print #ffile, "--------end--------------"
    Close #ffile   ' 파일을 닫습니다.


    
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

Private Sub cmdMig_Click()
    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    
    If GCONN_USE = False Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = GetConnString
    
        conn.Open
        conn.BeginTrans
    Else
        gConn.BeginTrans
        
    End If

    
    
    '테이블 Truncate
    If chtTruncate.Value = "1" Then
    
    
    
            If optMigSome.Value = True And Trim(txtMoveCnt.Text) = "" Then
                gfMsgWrn ("부분 이관 건수를 입력하세요!!!!")
                txtMoveCnt.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
            
            
        
            If chtTruncate.Value = "1" Then
                If gfMsgQst("개발계 DB가 모두 삭제 후 이관 됩니다. 진행하시겠습니까?") = vbNo Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            Else
                If gfMsgQst("Delete all 하지 않은 경우 Duplicate 오류가 발생 할수 있습니다. 진행하시겠습니까?") = vbNo Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
            
            Screen.MousePointer = vbHourglass
            

        
        If chkNoExe.Value <> "1" Then
            

        
        
             sSQL = ""
             sSQL = sSQL & vbCrLf & " DELETE FROM " & Trim(txtTgtOwner.Text) & "." & Trim(txtTgtTableId.Text)
             
            
             If GCONN_USE = True Then
                 gConn.Execute sSQL
             Else
                 conn.Execute sSQL
             End If
        End If
    
    End If

    sSQL = fGetSelectInsertSQL
    
    
    
    
    sSQL = sSQL & vbCrLf
    If optMigSome.Value = True Then
        sSQL = sSQL & " WHERE ROWNUM  <= " & txtMoveCnt.Text
    End If
    
    
    txtMigSQL.Text = sSQL
    
    
    If chkNoExe.Value <> "1" Then
    
        If GCONN_USE = True Then
            gConn.Execute sSQL
        Else
            conn.Execute sSQL
        End If
        
        If GCONN_USE = True Then
            gConn.CommitTrans
        Else
            conn.CommitTrans
        End If
        gfMsg ("등록되었습니다.")
    Else
    
        gfMsg ("SQL문 생성 되었습니다.")
    End If
    
    
    
    
    
    

    

    
    
    Screen.MousePointer = vbDefault

    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    Exit Sub
ErrorHandler:
    
    If GCONN_USE = True Then
        gConn.RollbackTrans
        
    Else
        conn.RollbackTrans
        
    End If
    
    
    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description
'
End Sub
Private Sub cmdMinMax_Click(Index As Integer)
    If Index = 0 Then
        SSTab2.Top = 5520
        SSTab2.Height = 9255
        fpSample.Height = 8655
    Else
    
        SSTab2.Top = 11040
        SSTab2.Height = 3855
        fpSample.Height = 3375
    
    End If
End Sub



'---------------------------------------
'이행 DB 이관용 selec insert sql
'-----------------------------------------
Private Function fGetSelectInsertSQL()
    Dim i As Integer
    
    
    Dim sCode As String
    Dim sName As String
    
    Dim iMaxLen As Integer
    iMaxLen = 0
    Dim iSpace As Integer
    iSpace = 10
    Dim sTy As String
    
    '최대 컬럼 사이즈 계산
    For i = 1 To fpCol.MaxRows
        fpCol.Row = i
        fpCol.Col = 1
        
        If Len(Trim(fpCol.Value)) > iMaxLen Then
            iMaxLen = Len(Trim(fpCol.Value))
        End If
        
    Next i
    
    
    Dim sISQL As String
    Dim sISQLValue As String
    sISQLValue = ""
    
    
    Dim sIQuerySQL As String
    Dim sIQuerySQLValue As String
    sIQuerySQLValue = ""
    
    Dim sUSQL As String
    Dim sUSQLWhere As String
    sUSQLWhere = ""
    
    Dim sSSQL As String
    Dim sSSQL2 As String
    
    
    Dim sGSQL As String
    
    sISQL = "INSERT INTO " & Trim(txtTgtOwner.Text) & "." & Trim(txtTgtTableId.Text) & " ( " & "        /* " & txtTgtTableId.Text & "*/" & vbCrLf
    sISQLValue = "( "
    
    
    sIQuerySQL = "INSERT INTO " & edtTable.Text & " ( " & vbCrLf
    sIQuerySQLValue = "( "
    
    sUSQL = "UPDATE " & edtTable.Text & "        /* " & edtTablenm.Text & "*/" & vbCrLf & "   SET " & vbCrLf
    sUSQLWhere = " WHERE 1=1" & vbCrLf
    
    
    sSSQL = "SELECT "
    sSSQL2 = "SELECT "
    
    sGSQL = " "
    

    
    For i = 1 To fpCol.MaxRows
        fpCol.Row = i
        fpCol.Col = 1
        sCode = Trim(fpCol.Value)
        fpCol.Col = 2
        sName = Trim(fpCol.Value)
        fpCol.Col = 3
        sTy = Left(fpCol.Value, 1)
        

        
        'insert field
        If i <> 1 Then
            sISQL = sISQL & "         ,"
            sIQuerySQL = sIQuerySQL & "         ,"
        Else
            sISQL = sISQL & "          "
            sIQuerySQL = sIQuerySQL & "          "
        End If
        
        
        sISQL = sISQL & "" & sCode
        sISQL = sISQL & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
        
        
        sIQuerySQL = sIQuerySQL & "" & sCode
        sIQuerySQL = sIQuerySQL & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)

        
        If i <> fpCol.MaxRows Then
            sISQL = sISQL & "   -- " & sName
            
            sIQuerySQL = sIQuerySQL & "   /* " & sName & " */"
            
        Else
            sISQL = sISQL & "     -- " & sName
            
            sIQuerySQL = sIQuerySQL & "   /* " & sName & " */"
        End If
        sISQL = sISQL & vbCrLf
        sIQuerySQL = sIQuerySQL & vbCrLf
        
        
        
        'insert values
        If i <> 1 Then
            sISQLValue = sISQLValue & "       ,"
            
            sIQuerySQLValue = sIQuerySQLValue & "       ,"
            
        Else
            sISQLValue = sISQLValue & "      "
            
            sIQuerySQLValue = sIQuerySQLValue & "      "
        End If
        
        
        
        
        If sTy = "N" Or sTy = "I" Then
            sISQLValue = sISQLValue & "XXX"
            sISQLValue = sISQLValue & "            -- " & sName
        Else
            sISQLValue = sISQLValue & "'XXX'"
            sISQLValue = sISQLValue & "          -- " & sName
        End If
        
        
        
        'IQuery의 경우
        
        If sCode = "SYS_REG_DTTM" Then
            sIQuerySQLValue = sIQuerySQLValue & "SYSDATE"
        
        ElseIf sCode = "SYS_REG_BR_CD" Then
            sIQuerySQLValue = sIQuerySQLValue & "#SYS_BR_CD#  "
        
        ElseIf sCode = "SYS_RGSR_EMNO" Then
            sIQuerySQLValue = sIQuerySQLValue & "#SYS_EMNO# "
        
        ElseIf sCode = "SYS_CHG_DTTM" Then
            sIQuerySQLValue = sIQuerySQLValue & "SYSDATE"
        
        ElseIf sCode = "SYS_CHG_BR_CD" Then
            sIQuerySQLValue = sIQuerySQLValue & "#SYS_BR_CD#"
        
        ElseIf sCode = "SYS_EDIR_EMNO" Then
            sIQuerySQLValue = sIQuerySQLValue & "#SYS_EMNO#"
        
        ElseIf sTy = "N" Or sTy = "I" Then
           
            sIQuerySQLValue = sIQuerySQLValue & "NVL(#" & sCode & "#,0)"
            
        Else
            
            sIQuerySQLValue = sIQuerySQLValue & "#" & sCode & "#"
            
        End If
        

        If sTy = "N" Or sTy = "I" Then
            
            
            sIQuerySQLValue = sIQuerySQLValue & gfMakeString(" ", iMaxLen - Len(sCode) - 7 + iSpace)
            sIQuerySQLValue = sIQuerySQLValue & "   /* " & sName & " */"
            
        Else
            
            sIQuerySQLValue = sIQuerySQLValue & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
            sIQuerySQLValue = sIQuerySQLValue & "   /* " & sName & " */"
            
        End If
        
        
        
        sISQLValue = sISQLValue & vbCrLf
        
        sIQuerySQLValue = sIQuerySQLValue & vbCrLf
        
        
        'update field
        If i <> 1 Then
            sUSQL = sUSQL & "       ,"
        Else
            sUSQL = sUSQL & "        "
        End If
        
        sUSQL = sUSQL & "" & sCode
        sUSQL = sUSQL & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
        
        
        If sTy = "N" Or sTy = "I" Then
            sUSQL = sUSQL & "    = XXX"
            sUSQL = sUSQL & "            -- " & sName
        Else
            sUSQL = sUSQL & "    = 'XXX'"
            sUSQL = sUSQL & "          -- " & sName
        End If
        
        sUSQL = sUSQL & vbCrLf
        
        'where
        'If Trim(sUSQLWhere) = "" Then
        '    sUSQLWhere = sUSQLWhere & " WHERE "
        'Else
        '    sUSQLWhere = sUSQLWhere & "   AND "
        'End If
            sUSQLWhere = sUSQLWhere & "   AND "
        
        sUSQLWhere = sUSQLWhere & IIf(Trim(txtAlias.Text) <> "", Trim(txtAlias.Text) & ".", "") & sCode
        sUSQLWhere = sUSQLWhere & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
        If sTy = "N" Or sTy = "I" Then
            sUSQLWhere = sUSQLWhere & "    = XXX"
            sUSQLWhere = sUSQLWhere & "            -- " & sName
        Else
            sUSQLWhere = sUSQLWhere & "    = 'XXX'"
            sUSQLWhere = sUSQLWhere & "          -- " & sName
        End If
        
        sUSQLWhere = sUSQLWhere & vbCrLf
        
        
        'SELECT FIELD
        
        If i <> 1 Then
            sSSQL = sSSQL & "       ,"
            sSSQL2 = sSSQL2 & "       ,"
            
            fpCol.Col = 5
            If Trim(fpCol.Value) <> "" Then
                sGSQL = sGSQL & "       ,"
            End If
        End If
        
'-------------------------------------
        
        
        If i = 1 Then
            sSSQL = sSSQL & " " & IIf(Trim(txtAlias.Text) <> "", Trim(txtAlias.Text) & ".", "") & sCode
            sSSQL2 = sSSQL2 & " " & IIf(Trim(txtAlias.Text) <> "", Trim(txtAlias.Text) & ".", "") & sCode
            
            fpCol.Col = 5
            If Trim(fpCol.Value) <> "" Then
                sGSQL = sGSQL & " " & sCode
            End If
                
            
        Else
            sSSQL = sSSQL & "" & IIf(Trim(txtAlias.Text) <> "", Trim(txtAlias.Text) & ".", "") & sCode
            sSSQL2 = sSSQL2 & "" & IIf(Trim(txtAlias.Text) <> "", Trim(txtAlias.Text) & ".", "") & sCode
        
            fpCol.Col = 5
            If Trim(fpCol.Value) <> "" Then
                sGSQL = sGSQL & " " & sCode
            End If
        

        
        
        End If
        
        
        
        sSSQL = sSSQL & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
        sSSQL2 = sSSQL2 & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
        
        
        
        
        'as 논리명 처리
        If chkAS.Value = 1 Then
            sSSQL = sSSQL & "   AS """ & sName & """ "
            sSSQL2 = sSSQL2 & "   AS """ & sName & """ "
        Else
            sSSQL = sSSQL & "   -- " & sName
            sSSQL2 = sSSQL2 & "   AS " & sCode
            sSSQL2 = sSSQL2 & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
            sSSQL2 = sSSQL2 & "   -- " & sName
        End If
        
        
        
        
'-----------------------------------


  
                
                
        
        If Right(sName, 2) = "코드" Or Trim(sName) = "고객관리번호" Or Right(sName, 4) = "부점코드" Or Right(sName, 4) = "직원번호" Then
            
            sSSQL2 = sSSQL2 & vbCrLf
            sSSQL2 = sSSQL2 & "       ,"
            sSSQL2 = sSSQL2 & "" & IIf(Trim(txtAlias.Text) <> "", Trim(txtAlias.Text) & ".", "") & sCode
        
            sSSQL2 = sSSQL2 & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
            
            sSSQL2 = sSSQL2 & "   AS " & sCode & "_NM"
            sSSQL2 = sSSQL2 & gfMakeString(" ", iMaxLen - Len(sCode) - 3 + iSpace)
            sSSQL2 = sSSQL2 & "   -- " & sName & "명"
        End If

        
        
        
        
        sSSQL = sSSQL & vbCrLf
        sSSQL2 = sSSQL2 & vbCrLf
        
        fpCol.Col = 5
        If Trim(fpCol.Value) <> "" Then
            sGSQL = sGSQL & vbCrLf
        End If
        
        
    Next i
    
    

    
    'sIQuerySQL = sIQuerySQL & ")"
    'sIQuerySQLValue = sIQuerySQLValue & ")"
    'txtIQuery.Text = txtIQuery.Text & vbCrLf & sIQuerySQL & " VALUES " & vbCrLf & sIQuerySQLValue
    
    
    sISQL = sISQL & ")"
    sISQLValue = sISQLValue & ")"
    
    'txtInsertSQL.Text = sISQL & " VALUES " & vbCrLf & sISQLValue
    
    'sISQL = sISQL & " VALUES " & vbCrLf & sISQLValue
    
    
    'txUpdateSQL.Text = sUSQL & sUSQLWhere
      
    'txtSelectSQL.Text = sSSQL & "  FROM " & edtTable.Text & IIf(Trim(txtAlias.Text) <> "", "  " & Trim(txtAlias.Text), "") & _
                                "        /* " & edtTablenm.Text & "*/" & vbCrLf & sUSQLWhere
                                
                                
    'txtSelectSQL.Text = sSSQL & "  FROM " & edtTable.Text & IIf(chkASIS.Value = 1, "@DL_IBKCDEV", "") & IIf(Trim(txtAlias.Text) <> "", "  " & Trim(txtAlias.Text), "") & _
                                "        /* " & edtTablenm.Text & "*/" & vbCrLf & sUSQLWhere
                                
    'txtSQLcdnm.Text = sSSQL2 & "  FROM " & edtTable.Text & IIf(chkASIS.Value = 1, "@DL_IBKCDEV", "") & IIf(Trim(txtAlias.Text) <> "", "  " & Trim(txtAlias.Text), "") & _
                                "        /* " & edtTablenm.Text & "*/" & vbCrLf & sUSQLWhere
    
    'txtInsSelect.Text = sISQL & vbCrLf & sSSQL & "  FROM " & edtTable.Text & vbCrLf & sUSQLWhere
    Dim sRet As String
    sRet = sISQL & vbCrLf & sSSQL & "  FROM " & Trim(txtMigOwner.Text) & "." & Trim(txtMigTableID.Text) & "@LK_STG3" '-- & vbCrLf & sUSQLWhere
    
    fGetSelectInsertSQL = sRet

End Function



Private Sub cmdNeed_Click()
    

On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    
    Dim i As Integer
    Dim j As Integer
    
    txtNeed.Text = ""
    txtLogger.Text = ""
    
    Dim sColid As String
    Dim sColnm As String
    Dim sCodeDomain As String
    Dim sColBigo As String
    Dim sColBigo2 As String
    Dim sType As String
    Dim sTxt As String
    
    
    For i = 1 To fpCol.MaxRows
        
        fpCol.Row = i
        fpCol.Col = 1: sColid = Trim(fpCol.Value)
        fpCol.Col = 2: sColnm = fpCol.Value
        fpCol.Col = 3: sType = fpCol.Value
        fpCol.Col = 6: sCodeDomain = fpCol.Value
        fpCol.Col = 7: sColBigo = fpCol.Value
        fpCol.Col = 10: sColBigo2 = fpCol.Value
        If InStr(1, sColBigo2, "필수") > 0 Then
        
            
            txtNeed.Text = txtNeed.Text & "//" & sColnm & vbCrLf
            
            If InStr(1, sType, "CHAR") > 0 Then
                
                txtNeed.Text = txtNeed.Text & "if(StringUtil.isEmpty(vo.get" & sColid & "())) {" & vbCrLf
                            
            Else
                txtNeed.Text = txtNeed.Text & "if(vo.get" & sColid & "() == null || vo.get" & sColid & "().doubleValue() <= 0) {" & vbCrLf
            
            End If
                        
            txtNeed.Text = txtNeed.Text & "    throw new BizException(MsgConst.EESY0028, """ & sColnm & """);" & vbCrLf
            txtNeed.Text = txtNeed.Text & "}" & vbCrLf
            txtNeed.Text = txtNeed.Text & "" & vbCrLf
            
            
        End If
        
        sTxt = "logger.debug(""" & sColnm & "[" & sColid & "]"
        sTxt = Trim(sTxt)
        sTxt = sTxt & gfMakeString(" ", 70 - nLeft(sTxt) + 1) & ": "" + "
        sTxt = sTxt & "formObject.get" & sColid & "()"
        sTxt = sTxt & gfMakeString(" ", 130 - nLeft(sTxt) + 1) & ");"

        
        txtLogger.Text = txtLogger.Text & sTxt & vbCrLf
    
        
    'gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
    Next i
    
    frmNeed.Visible = True
    Screen.MousePointer = vbDefault
    Exit Sub
ErrorHandler:
    
   
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description
End Sub

'테이블 정의서 출력
Private Sub cmdPrint_Click()




    If Trim(edtTable.Text) = "" Then
        gfMsg "테이블을 선택하십시오"
        Exit Sub
    End If
    
    fpCol.Col = 9
    fpCol.ColHidden = True
    fpCol.Col = 8
    fpCol.ColHidden = True

    Call subSpreadSmart(PrintOrientationPortrait, 50, 50, fpCol, edtTable.Text & "      " & edtTablenm.Text, "")
    
    
    fpCol.Col = 9
    fpCol.ColHidden = False
    fpCol.Col = 8
    fpCol.ColHidden = False
End Sub

Private Sub cmdReload_Click()
    Call subSPlistSetup
End Sub


'코드 저장
Private Sub cmdSave_Click()

    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    
    If GCONN_USE = False Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = GetConnString
    
        conn.Open
        conn.BeginTrans
    Else
        gConn.BeginTrans
        
    End If
    
    Dim i As Integer
    
    
    
    Dim sCODDOM As String
    Dim sCDVAL       As String
    Dim sCOMDOMNM    As String
    Dim sCDVALNM     As String
    Dim sBIGO        As String
       
    
    
    
    
    
    '전체 삭제
    fpcode.Row = 1
    fpcode.Col = 1
    
    sCODDOM = Trim(fpcode.Value)
    
    If sCODDOM <> "" Then
        
        sSQL = ""
        sSQL = sSQL & vbCrLf & " DELETE FROM ln.TB_DIC_COD "
        sSQL = sSQL & vbCrLf & "  WHERE CODDOM  = '" & sCODDOM & "'"
        
    
        txtSQL.Text = sSQL
        If GCONN_USE = True Then
            gConn.Execute sSQL
        Else
            conn.Execute sSQL
        End If
        
    '채번
    Else
    
            sSQL = ""
            sSQL = sSQL & vbCrLf & "SELECT 'XX' || LTRIM(TO_CHAR( TO_NUMBER(SUBSTR(MAX(CODDOM) ,3,4)) + 1, '0000'))"
            sSQL = sSQL & vbCrLf & " From Ln.TB_DIC_COD "
    
            Set rsTemp = New ADODB.Recordset
            rsTemp.CursorLocation = adUseClient

       
            txtSQL.Text = sSQL
            If GCONN_USE = True Then
                rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
            Else
                rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
            End If
        
            
            sCODDOM = rsTemp.Fields(0).Value
    
    End If
    
    
    
    
    fpcode.Col = 2: sCOMDOMNM = fpcode.Value
        
    
    For i = 1 To fpcode.MaxRows
        
        fpcode.Row = i
        fpcode.Col = 4: sCDVAL = Trim(fpcode.Value)
        fpcode.Col = 5: sCDVALNM = Trim(fpcode.Value)
        fpcode.Col = 7: sBIGO = Trim(fpcode.Value)
    
        
        If sCDVAL <> "" And sCDVALNM = "" Then
            
            gfMsgInf (i & "행 코드명이 누락되었습니다.")
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        
        If sCDVAL = "" And sCDVALNM <> "" Then
            
            gfMsgInf (i & "행 코드값이 누락되었습니다.")
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        If sCDVAL <> "" And sCDVALNM <> "" Then
            sSQL = ""
            sSQL = sSQL & vbCrLf & " INSERT INTO ln.TB_DIC_COD (CODDOM, COMDOMNM,CDVAL, CDVALNM, BIGO) "
            sSQL = sSQL & vbCrLf & "  VALUES ('" & sCODDOM & "' ,"
            sSQL = sSQL & vbCrLf & "          '" & sCOMDOMNM & "' ,"
            sSQL = sSQL & vbCrLf & "          '" & sCDVAL & "' ,"
            sSQL = sSQL & vbCrLf & "          '" & sCDVALNM & "' ,"
            sSQL = sSQL & vbCrLf & "          '" & sBIGO & "' )"
            
            
            txtSQL.Text = sSQL
            If GCONN_USE = True Then
                gConn.Execute sSQL
            Else
                conn.Execute sSQL
            End If
        End If
    Next i
    
    
    gfMsg ("등록되었습니다.")

    'commit
    If GCONN_USE = True Then
        gConn.CommitTrans
    Else
        conn.CommitTrans
    End If
    
   
    '재조회
    subSelectCodeList (sCODDOM)
    
    
    Screen.MousePointer = vbDefault

    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    Exit Sub
ErrorHandler:
    
    If GCONN_USE = True Then
        gConn.RollbackTrans
        
    Else
        conn.RollbackTrans
        
    End If
    
    
    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description


End Sub














Private Sub cmdSock_Click()
    frmSock.Show
End Sub

Private Sub cmdSPSearch_Click()
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


    
    If Trim(txtSPconect.Text <> "") Then
    
        sSQL = ""
        sSQL = sSQL & vbCrLf & "  SELECT  DISTINCT OBJECT_NAME   "
        sSQL = sSQL & vbCrLf & "         ,OBJECT_TYPE            "
        sSQL = sSQL & vbCrLf & "  FROM USER_PROCEDURES A         "
        sSQL = sSQL & vbCrLf & "    ,USER_SOURCE   B             "
        sSQL = sSQL & vbCrLf & " Where 1 = 1                     "
        sSQL = sSQL & vbCrLf & " AND OBJECT_NAME = NAME          "
        sSQL = sSQL & vbCrLf & " AND TEXT LIKE '%" & Trim(txtSPconect.Text) & "%'"
        sSQL = sSQL & vbCrLf & " ORDER BY OBJECT_TYPE, OBJECT_NAME  "
        
    
    Else
        sSQL = ""
        sSQL = sSQL & vbCrLf & "SELECT OBJECT_NAME, OBJECT_TYPE    "
        sSQL = sSQL & vbCrLf & "  FROM USER_PROCEDURES   "
        sSQL = sSQL & vbCrLf & " Where 1 = 1                     "
        
        If Trim(txtSPname.Text) <> "" Then
            sSQL = sSQL & vbCrLf & " AND OBJECT_NAME LIKE '%" & Trim(UCase(txtSPname.Text)) & "%'"
        End If
        sSQL = sSQL & vbCrLf & " ORDER BY OBJECT_TYPE, OBJECT_NAME  "
    
    End If
   
    txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If

    fpSP.MaxRows = 0
    If rsTemp.RecordCount < 1 Then
        fpSP.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpSP, rsTemp)
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

Private Sub cmdTableSrch_Click()
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


Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", "OWNER_LIST", "OWNER", UCase(cboShema.Text))

'
'If strReal <> "T" And strReal <> "I" Then
'
'    sSQL = sSQL & vbCrLf & "SELECT TABLE_NAME, TABLE_COMM, BIGO1, TABLESPACE_NAME "
'    sSQL = sSQL & vbCrLf & "  FROM ("
'            sSQL = sSQL & vbCrLf & "SELECT  TABLE_NAME                                      "
'            sSQL = sSQL & vbCrLf & "       ,TABLE_COMM                                        "
'            sSQL = sSQL & vbCrLf & "       , BIGO1                                           "
'
'            sSQL = sSQL & vbCrLf & "       ,CASE WHEN TABLE_COMM IS NULL THEN DIC_COMM      "
'            sSQL = sSQL & vbCrLf & "             ELSE TABLE_COMM                            "
'            sSQL = sSQL & vbCrLf & "        END TABLE_COMM                                  "
'            sSQL = sSQL & vbCrLf & "       ,TABLESPACE_NAME                                 "
'            sSQL = sSQL & vbCrLf & "  FROM (                                                "
'            sSQL = sSQL & vbCrLf & "        SELECT  A.TABLE_NAME                            "
'            sSQL = sSQL & vbCrLf & "               ,(SELECT B.COMMENTS                      "
'            sSQL = sSQL & vbCrLf & "                   FROM ALL_TAB_COMMENTS" & IIf(chkASIS.Value = 1, "@DL_IBKCDEV", "") & " B              "
'            sSQL = sSQL & vbCrLf & "                  WHERE A.OWNER = B.OWNER               "
'            sSQL = sSQL & vbCrLf & "                   AND A.TABLE_NAME = B.TABLE_NAME      "
'            sSQL = sSQL & vbCrLf & "                ) DIC_COMM                              "
'            sSQL = sSQL & vbCrLf & "               ,C.TABLE_COMM                            "
'            sSQL = sSQL & vbCrLf & "               ,C.BIGO1                                 "
'            sSQL = sSQL & vbCrLf & "               ,C.BIGO2                                 "
'            sSQL = sSQL & vbCrLf & "               ,A.TABLESPACE_NAME                       "
'            sSQL = sSQL & vbCrLf & "          FROM ALL_TABLES" & IIf(chkASIS.Value = 1, "@DL_IBKCDEV", "") & " A                             "
'            sSQL = sSQL & vbCrLf & "                   LEFT OUTER JOIN                      "
'            sSQL = sSQL & vbCrLf & "               ln.TB_DIC_TAB" & IIf(chkASIS.Value = 1, "@DL_IBKCDEV", IIf(strReal = "R", "@DL_DEOSDB", "")) & " C  "
'            sSQL = sSQL & vbCrLf & "                   ON A.OWNER      = C.OWNER            "
'            sSQL = sSQL & vbCrLf & "                  AND A.TABLE_NAME = C.TABLE_NAME       "
'            sSQL = sSQL & vbCrLf & "         WHERE A.OWNER = '" & IIf(chkASIS.Value = 1, "IBKCDEV", UCase(cboShema.Text)) & "'"
'            sSQL = sSQL & vbCrLf & "        )                                               "
'    sSQL = sSQL & vbCrLf & "         )"
'    sSQL = sSQL & vbCrLf & "  WHERE 1=1"
'
'    If Trim(cboSubject.Text) <> "" Then
'        sSQL = sSQL & vbCrLf & " AND UPPER(TABLE_NAME) LIKE 'T" & cboSubject.Text & "%'"
'    End If
'    If Trim(txtTableSrchEng.Text) <> "" Then
'        sSQL = sSQL & vbCrLf & " AND UPPER(TABLE_NAME) LIKE '%" + UCase(Trim(txtTableSrchEng.Text)) + "%'"
'    ElseIf Trim(txtTableSrchKor.Text) <> "" Then
'        sSQL = sSQL & vbCrLf & " AND  UPPER(TABLE_COMM)  LIKE '%" + UCase(Trim(txtTableSrchKor.Text)) + "%'"
'    ElseIf Trim(txtTableSrchBigo.Text) <> "" Then
'        sSQL = sSQL & vbCrLf & " AND  UPPER(BIGO1)  LIKE '%" + UCase(Trim(txtTableSrchBigo.Text)) + "%'"
'    End If
'
'Else

    sSQL = sSQL & vbCrLf & "SELECT TABLE_NAME, TABLE_COMM,  BIGO1, TABLESPACE_NAME, OWNER "
    sSQL = sSQL & vbCrLf & "  FROM ("
            sSQL = sSQL & vbCrLf & "SELECT  TABLE_NAME AS TABLE_NAME                                   "
            sSQL = sSQL & vbCrLf & "       ,DIC_COMM   AS TABLE_COMM                                  "
            sSQL = sSQL & vbCrLf & "       ,' '        AS BIGO1                                           "
            sSQL = sSQL & vbCrLf & "       ,TABLESPACE_NAME                                 "
            sSQL = sSQL & vbCrLf & "       ,OWNER                                 "
            sSQL = sSQL & vbCrLf & "  FROM (                                                "
            sSQL = sSQL & vbCrLf & "        SELECT  A.TABLE_NAME                            "
            sSQL = sSQL & vbCrLf & "               ,(SELECT B.COMMENTS                      "
            sSQL = sSQL & vbCrLf & "                   FROM ALL_TAB_COMMENTS" & IIf(strReal = "A", "@LK_STG3", "") & " B              "
            sSQL = sSQL & vbCrLf & "                  WHERE A.OWNER = B.OWNER               "
            sSQL = sSQL & vbCrLf & "                   AND A.TABLE_NAME = B.TABLE_NAME      "
            sSQL = sSQL & vbCrLf & "                ) DIC_COMM                              "
            'sSQL = sSQL & vbCrLf & "               ,C.TABLE_COMM                            "
            'sSQL = sSQL & vbCrLf & "               ,C.BIGO1                                 "
            'sSQL = sSQL & vbCrLf & "               ,C.BIGO2                                 "
            sSQL = sSQL & vbCrLf & "               ,A.TABLESPACE_NAME                       "
            sSQL = sSQL & vbCrLf & "               ,A.OWNER                       "
            sSQL = sSQL & vbCrLf & "          FROM ALL_TABLES" & IIf(strReal = "A", "@LK_STG3", "") & " A                             "
            'sSQL = sSQL & vbCrLf & "                   LEFT OUTER JOIN                      "
            'sSQL = sSQL & vbCrLf & "               ln.TB_DIC_TAB" & IIf(chkASIS.Value = 1, "@DL_IBKCDEV", IIf(strReal = "1", "@DL_DEOSDB", "")) & " C                             "
            'sSQL = sSQL & vbCrLf & "                   ON A.OWNER      = C.OWNER            "
            'sSQL = sSQL & vbCrLf & "                  AND A.TABLE_NAME = C.TABLE_NAME       "
            If Trim(cboShema.Text) <> "" Then
                sSQL = sSQL & vbCrLf & "         WHERE A.OWNER = '" & IIf(chkASIS.Value = 1, UCase(cboShema.Text), UCase(cboShema.Text)) & "'"
            Else
                sSQL = sSQL & vbCrLf & "         WHERE A.OWNER      NOT IN  " & C_NOTIN
            End If
            sSQL = sSQL & vbCrLf & "        )                                               "
            
    sSQL = sSQL & vbCrLf & "         )"
    sSQL = sSQL & vbCrLf & "  WHERE 1=1"

'엑셀에서 복사한 경우 enter key 제거
txtTableSrchEng.Text = Replace(txtTableSrchEng.Text, vbCrLf, "")

    If Trim(cboSubject.Text) <> "" Then
        sSQL = sSQL & vbCrLf & " AND UPPER(TABLE_NAME) LIKE 'T" & cboSubject.Text & "%'"
    End If
    If Trim(txtTableSrchEng.Text) <> "" Then
        sSQL = sSQL & vbCrLf & " AND (       UPPER(TABLE_NAME) LIKE '%" + UCase(Trim(txtTableSrchEng.Text)) + "%'"
        sSQL = sSQL & vbCrLf & "         OR  UPPER(TABLE_COMM)  LIKE '%" + UCase(Trim(txtTableSrchEng.Text)) + "%' )"
    ElseIf Trim(txtTableSrchKor.Text) <> "" Then
        sSQL = sSQL & vbCrLf & " AND  UPPER(TABLE_COMM)  LIKE '%" + UCase(Trim(txtTableSrchEng.Text)) + "%'"
    ElseIf Trim(txtTableSrchBigo.Text) <> "" Then
        sSQL = sSQL & vbCrLf & " AND  UPPER(BIGO1)  LIKE '%" + UCase(Trim(txtTableSrchBigo.Text)) + "%'"
    End If
'End If


    sSQL = sSQL & vbCrLf & " ORDER BY TABLE_NAME"
    
   
    txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If

    fpTable.MaxRows = 0
    If rsTemp.RecordCount < 1 Then
        fpTable.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpTable, rsTemp)
    End If
    
    If rsTemp.RecordCount = 1 Then
        Call fpTable_DblClick(1, 1)
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
'------------------------------------------
'테이블 명 갱신
'------------------------------------------
Private Sub cmdTableUpdate_Click()
    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    
    If GCONN_USE = False Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = GetConnString
    
        conn.Open
        conn.BeginTrans
    Else
        gConn.BeginTrans
        
    End If
    



    sSQL = ""
    sSQL = sSQL & vbCrLf & " DELETE FROM ln.TB_DIC_TAB" & IIf(strReal = "A", "@LK_STG3", "") & " "
    sSQL = sSQL & vbCrLf & "  WHERE TABLE_NAME     = '" & edtTable.Text & "'"
    If Trim(cboShema.Text) <> "" Then
        sSQL = sSQL & vbCrLf & "    AND OWNER  = '" & IIf(chkASIS.Value = 1, "IBKCDEV", UCase(cboShema.Text)) & "'"
    Else
        sSQL = sSQL & vbCrLf & "    AND OWNER      NOT IN  " & C_NOTIN
    End If
    
    If GCONN_USE = True Then
        gConn.Execute sSQL
    Else
        conn.Execute sSQL
    End If
    

    
    sSQL = ""
    sSQL = sSQL & vbCrLf & " INSERT INTO ln.TB_DIC_TAB (OWNER, TABLE_NAME, TABLE_COMM, BIGO1, BIGO2) "
    sSQL = sSQL & vbCrLf & "  VALUES ('" & cboShema.Text & "' ,"
    sSQL = sSQL & vbCrLf & "          '" & edtTable.Text & "' ,"
    sSQL = sSQL & vbCrLf & "          '" & Trim(edtTablenm.Text) & "' ,"
    sSQL = sSQL & vbCrLf & "          '" & Trim(edtTableBigo.Text) & "' ,"
    sSQL = sSQL & vbCrLf & "          '" & " " & "' )"
    
    

    If GCONN_USE = True Then
        gConn.Execute sSQL
    Else
        conn.Execute sSQL
    End If
    
    
    If GCONN_USE = True Then
        gConn.CommitTrans
    Else
        conn.CommitTrans
    End If
    
    
    gfMsg ("등록되었습니다.")
    

    'commit
    cmdTableSrch_Click
    
    
    Screen.MousePointer = vbDefault

    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    Exit Sub
ErrorHandler:
    
    If GCONN_USE = True Then
        gConn.RollbackTrans
        
    Else
        conn.RollbackTrans
        
    End If
    
    
    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description
'
End Sub

Private Sub cmdValueSelect_Click()
    Call subSelectData("")
End Sub

Private Sub cmdZCamp_Click()
    WebBrowser1.Navigate ("http://portal.ibk.co.kr/Kplusweb/index.aspx")

End Sub

Private Sub cmdViewSrc_Click()
   
    
    

On Error GoTo ErrorHandler
    
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim sSQL As String
    Screen.MousePointer = vbHourglass
    

   'frmSrc.Visible = True



    sSQL = ""
    sSQL = sSQL & vbCrLf & "  select REPLACE(REPLACE(TEXT, CHR(9), '') , CHR(13), '') AS TEXT                       "
    sSQL = sSQL & vbCrLf & "   from user_SOURCE" & IIf(strReal = "A", "@LK_STG3", "") & "              "
    sSQL = sSQL & vbCrLf & "   WHERE 1=1  "
    'sSQL = sSQL & vbCrLf & "    AND OWNER = '" & cboShema.Text & "' "
    sSQL = sSQL & vbCrLf & "      AND NAME = '" & cmbPgmIDLog.Text & "'  "
    sSQL = sSQL & vbCrLf & "   ORDER BY LINE  "
    
    
    Set rsTemp = gfGlobalSQL(sSQL)
    
    rsTemp.MoveFirst
    
    txtSrc.Text = ""
    Dim sTemp As String
    
    For i = 1 To rsTemp.RecordCount
        
        sTemp = rsTemp.Fields(0).Value
        sTemp = Replace(sTemp, Chr(10), "")
        sTemp = Replace(sTemp, Chr(13), "")
        
        txtSrc.Text = txtSrc.Text & sTemp & vbCrLf
        rsTemp.MoveNext
    Next i
   
    
    frmSrc.Show
    frmSrc.txtSrc.Text = txtSrc.Text
    
    
        
    
    
    Screen.MousePointer = vbDefault

    Set rsTemp = Nothing
    
    
    Exit Sub
ErrorHandler:

    Set rsTemp = Nothing
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description
'
    
    
End Sub

Private Sub Command3_Click()
    Clipboard.Clear
    Clipboard.SetText (txtMapping.Text)
End Sub

Private Sub Command4_Click()
    WebBrowser3.Navigate ("http://db2.ibk.co.kr/db/index.htm") '차세대
End Sub




Private Sub Command5_Click()
    Clipboard.Clear
    Clipboard.SetText (txtTabMapping.Text)
End Sub

Private Sub Command6_Click()
    'SQL display
    Call subMakeSQL
End Sub




Private Sub Form_Activate()
    'WebBrowser1.Navigate ("http://iworld.ibksystem.co.kr/iFrame/out/main.xml")
    WebBrowser1.Navigate ("http://172.16.1.184:9800/")
    'WebBrowser1.Navigate ("http://portal.ibk.co.kr")
    'WebBrowser2.Navigate ("http://134.100.207.40/fine_db/EDW/table_group.asp")  'EDW
    WebBrowser2.Navigate ("http://pms.ibksystem.co.kr/iFrame?Class=com.ibks.cmmi.com.LoginCheck")  'PMS
    WebBrowser3.Navigate ("http://db2.ibk.co.kr/db/index.htm") '차세대
    
    txtTableSrchEng.SetFocus
    
    
    '최근 검색 불러오기
    Call ReadFile
    
    '내 owner 지정
    cboShema.Text = GetSetting("ORAIB_TAB_MAN_BY_HAK", "DB_OWNER", "OWNER_" & cboSystemList.Text, "")

End Sub

Private Sub Form_Load()
    '콤보박스 셋팅
'    If strReal <> "I" Then
'        cboShema.AddItem (strUser)
'    Else
'        cboShema.AddItem ("ODS_TBOWN")
'    End If
'    cboShema.ListIndex = 0
'    If strReal = "T" Then
'        cboShema.Text = "TE_EOS_TBOWN"
'    ElseIf strReal = "I" Then
'        cboShema.Text = "ODS_TBOWN"
'        cboShema.AddItem ("OAS_TBOWN")
'        cboShema.AddItem ("DWS_TBOWN")
'        cboShema.AddItem ("DMS_TBOWN")
'    Else
'        cboShema.Text = "EOS_TBOWN"
'    End If

    
    
    '공통코드 전체 호출
    'Call getCommonCode
    
    If GCONN_USE = True Then
        If Not fConn Then
            gfMsg ("DB 접속오류")
        End If
    End If
    
    
    cboShema.Text = GetSetting("ORAIB_TAB_MAN_BY_HAK", "OWNER_LIST", "OWNER", "LN")
    
    
    'WebBrowser1.Navigate ("http://iworld.ibksystem.co.kr/iFrame/out/main.xml")
    'WebBrowser1.Navigate ("http://portal.ibk.co.kr")
    'WebBrowser2.Navigate ("http://134.100.207.40/fine_db/EDW/table_group.asp")  'EDW
    'WebBrowser3.Navigate ("http://db2.ibk.co.kr/db/index.htm") '차세대
    
    'sp list setup
    'Call subSPlistSetup
    
    Call subLicenceChek
    
    frameCode.Visible = False
    frameSrc.Visible = False
    
    SSTab1.Tab = 1
    SSTab2.Tab = 0
    SSTab3.Tab = 0
    SSTab4.Tab = 1
    SSTab5.Tab = 0
    SSTab8.Tab = 0
    SSTab9.Tab = 0
    
    
    Me.Left = 200
    Me.Top = 0
    
            
            
    Dim tempDay As String
    tempDay = DateAdd("m", -1, Now)
    
    txtStDay.Text = Mid(tempDay, 1, 4) & Mid(tempDay, 6, 2) & Mid(tempDay, 9, 2)
    txtEdDay.Text = Mid(Now, 1, 4) & Mid(Now, 6, 2) & Mid(Now, 9, 2)
        
        
        
    'owner list 설정
    Call setOwnerList


        
        
        
    '레지스트리에서 등록된 시스템 정보를 읽어와 콤보박스에 셋팅한다.
    Dim i As Integer
    Dim iMaxListCount As Integer
    Dim sTemp As String
    sTemp = cboSystemList.Text  '임시저장했다가 다시 설정
    iMaxListCount = GetSetting("ORAIB_TAB_MAN_BY_HAK", "SYSTEM_LIST", "MAX_LIST_COUNT", "0")
    
    Dim sSystemList As String

    For i = 1 To CInt(iMaxListCount)
    
        sSystemList = GetSetting("ORAIB_TAB_MAN_BY_HAK", "SYSTEM_LIST", "LIST_" & i, "")
        cboSystemList.AddItem (sSystemList)
        
    Next i
    
    cboSystemList.Text = sTemp
    


    
    '즐겨찾기 목록 불러오기
    Call getFavorite

    
    cboShema.ListIndex = 0
    
    
    
End Sub



Private Sub getFavorite()
    
    Dim i As Integer
    Dim iCnt As Integer
    
    
    iCnt = GetSetting("ORAIB_TAB_MAN_BY_HAK", "TAB_FAVORITE", "TOTAL_COUNT", "0")
    
    fpFavorite.MaxRows = iCnt
    For i = 1 To iCnt
        fpFavorite.Row = i
        fpFavorite.Col = 1
        fpFavorite.Value = GetSetting("ORAIB_TAB_MAN_BY_HAK", "TAB_FAVORITE", "TAB_ID_" & i, "")
        
        fpFavorite.Col = 2
        fpFavorite.Value = GetSetting("ORAIB_TAB_MAN_BY_HAK", "TAB_FAVORITE", "TAB_NAME_" & i, "")
        
        fpFavorite.Col = 5
        fpFavorite.Value = GetSetting("ORAIB_TAB_MAN_BY_HAK", "TAB_FAVORITE", "OWNER_NAME_" & i, "")
        
        
    Next i
End Sub

Private Sub setFavorite()
    Dim i As Integer
        
    Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", "TAB_FAVORITE", "TOTAL_COUNT", fpFavorite.MaxRows)
        
    For i = 1 To fpFavorite.MaxRows
        fpFavorite.Row = i
        fpFavorite.Col = 1
        Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", "TAB_FAVORITE", "TAB_ID_" & i, fpFavorite.Value)
        
        fpFavorite.Col = 2
        Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", "TAB_FAVORITE", "TAB_NAME_" & i, fpFavorite.Value)
    
        fpFavorite.Col = 5
        Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", "TAB_FAVORITE", "OWNER_NAME_" & i, fpFavorite.Value)
    
    
    Next i

End Sub


Private Sub subLicenceChek()



On Error GoTo ErrorHandler
    
    Dim conn As ADODB.Connection

    
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim sSQL As String
    Screen.MousePointer = vbHourglass
    

    If GCONN_USE = False Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = GetConnString
    
        conn.Open
    End If
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient




    sSQL = ""
    sSQL = sSQL & vbCrLf & "  SELECT TO_CHAR(SYSDATE,'YYYYMMDD')  FROM DUAL "
    
        

    txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If
    
    
        
'    If Trim(rsTemp.Fields(0).Value) > "20130631" Then
'
'
'
'        MsgBox "#@^%YOU#!$&^*(*& $#@^@$*U^&%^$J~CAN4321~~QWE%@#!Y%GFDDGFVJH$T^$@#%%" & vbCrLf & _
'               "#@G67NOT583493-23904943 %!THIS2342#@$QGSDAdsag" & vbCrLf & _
'               "#@%PROGRAM$#@!%$GgSod#$@#^$#^t3485312BYE#@!$RF352SORRY3f%@#%RDAFas ", vbCritical
'        End
'    End If
    
    
    Screen.MousePointer = vbDefault

    Set rsTemp = Nothing
    
    
    Exit Sub
ErrorHandler:

    Set rsTemp = Nothing
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description
'

End Sub


'------------------------------------------------------------
'sp list setup
'------------------------------------------------------------
Private Sub subSPlistSetup()
    

On Error GoTo ErrorHandler
    
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim sSQL As String
    Dim conn As ADODB.Connection
    Screen.MousePointer = vbHourglass
    

    If GCONN_USE = False Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = GetConnString
    
        conn.Open
    End If
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient




    sSQL = ""
    sSQL = sSQL & vbCrLf & "  SELECT  OBJECT_NAME                       "
    sSQL = sSQL & vbCrLf & "    FROM USER_PROCEDURES" & IIf(strReal = "A", "@LK_STG3", "") & "                 "
    sSQL = sSQL & vbCrLf & "   WHERE 1=1   "
    sSQL = sSQL & vbCrLf & "     AND  OBJECT_NAME LIKE 'SP_CRS%'   "
    sSQL = sSQL & vbCrLf & "   order by OBJECT_NAME "
    
    
    Set rsTemp = gfGlobalSQL(sSQL)
    
    rsTemp.MoveFirst
    
    cmbPgmIDLog.Clear
    For i = 1 To rsTemp.RecordCount
    
        cmbPgmIDLog.AddItem (Trim(rsTemp.Fields(0).Value))
        cmbPgmIDLog.ListIndex = 0
            
        rsTemp.MoveNext
    Next i
    
    
    Screen.MousePointer = vbDefault

    Set rsTemp = Nothing
    
    
    Exit Sub
ErrorHandler:

    Set rsTemp = Nothing
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description
'


End Sub




'golbal connection
Private Function fConn() As Boolean
On Error GoTo ErrorHandler
    
    Set gConn = New ADODB.Connection
    gConn.ConnectionString = GetConnString
    gConn.Open
            
    
    fConn = True
    
    Exit Function
ErrorHandler:
    gfMsg (Err.Description)
    
    fConn = False
    

End Function

'Form Unload
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub fpcode_Click(ByVal Col As Long, ByVal Row As Long)

    If Row < 1 Then Exit Sub
    
    txtCodeWhere(0).Text = ""
    fpcode.Row = Row
    fpcode.Col = 3: txtCodeWhere(0).Text = txtCodeWhere(0).Text & "   AND " & fpcode.Value & " = '"
    fpcode.Col = 4: txtCodeWhere(0).Text = txtCodeWhere(0).Text & fpcode.Value & "'"
    fpcode.Col = 2: txtCodeWhere(0).Text = txtCodeWhere(0).Text & "           -- " & fpcode.Value
    fpcode.Col = 5: txtCodeWhere(0).Text = txtCodeWhere(0).Text & " : " & fpcode.Value & " "
    
    txtCodeWhere(1).Text = ""
    fpcode.Row = Row
    fpcode.Col = 3: txtCodeWhere(1).Text = txtCodeWhere(1).Text & "   AND " & fpcode.Value & " != '"
    fpcode.Col = 4: txtCodeWhere(1).Text = txtCodeWhere(1).Text & fpcode.Value & "'"
    fpcode.Col = 2: txtCodeWhere(1).Text = txtCodeWhere(1).Text & "          -- " & fpcode.Value
    fpcode.Col = 5: txtCodeWhere(1).Text = txtCodeWhere(1).Text & " : " & fpcode.Value & " "
    
    
    txtCodeWhere(2).Text = ""
    fpcode.Row = Row
    fpcode.Col = 3: txtCodeWhere(2).Text = txtCodeWhere(2).Text & "   AND " & fpcode.Value & " IN ('"
    fpcode.Col = 4: txtCodeWhere(2).Text = txtCodeWhere(2).Text & fpcode.Value & "')"
    fpcode.Col = 2: txtCodeWhere(2).Text = txtCodeWhere(2).Text & "          -- " & fpcode.Value
    fpcode.Col = 5: txtCodeWhere(2).Text = txtCodeWhere(2).Text & " : " & fpcode.Value & " "
    
    
    
End Sub

Private Sub fpcode_DblClick(ByVal Col As Long, ByVal Row As Long)
    'Image1_Click
End Sub

Private Sub fpCol_Click(ByVal Col As Long, ByVal Row As Long)
    
    fpCol.Row = Row
    fpCol.Col = Col
    
    
    Dim sCode As String
    
    
 Dim slen() As String
    
    
    
    txtMapping.Text = "KEXIM1" & "|" & _
                      cboShema.Text & "|" & _
                      edtTable.Text & "|" & _
                      edtTablenm.Text & "|"
    fpCol.Col = 1
    txtMapping.Text = txtMapping.Text & fpCol.Value & "|"
    
    fpCol.Col = 2
    txtMapping.Text = txtMapping.Text & fpCol.Value & "|"
    
    fpCol.Col = 8 '컬럼순서
    txtMapping.Text = txtMapping.Text & fpCol.Value & "|"
    
    fpCol.Col = 3 '데이터타입
    txtMapping.Text = txtMapping.Text & fpCol.Value & "|"
                          
    If fpCol.Value = "NUMBER" Then
        fpCol.Col = 4 'data lentgh
        slen = Split(fpCol.Value, ",")
        txtMapping.Text = txtMapping.Text & slen(0) & "|"
        txtMapping.Text = txtMapping.Text & slen(1) & "|"
        
    Else
    
        fpCol.Col = 4 'data lentgh
        txtMapping.Text = txtMapping.Text & fpCol.Value & "|"
        txtMapping.Text = txtMapping.Text & "" & "|"
                      
    End If
                      
    fpCol.Col = 5
    
    If Trim(fpCol.Value) <> "" Then
        txtMapping.Text = txtMapping.Text & "Y" & "|"
        
    Else
        txtMapping.Text = txtMapping.Text & "N" & "|"
    End If
    
    txtMapping.Text = txtMapping.Text & fpCol.Value & "|"
    
    
    fpCol.Col = 11
    txtMapping.Text = txtMapping.Text & fpCol.Value ' IIf(fpCol.Value = "Y", "N", "Y")
    
    
    
    fpCol.Col = Col
    
    
    If Col <> 6 Then Exit Sub
    
   
    If Trim(fpCol.Value) = "" Then
    
    
        Exit Sub
        
    End If
    
    'RSM 공통코드 조회
    fpCol.Col = 2
    If UCase(Left(fpCol.Value, 2)) = "CD" Then
        sCode = "코드"
    Else
        sCode = fpCol.Value
    End If
    
    
    fpCol.Col = 1
    txtcolname.Text = fpCol.Value
    
    subSelectCodeList (sCode)
    
    txtCodeWhere(0).Text = ""
    txtCodeWhere(1).Text = ""
    txtCodeWhere(2).Text = ""
    
    
    
    txtCodeWhere(3).Text = ""
    fpcode.Row = 1
    
    txtCodeWhere(3).Text = ", SF_CO_SELCDNM('"
    fpcode.Col = 1: txtCodeWhere(3).Text = txtCodeWhere(3).Text & fpcode.Value & "', "
    fpcode.Col = 1: txtCodeWhere(3).Text = txtCodeWhere(3).Text & txtcolname.Text & ")  AS " & txtcolname.Text & "_NM"
    fpcode.Col = 2: txtCodeWhere(3).Text = txtCodeWhere(3).Text & "  --" & fpcode.Value & "명"
    
    
    
    
    frameCode.Visible = True
    txtCodeSrch.SetFocus
   
    
                   
    
End Sub

'**********************************************************
'컬럼 정보 더블 클릭 이벤트 처리
'**********************************************************

Private Sub fpCol_DblClick(ByVal Col As Long, ByVal Row As Long)

    

On Error GoTo ErrorHandler
            
    Dim sCode As String
            
    If Col = 6 Then
        fpCol.Row = Row
        fpCol.Col = Col
    
    
        '코드가 공백인데 더블클릭이면 등록 할 수 있게 띄운다.
        If Trim(fpCol.Value) = "" Then
            fpcode.MaxRows = 0
            fpcode.MaxRows = 20
            
            fpcode.Row = 1
            fpcode.Col = 1
            'fpcode.Value = "XX" & Format(Now, "MMDDHHMMSS")
            
            fpCol.Col = 2
            
            fpcode.Col = 2
            fpcode.Value = fpCol.Value
            
            frameCode.Visible = True
            
                    
            cmdSave.Visible = True
            cmdADD.Visible = True
            cmdDel.Visible = True
            
            
    
        End If
        
    End If
    
    
    
    
    
            
            
    '----------------------
            
            
    '기본정보 세팅
    With fpCol
        .Row = Row
        .Col = 1: txtColId.Text = .Value
        .Col = 2: txtColNm.Text = .Value
        .Col = 6: txtCodeDomain.Text = .Value
        .Col = 7: txtColBigo.Text = .Value
        
    End With
        

    
    Exit Sub
ErrorHandler:
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description
    
    
'



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


  
    fpCol.Row = fpCol.ActiveRow
    fpCol.Col = 1
    Dim sEng_Nm As String
    sEng_Nm = fpCol.Value


    'AS-IS인 경우 코드 조회 sql
    If strReal = "A" Then
    
        sSQL = sSQL & vbCrLf & "SELECT A.CD_GRP_ID"
        sSQL = sSQL & vbCrLf & "      ,A.CD_GRP_NM"
        sSQL = sSQL & vbCrLf & "      ,''   as eng_nm"
        sSQL = sSQL & vbCrLf & "      ,B.CD_ID"
        sSQL = sSQL & vbCrLf & "      ,B.CD_NM"
        sSQL = sSQL & vbCrLf & "      ,''       AS USE_YN"
        sSQL = sSQL & vbCrLf & "      ,''       AS REMARKS"
        sSQL = sSQL & vbCrLf & "  FROM TB_ZZ1001_CDGRP A"
        sSQL = sSQL & vbCrLf & "      ,TB_ZZ1002_CD    B"
        sSQL = sSQL & vbCrLf & " WHERE 1=1"
        sSQL = sSQL & vbCrLf & "   AND A.CD_GRP_ID = B.CD_GRP_ID"
        
        '컬럼 정보에서 호출되었으면 해당 코드로
        If Trim(sCode) <> "" Then
        
        '아니면 검색 조건으로 인한 검색
        Else
            sSQL = sSQL & vbCrLf & "   AND (A.CD_GRP_NM LIKE '%" & Trim(txtCodeSrchNm.Text) & "%'"
            sSQL = sSQL & vbCrLf & "        OR A.CD_GRP_ID LIKE '%" & Trim(txtCodeSrchNm.Text) & "%' "
            sSQL = sSQL & vbCrLf & "       )"
        End If
        
        sSQL = sSQL & vbCrLf & "ORDER BY A.CD_GRP_ID, B.CD_ID    "
    
    
    
    Else
    
'            sSQL = ""
'            sSQL = sSQL & vbCrLf & "  SELECT  CD_DMN_NO                       "
'            sSQL = sSQL & vbCrLf & "         ,CD_DMN_KRN_NM                   "
'            sSQL = sSQL & vbCrLf & "         , '" & sEng_Nm & "'"
'            sSQL = sSQL & vbCrLf & "         ,CD_INST_VL                        "
'            sSQL = sSQL & vbCrLf & "         ,CD_INST_NM                        "
'            sSQL = sSQL & vbCrLf & "         ,'' USE_YN                            "
'            sSQL = sSQL & vbCrLf & "         , REMARK                              "
'
'
'
'        sSQL = sSQL & vbCrLf & "         FROM (    SELECT B.C_CLAS_NO   AS CD_DMN_NO"
'        sSQL = sSQL & vbCrLf & "                         ,B.C_CLAS_NM   AS CD_DMN_KRN_NM"
'        sSQL = sSQL & vbCrLf & "                         ,A.C_NO        AS CD_INST_VL"
'        sSQL = sSQL & vbCrLf & "                         ,A.STND_C_NM   AS CD_INST_NM"
'        sSQL = sSQL & vbCrLf & "                         ,''            AS REMARK"
'        sSQL = sSQL & vbCrLf & "                     FROM CO.TBGM_COC_TBL  A"
'        sSQL = sSQL & vbCrLf & "                        ,CO.TBGM_COC_CLAS B"
'        sSQL = sSQL & vbCrLf & "                     Where 1 = 1"
'        sSQL = sSQL & vbCrLf & "                       AND A.C_CLAS_NO = B.C_CLAS_NO  UNION ALL"
'        sSQL = sSQL & vbCrLf & "                   SELECT CODDOM            "
'        sSQL = sSQL & vbCrLf & "                         ,COMDOMNM          "
'        sSQL = sSQL & vbCrLf & "                         ,CDVAL             "
'        sSQL = sSQL & vbCrLf & "                         ,CDVALNM           "
'        sSQL = sSQL & vbCrLf & "                         ,BIGO              "
'        sSQL = sSQL & vbCrLf & "                   From LN.TB_DIC_COD          "
'        sSQL = sSQL & vbCrLf & "            )"
           
            sSQL = ""
            sSQL = sSQL & vbCrLf & "SELECT A.CD_GRP_ID"
            sSQL = sSQL & vbCrLf & "      ,A.CD_GRP_NM"
            sSQL = sSQL & vbCrLf & "      ,''   as eng_nm"
            sSQL = sSQL & vbCrLf & "      ,B.CD_ID"
            sSQL = sSQL & vbCrLf & "      ,B.CD_NM"
            sSQL = sSQL & vbCrLf & "      ,''       AS USE_YN"
            sSQL = sSQL & vbCrLf & "      ,''       AS REMARKS"
            sSQL = sSQL & vbCrLf & "  FROM USCO.TB_COA310M_CDGRPBSC A"
            sSQL = sSQL & vbCrLf & "      ,USCO.TB_COA311M_CDBSC B"
            sSQL = sSQL & vbCrLf & " WHERE 1=1"
            sSQL = sSQL & vbCrLf & "   AND A.CD_GRP_ID = B.CD_GRP_ID"
            
            
            
        
        

                    
            
            If sCode = "코드" Then
                
                fpCol.Row = fpCol.ActiveRow
                fpCol.Col = 2
                sSQL = sSQL & vbCrLf & "     AND TRIM(CD_GRP_NM) = '" & fpCol.Value & "'              "
            Else
                sSQL = sSQL & vbCrLf & "     AND upper(CD_GRP_NM) = '" & UCase(sCode) & "'              "
            End If
            sSQL = sSQL & vbCrLf & "ORDER BY A.CD_GRP_ID, B.CD_ID    "
            
    End If
    
    
    txtSQL.Text = sSQL
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

    
    
    
    fpcode.Row = 1
    fpcode.Col = 1
    
    

    
    
    
    
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

Private Sub fpDebugLogHist_DblClick(ByVal Col As Long, ByVal Row As Long)
  Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
   
    
    If GCONN_USE = False Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = GetConnString
    
        conn.Open
    End If
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
    If Trim(cmbPgmIDLog.Text) = "" Then
        MsgBox "Input Program Id "
        Exit Sub
    End If

    Dim sPgm As String
    Dim sSeq As String
    fpDebugLogHist.Row = Row
    fpDebugLogHist.Col = 1
    sPgm = fpDebugLogHist.Value
    
    fpDebugLogHist.Row = Row
    fpDebugLogHist.Col = 2
    sSeq = fpDebugLogHist.Value


    Screen.MousePointer = vbHourglass

    sSQL = ""
    sSQL = sSQL & vbCrLf & "SELECT                                                                                                     "
    sSQL = sSQL & vbCrLf & "       LAG(PRRM_ID,1,PRRM_ID) OVER (ORDER BY LOG_SRNO)  AS PRRM_ID                          "
    sSQL = sSQL & vbCrLf & "      ,MNNO                                                                                                "
    sSQL = sSQL & vbCrLf & "      ,CASE WHEN LOG_SRNO = 99999 THEN LAG(LOG_SRNO,1,LOG_SRNO) OVER (ORDER BY LOG_SRNO)  + 1             "
    sSQL = sSQL & vbCrLf & "             ELSE LOG_SRNO                                                                                 "
    sSQL = sSQL & vbCrLf & "       END  AS  LOG_SRNO                                                                                  "
    sSQL = sSQL & vbCrLf & "      ,BULT_CNTN                                                                                          "
    sSQL = sSQL & vbCrLf & "      ,AF_TIME                                                                                            "
    sSQL = sSQL & vbCrLf & "      ,TO_CHAR(FLOOR(AF_TIME - BF_TIME))  ||  '일 ' ||                                                     "
    sSQL = sSQL & vbCrLf & "       TO_CHAR(                                                                                            "
    sSQL = sSQL & vbCrLf & "                FLOOR( (MOD(AF_TIME - BF_TIME, 1) * 24) )                                                  "
    sSQL = sSQL & vbCrLf & "              , '00') || ':' ||                                                                            "
    sSQL = sSQL & vbCrLf & "       TO_CHAR(                                                                                            "
    sSQL = sSQL & vbCrLf & "                   FLOOR( (MOD( (MOD(AF_TIME - BF_TIME, 1                                                  "
    sSQL = sSQL & vbCrLf & "                                    ) * 24                                                                 "
    sSQL = sSQL & vbCrLf & "                                 ), 1                                                                      "
    sSQL = sSQL & vbCrLf & "                              ) * 60                                                                       "
    sSQL = sSQL & vbCrLf & "                          )                                                                                "
    sSQL = sSQL & vbCrLf & "                        )                                                                                  "
    sSQL = sSQL & vbCrLf & "             , '00' ) || ':' ||                                                                            "
    sSQL = sSQL & vbCrLf & "       TO_CHAR(                                                                                            "
    sSQL = sSQL & vbCrLf & "                   FLOOR( MOD (                                                                            "
    sSQL = sSQL & vbCrLf & "                                   (MOD( (MOD(AF_TIME - BF_TIME, 1                                         "
    sSQL = sSQL & vbCrLf & "                                            ) * 24                                                         "
    sSQL = sSQL & vbCrLf & "                                         ), 1                                                              "
    sSQL = sSQL & vbCrLf & "                                      ) * 60                                                               "
    sSQL = sSQL & vbCrLf & "                                  )                                                                        "
    sSQL = sSQL & vbCrLf & "                                , 1                                                                        "
    sSQL = sSQL & vbCrLf & "                              ) * 60                                                                       "
    sSQL = sSQL & vbCrLf & "                        )                                                                                  "
    sSQL = sSQL & vbCrLf & "              , '00' )                                                                                     "
    sSQL = sSQL & vbCrLf & "  FROM                                                                                                     "
    sSQL = sSQL & vbCrLf & "        (                                                                                                  "
    sSQL = sSQL & vbCrLf & "            SELECT  PRRM_ID                                                                           "
    sSQL = sSQL & vbCrLf & "                   ,LOG_SRNO                                                                               "
    sSQL = sSQL & vbCrLf & "                   ,LAG(LOG_TRYMD,1, LOG_TRYMD) OVER (ORDER BY LOG_SRNO)  AS BF_TIME                       "
    sSQL = sSQL & vbCrLf & "                   ,LOG_TRYMD                                  AS AF_TIME                                  "
    sSQL = sSQL & vbCrLf & "                   ,BULT_CNTN ,  MNNO                                                                           "
    sSQL = sSQL & vbCrLf & "              FROM (                                                                                       "
    sSQL = sSQL & vbCrLf & "                    SELECT  PRRM_ID                                                                   "
    sSQL = sSQL & vbCrLf & "                           ,LOG_SRNO                                                                       "
    sSQL = sSQL & vbCrLf & "                           ,LOG_TRYMD                                                                      "
    sSQL = sSQL & vbCrLf & "                           ,BULT_CNTN       ,  MNNO                                                                "
    sSQL = sSQL & vbCrLf & "                      FROM TB_CRS_CM013_L_A                                                                "
    sSQL = sSQL & vbCrLf & "                     WHERE PRRM_ID = UPPER('" & cmbPgmIDLog.Text & "')                                "
    sSQL = sSQL & vbCrLf & "                       AND MNNO = '" & sSeq & "'                                                      "
    sSQL = sSQL & vbCrLf & "                    UNION ALL                                                                              "
    sSQL = sSQL & vbCrLf & "                    SELECT  'ZZZZZZ'                                                                       "
    sSQL = sSQL & vbCrLf & "                           ,99999                                                                          "
    sSQL = sSQL & vbCrLf & "                           ,SYSDATE                                                                        "
    sSQL = sSQL & vbCrLf & "                           ,'현재 시각'    , ' '                                                                "
    sSQL = sSQL & vbCrLf & "                      FROM DUAL                                                                            "
    sSQL = sSQL & vbCrLf & "                    ) X                                                                                    "
    sSQL = sSQL & vbCrLf & "         )                                                                                                 "
    sSQL = sSQL & vbCrLf & "   ORDER BY   LOG_SRNO                                                                                     "
    
    

    txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If

    fpDebugLog.MaxRows = 0
    If rsTemp.RecordCount < 1 Then
        fpDebugLog.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpDebugLog, rsTemp)
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



Private Sub getCommonCode()
    Dim conn As ADODB.Connection
    Dim sSQL As String
    

On Error GoTo ErrorHandler
     

    Screen.MousePointer = vbHourglass

'
'    '테이블상태 정보 조회
'     Call subTableStat
'
    If GCONN_USE = False Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = GetConnString
        conn.Open
    End If

    Set rscode = New ADODB.Recordset
    rscode.CursorLocation = adUseClient




    sSQL = ""
    sSQL = sSQL & vbCrLf & "SELECT  SMPL_LCCD                      -- 단순대분류코드"
    sSQL = sSQL & vbCrLf & "   ,SMPL_MCCD                      -- 단순중분류코드"
    sSQL = sSQL & vbCrLf & "   ,SMPL_SCCD                      -- 단순소분류코드"
    sSQL = sSQL & vbCrLf & "   ,SMPL_LCCD_NM                   -- 단순대분류코드명"
    sSQL = sSQL & vbCrLf & "   ,SMPL_MCCD_NM                   -- 단순중분류코드명"
    sSQL = sSQL & vbCrLf & "   ,SMPL_SCCD_NM                   -- 단순소분류코드명"
    sSQL = sSQL & vbCrLf & "FROM TCMSMPLCODEC        /* 공통단순코드*/"


    txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rscode.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rscode.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If

        
    Screen.MousePointer = vbDefault

    Set conn = Nothing
    
    
    Exit Sub
ErrorHandler:
    Set conn = Nothing
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description

End Sub

Private Sub fpFavorite_DblClick(ByVal Col As Long, ByVal Row As Long)

    Call gfGrid_Sort(fpFavorite, Col, Row)
                        
    If Row < 1 Then Exit Sub

    fpTable.MaxRows = fpTable.MaxRows + 1
    fpTable.Row = fpTable.MaxRows
    
    fpFavorite.Row = Row
    fpFavorite.Col = 1
    
    fpTable.Col = 1
    fpTable.Value = fpFavorite.Value
    
    fpFavorite.Col = 2
    fpTable.Col = 2
    fpTable.Value = fpFavorite.Value
    
    fpFavorite.Col = 5
    fpTable.Col = 5
    fpTable.Value = fpFavorite.Value
    
    
    
    Call fpTable_DblClick(1, fpTable.MaxRows)
    
    fpTable.MaxRows = fpTable.MaxRows - 1
End Sub

Private Sub fpNear_DblClick(ByVal Col As Long, ByVal Row As Long)

    Call gfGrid_Sort(fpNear, Col, Row)
                        
    If Row < 1 Then Exit Sub

    fpTable.MaxRows = fpTable.MaxRows + 1
    fpTable.Row = fpTable.MaxRows
    
    fpNear.Row = Row
    fpNear.Col = 1
    
    fpTable.Col = 1
    fpTable.Value = fpNear.Value
    
    fpNear.Col = 2
    fpTable.Col = 2
    fpTable.Value = fpNear.Value
    
    fpNear.Col = 5
    fpTable.Col = 5
    fpTable.Value = fpNear.Value
    
    
    
    Call fpTable_DblClick(1, fpTable.MaxRows)
    
    fpTable.MaxRows = fpTable.MaxRows - 1
End Sub

Private Sub fpSample_Change(ByVal Col As Long, ByVal Row As Long)
    With fpSample
    
        
        .Row = Row
        
        .Col = 1
        If .BackColor = C_ADD_COLOR Or .BackColor = S_DARK_RED Then Exit Sub
        
        .Col = -1
        .BackColor = S_DARK_CYAN
    End With
End Sub

Private Sub fpSample_Click(ByVal Col As Long, ByVal Row As Long)
    
     If C_SPECIAL <> "N" Then
        Exit Sub
    End If
    fpSample.Col = Col
    fpSample.Row = Row
    
    
    'rscode.Filter = adFilterNone
    'rscode.Filter = "SMPL_SCCD='" & fpSample.Value & "'"
    
On Error Resume Next
    'lblHint.Caption = rscode.Fields(5)
    
End Sub

Private Sub fpSP_DblClick(ByVal Col As Long, ByVal Row As Long)

On Error GoTo ErrorHandler
    
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim sSQL As String
    Screen.MousePointer = vbHourglass
    

   'frmSrc.Visible = True

    fpSP.Row = Row
    fpSP.Col = Col

    sSQL = ""
    sSQL = sSQL & vbCrLf & "  select REPLACE(REPLACE(TEXT, CHR(9), '') , CHR(13), '') AS TEXT                       "
    sSQL = sSQL & vbCrLf & "   from user_SOURCE "
    sSQL = sSQL & vbCrLf & "   WHERE 1=1  "
    sSQL = sSQL & vbCrLf & "      AND NAME = '" & fpSP.Value & "'  "
    sSQL = sSQL & vbCrLf & "   ORDER BY LINE  "
    
    
    Set rsTemp = gfGlobalSQL(sSQL)
    
    rsTemp.MoveFirst
    
    txtSrc.Text = ""
    Dim sTemp As String
    
    txtSRC2.Text = ""
    For i = 1 To rsTemp.RecordCount
        
        sTemp = rsTemp.Fields(0).Value
        sTemp = Replace(sTemp, Chr(10), "")
        sTemp = Replace(sTemp, Chr(13), "")
        
        txtSRC2.Text = txtSRC2.Text & sTemp & vbCrLf
        rsTemp.MoveNext
    Next i
   
        
    
    
    Screen.MousePointer = vbDefault

    Set rsTemp = Nothing
    
    
    Exit Sub
ErrorHandler:

    Set rsTemp = Nothing
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description
End Sub

'첫글자만 대문자로
Public Function gfUpperFirst(sInStr) As String

    Dim i As Integer
    Dim sReturn As String
    
    For i = 1 To Len(sInStr)
        
        If i = 1 Then
            sReturn = sReturn & UCase(Mid(sInStr, i, 1))
        Else
            sReturn = sReturn & LCase(Mid(sInStr, i, 1))
        End If
    
    Next i
    
    gfUpperFirst = sReturn

End Function
'---------------------------------------------------------------------------------------
'table grid doublce click event
'---------------------------------------------------------------------------------------
Private Sub fpTable_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
                    
    Call gfGrid_Sort(fpTable, Col, Row)
                        
    If Row < 1 Then Exit Sub
                            
                            


    '기본정보 세팅
    With fpTable
        .Row = Row
        .Col = 1: edtTable.Text = .Value
        .Col = 2: edtTablenm.Text = .Value
        .Col = 3: edtTableBigo.Text = .Value
        .Col = 4: edtTableCrdate.Text = .Value
        .Col = 5: txtOwner.Text = .Value
        
    End With
    
    
    '마이그 정보 셋팅
    txtMigOwner.Text = "USHG"
'txtMigOwner.Text = txtOwner.Text
    txtMigTableID.Text = edtTable.Text
    txtMigTableNm.Text = edtTablenm.Text
    txtMigCnt.Text = ""
    
    
    txtTgtOwner.Text = txtOwner.Text
    txtTgtTableId.Text = edtTable.Text
    txtTgtTableNm.Text = edtTablenm.Text
    txtTgtCnt.Text = ""
    
    txtMigSQL.Text = ""
    
    
    
    
    'Dao명 생성
    Dim sArr() As String
    Dim i As Integer
    
    sArr = Split(edtTable.Text, "_")
    txtTableDao.Text = ""
    For i = LBound(sArr) To UBound(sArr)
        txtTableDao.Text = txtTableDao.Text & gfUpperFirst(sArr(i))
    Next i
    txtTableDao.Text = txtTableDao.Text & "001Dao"
    
    
    txtTabMapping.Text = "KEXIM1" & "|" & _
                          cboShema.Text & "|" & _
                          edtTable.Text & "|" & _
                          edtTablenm.Text

    

    '컬럼정보CLEAR
    txtColId.Text = ""
    txtColNm.Text = ""
    txtCodeDomain.Text = ""
    txtColBigo.Text = ""
'
'    ftPartition.MaxRows = 0

    Screen.MousePointer = vbHourglass

'
'    '테이블상태 정보 조회
'     Call subTableStat
'
    If GCONN_USE = False Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = GetConnString
        conn.Open
    End If

    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
'
'
'If strReal <> "T" And strReal <> "I" Then
'
'    sSQL = ""
'    sSQL = sSQL & vbCrLf & "SELECT  COLUMN_NAME "
'    sSQL = sSQL & vbCrLf & "       ,COLUMN_COMM "
'    sSQL = sSQL & vbCrLf & "       ,DATA_TYPE "
'    sSQL = sSQL & vbCrLf & "       ,DATA_LENGTH "
'    sSQL = sSQL & vbCrLf & "       ,POSITION "
'    sSQL = sSQL & vbCrLf & "       ,CODE_DOM "
''    sSQL = sSQL & vbCrLf & "       ,CASE WHEN CODE_DOM IS NULL OR  TRIM(CODE_DOM) = '' THEN                                   "
''    sSQL = sSQL & vbCrLf & "                  CASE WHEN COLUMN_COMM LIKE '%코드%' AND COLUMN_COMM NOT LIKE '%부점코드%' THEN                                         "
''    sSQL = sSQL & vbCrLf & "                                                       '메타'                            "
''    sSQL = sSQL & vbCrLf & "                END                                                                          "
''    sSQL = sSQL & vbCrLf & "              ELSE CODE_DOM                                                                       "
''    sSQL = sSQL & vbCrLf & "        END CODE_DOM                                                                                "
'    sSQL = sSQL & vbCrLf & "       ,BIGO1 "
'    sSQL = sSQL & vbCrLf & "       ,COLUMN_ID"
'    sSQL = sSQL & vbCrLf & "       ,DATA_VALUE "
'    sSQL = sSQL & vbCrLf & "       ,BIGO2 "
'    sSQL = sSQL & vbCrLf & "  FROM ( "
'            sSQL = sSQL & vbCrLf & "SELECT  A.COLUMN_NAME                                                                       "
'            sSQL = sSQL & vbCrLf & "       ,CASE WHEN A.COLUMN_COMM IS NULL THEN A.DIC_COMM                                     "
'            sSQL = sSQL & vbCrLf & "             ELSE A.COLUMN_COMM                                                             "
'            sSQL = sSQL & vbCrLf & "        END COLUMN_COMM                                                                     "
'            sSQL = sSQL & vbCrLf & "       ,A.DATA_TYPE                                                                         "
'            sSQL = sSQL & vbCrLf & "       ,CASE A.DATA_TYPE WHEN 'NUMBER' THEN A.DATA_PRECISION || ',' || A.DATA_SCALE         "
'            sSQL = sSQL & vbCrLf & "                         ELSE TO_CHAR(A.DATA_LENGTH)                                        "
'            sSQL = sSQL & vbCrLf & "        END  DATA_LENGTH                                                                    "
'            sSQL = sSQL & vbCrLf & "       ,B.POSITION                                                                          "
'            sSQL = sSQL & vbCrLf & "       ,A.CODE_DOM                                                                          "
'            sSQL = sSQL & vbCrLf & "       ,A.BIGO1                                                                             "
'            sSQL = sSQL & vbCrLf & "       ,A.COLUMN_ID                                                                         "
'            sSQL = sSQL & vbCrLf & "       ,'' AS DATA_VALUE                                                                    "
'            sSQL = sSQL & vbCrLf & "       ,A.BIGO2                                                                             "
'            sSQL = sSQL & vbCrLf & "  FROM (                                                                                    "
'            sSQL = sSQL & vbCrLf & "        SELECT  A.COLUMN_NAME                                                               "
'            sSQL = sSQL & vbCrLf & "               ,A.DATA_TYPE                                                                 "
'            sSQL = sSQL & vbCrLf & "               ,A.DATA_LENGTH                                                               "
'            sSQL = sSQL & vbCrLf & "               ,A.DATA_PRECISION                                                            "
'            sSQL = sSQL & vbCrLf & "               ,A.DATA_SCALE                                                                "
'            sSQL = sSQL & vbCrLf & "               ,A.NULLABLE                                                                  "
'            sSQL = sSQL & vbCrLf & "               ,A.COLUMN_ID                                                                 "
'            sSQL = sSQL & vbCrLf & "               ,(SELECT COMMENTS                                                            "
'            sSQL = sSQL & vbCrLf & "                   FROM ALL_COL_COMMENTS" & IIf(chkASIS.Value = 1, "@DL_IBKCDEV", "") & " B                                                  "
'            sSQL = sSQL & vbCrLf & "                  WHERE B.OWNER = A.OWNER                                                   "
'            sSQL = sSQL & vbCrLf & "                    AND B.TABLE_NAME = A.TABLE_NAME                                         "
'            sSQL = sSQL & vbCrLf & "                    AND B.COLUMN_NAME = A.COLUMN_NAME                                       "
'            sSQL = sSQL & vbCrLf & "                )  DIC_COMM                                                                 "
'            sSQL = sSQL & vbCrLf & "                ,C.COLUMN_COMM                                                              "
'            sSQL = sSQL & vbCrLf & "                ,C.CODE_DOM                                                                 "
'            sSQL = sSQL & vbCrLf & "                ,C.BIGO1                                                                    "
'            sSQL = sSQL & vbCrLf & "                ,C.BIGO2                                                                    "
'
'            'sSQL = sSQL & vbCrLf & "                ,' ' as COLUMN_COMM                                                              "
'            'sSQL = sSQL & vbCrLf & "                ,' ' as CODE_DOM                                                                 "
'            'sSQL = sSQL & vbCrLf & "                ,' ' as BIGO1                                                                    "
'            'sSQL = sSQL & vbCrLf & "                ,' ' as BIGO2                                                                    "
'
'
'            sSQL = sSQL & vbCrLf & "          FROM ALL_TAB_COLUMNS" & IIf(chkASIS.Value = 1, "@DL_IBKCDEV", "") & " A                                                            "
'            sSQL = sSQL & vbCrLf & "                   LEFT OUTER JOIN                                                          "
'            sSQL = sSQL & vbCrLf & "               ln.TB_DIC_COL" & IIf(chkASIS.Value = 1, "@DL_IBKCDEV", IIf(strReal = "R", "@DL_DEOSDB", "")) & "      C                                                            "
'            sSQL = sSQL & vbCrLf & "                   ON A.OWNER       = C.OWNER                                               "
'            sSQL = sSQL & vbCrLf & "                  AND A.TABLE_NAME  = C.TABLE_NAME                                          "
'            sSQL = sSQL & vbCrLf & "                  AND A.COLUMN_NAME = C.COLUMN_NAME                                         "
'            sSQL = sSQL & vbCrLf & "          WHERE A.TABLE_NAME = '" & edtTable.Text & "' "
'            sSQL = sSQL & vbCrLf & "             AND A.OWNER      = '" & IIf(chkASIS.Value = 1, "IBKCDEV", UCase(cboShema.Text)) & "'"
'            sSQL = sSQL & vbCrLf & "        )  A                                                                                "
'            sSQL = sSQL & vbCrLf & "            LEFT OUTER JOIN                                                                 "
'            sSQL = sSQL & vbCrLf & "        (                                                                                   "
'            sSQL = sSQL & vbCrLf & "        SELECT  B.COLUMN_NAME                                                               "
'            sSQL = sSQL & vbCrLf & "               ,B.POSITION                                                                  "
'            sSQL = sSQL & vbCrLf & "          FROM ALL_CONSTRAINTS" & IIf(chkASIS.Value = 1, "@DL_IBKCDEV", "") & "  A,                                                          "
'            sSQL = sSQL & vbCrLf & "               ALL_CONS_COLUMNS" & IIf(chkASIS.Value = 1, "@DL_IBKCDEV", "") & " B                                                           "
'            sSQL = sSQL & vbCrLf & "           WHERE A.TABLE_NAME = '" & edtTable.Text & "' "
'            sSQL = sSQL & vbCrLf & "             AND A.OWNER      = '" & IIf(chkASIS.Value = 1, "IBKCDEV", UCase(cboShema.Text)) & "'"
'            sSQL = sSQL & vbCrLf & "           AND A.CONSTRAINT_TYPE = 'P'                                                      "
'            sSQL = sSQL & vbCrLf & "           AND A.OWNER   = B.OWNER                                                          "
'            sSQL = sSQL & vbCrLf & "           AND A.CONSTRAINT_NAME = B.CONSTRAINT_NAME                                        "
'            sSQL = sSQL & vbCrLf & "           AND A.TABLE_NAME = B.TABLE_NAME                                                  "
'            sSQL = sSQL & vbCrLf & "         )  B                                                                               "
'            sSQL = sSQL & vbCrLf & "           ON A.COLUMN_NAME = B.COLUMN_NAME                                                 "
'    sSQL = sSQL & vbCrLf & "         )"
'    sSQL = sSQL & vbCrLf & " ORDER BY COLUMN_ID                                                                         "
'
'
'Else

   sSQL = ""
    sSQL = sSQL & vbCrLf & "SELECT  COLUMN_NAME "
    sSQL = sSQL & vbCrLf & "       ,COLUMN_COMM "
    sSQL = sSQL & vbCrLf & "       ,DATA_TYPE "
    sSQL = sSQL & vbCrLf & "       ,DATA_LENGTH "
    sSQL = sSQL & vbCrLf & "       ,POSITION "
    sSQL = sSQL & vbCrLf & "       ,CASE WHEN CODE_DOM IS NULL OR  TRIM(CODE_DOM) = '' THEN                                   "
    sSQL = sSQL & vbCrLf & "                  CASE WHEN COLUMN_COMM LIKE '%코드%' AND COLUMN_COMM NOT LIKE '%부점코드%' THEN                                         "
    sSQL = sSQL & vbCrLf & "                                                       '코드'                            "
    sSQL = sSQL & vbCrLf & "                END                                                                          "
    sSQL = sSQL & vbCrLf & "              ELSE CODE_DOM                                                                       "
    sSQL = sSQL & vbCrLf & "        END CODE_DOM                                                                                "

    sSQL = sSQL & vbCrLf & "       ,BIGO1 "
    sSQL = sSQL & vbCrLf & "       ,COLUMN_ID"
    sSQL = sSQL & vbCrLf & "       ,OPCODE"
    sSQL = sSQL & vbCrLf & "       ,DATA_VALUE "
    sSQL = sSQL & vbCrLf & "       ,BIGO2 "
    sSQL = sSQL & vbCrLf & "       ,NULLABLE "
    sSQL = sSQL & vbCrLf & "  FROM ( "
            sSQL = sSQL & vbCrLf & "SELECT  A.COLUMN_NAME                                                                       "
            sSQL = sSQL & vbCrLf & "       ,A.DIC_COMM     as  COLUMN_COMM                                                                     "
            sSQL = sSQL & vbCrLf & "       ,A.DATA_TYPE                                                                         "
            sSQL = sSQL & vbCrLf & "       ,CASE A.DATA_TYPE WHEN 'NUMBER' THEN A.DATA_PRECISION || ',' || A.DATA_SCALE         "
            sSQL = sSQL & vbCrLf & "                         ELSE TO_CHAR(A.DATA_LENGTH)                                        "
            sSQL = sSQL & vbCrLf & "        END  DATA_LENGTH                                                                    "
            sSQL = sSQL & vbCrLf & "       ,B.POSITION                                                                          "
            sSQL = sSQL & vbCrLf & "       ,''           AS CODE_DOM                                                                          "
            sSQL = sSQL & vbCrLf & "       ,''           AS BIGO1                                                                     "
            sSQL = sSQL & vbCrLf & "       ,A.COLUMN_ID  AS COLUMN_ID                                                                       "
            sSQL = sSQL & vbCrLf & "       ,'='          AS OPCODE                                                                       "
            sSQL = sSQL & vbCrLf & "       ,''           AS DATA_VALUE                                                                    "
            sSQL = sSQL & vbCrLf & "       ,''           AS BIGO2                                                                             "
            sSQL = sSQL & vbCrLf & "       , NULLABLE                                                                             "
            sSQL = sSQL & vbCrLf & "  FROM (                                                                                    "
            sSQL = sSQL & vbCrLf & "        SELECT  A.COLUMN_NAME                                                               "
            sSQL = sSQL & vbCrLf & "               ,A.DATA_TYPE                                                                 "
            sSQL = sSQL & vbCrLf & "               ,A.DATA_LENGTH                                                               "
            sSQL = sSQL & vbCrLf & "               ,A.DATA_PRECISION                                                            "
            sSQL = sSQL & vbCrLf & "               ,A.DATA_SCALE                                                                "
            sSQL = sSQL & vbCrLf & "               ,A.NULLABLE                                                                  "
            sSQL = sSQL & vbCrLf & "               ,A.COLUMN_ID                                                                 "
            sSQL = sSQL & vbCrLf & "               ,(SELECT COMMENTS                                                            "
            sSQL = sSQL & vbCrLf & "                   FROM ALL_COL_COMMENTS" & IIf(strReal = "A", "@LK_STG3", "") & " B                                                  "
            sSQL = sSQL & vbCrLf & "                  WHERE B.OWNER = A.OWNER                                                   "
            sSQL = sSQL & vbCrLf & "                    AND B.TABLE_NAME = A.TABLE_NAME                                         "
            sSQL = sSQL & vbCrLf & "                    AND B.COLUMN_NAME = A.COLUMN_NAME                                       "
            sSQL = sSQL & vbCrLf & "                )  DIC_COMM                                                                 "
            sSQL = sSQL & vbCrLf & "          FROM ALL_TAB_COLUMNS" & IIf(strReal = "A", "@LK_STG3", "") & " A                                                            "
            sSQL = sSQL & vbCrLf & "          WHERE A.TABLE_NAME = '" & edtTable.Text & "' "
            If Trim(cboShema.Text) <> "" Then
                sSQL = sSQL & vbCrLf & "             AND A.OWNER      = '" & IIf(chkASIS.Value = 1, "IBKCDEV", UCase(cboShema.Text)) & "'"
            Else
                sSQL = sSQL & vbCrLf & "             AND A.OWNER      NOT IN  " & C_NOTIN
            End If
            sSQL = sSQL & vbCrLf & "             AND A.OWNER      = '" & Trim(txtOwner.Text) & "'"
            sSQL = sSQL & vbCrLf & "        )  A                                                                                "
            sSQL = sSQL & vbCrLf & "            LEFT OUTER JOIN                                                                 "
            sSQL = sSQL & vbCrLf & "        (                                                                                   "
            sSQL = sSQL & vbCrLf & "        SELECT  B.COLUMN_NAME                                                               "
            sSQL = sSQL & vbCrLf & "               ,B.POSITION                                                                  "
            sSQL = sSQL & vbCrLf & "          FROM ALL_CONSTRAINTS" & IIf(strReal = "A", "@LK_STG3", "") & "  A,                                                          "
            sSQL = sSQL & vbCrLf & "               ALL_CONS_COLUMNS" & IIf(strReal = "A", "@LK_STG3", "") & " B                                                           "
            sSQL = sSQL & vbCrLf & "           WHERE A.TABLE_NAME = '" & edtTable.Text & "' "
            If Trim(cboShema.Text) <> "" Then
                sSQL = sSQL & vbCrLf & "             AND A.OWNER      = '" & IIf(chkASIS.Value = 1, "IBKCDEV", UCase(cboShema.Text)) & "'"
            Else
                sSQL = sSQL & vbCrLf & "             AND A.OWNER      NOT IN  " & C_NOTIN
            End If
            sSQL = sSQL & vbCrLf & "             AND A.OWNER      = '" & Trim(txtOwner.Text) & "'"
            sSQL = sSQL & vbCrLf & "           AND A.CONSTRAINT_TYPE = 'P'                                                      "
            sSQL = sSQL & vbCrLf & "           AND A.OWNER   = B.OWNER                                                          "
            sSQL = sSQL & vbCrLf & "           AND A.CONSTRAINT_NAME = B.CONSTRAINT_NAME                                        "
            sSQL = sSQL & vbCrLf & "           AND A.TABLE_NAME = B.TABLE_NAME                                                  "
            sSQL = sSQL & vbCrLf & "         )  B                                                                               "
            sSQL = sSQL & vbCrLf & "           ON A.COLUMN_NAME = B.COLUMN_NAME                                                 "
    sSQL = sSQL & vbCrLf & "         )"
    sSQL = sSQL & vbCrLf & " ORDER BY COLUMN_ID                                                                         "
    

'End If

    txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If




    fpCol.MaxRows = 0
    If rsTemp.RecordCount < 1 Then
        fpCol.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpCol, rsTemp)
    End If


    'SQL display
    Call subMakeSQL


    'PK 색칠하기
    For i = 1 To fpCol.MaxRows
        With fpCol
            .Row = i
            .Col = 5
            If Trim(.Value) <> "" Then
                .Col = -1
                .BackColor = &HFFC0FF


            End If

        End With
    Next i


    fpIndex.MaxRows = 0
    '----------------------------------------------------------------------
    'Index 조회
    '----------------------------------------------------------------------
    Set rsTemp = Nothing
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient

    sSQL = ""
    sSQL = sSQL & vbCrLf & "SELECT  A.INDEX_NAME                        "
    sSQL = sSQL & vbCrLf & "       ,B.COLUMN_NAME                       "
    sSQL = sSQL & vbCrLf & "       ,B.DESCEND                           "
    sSQL = sSQL & vbCrLf & "       ,A.INDEX_TYPE                        "
    sSQL = sSQL & vbCrLf & "       ,A.TABLESPACE_NAME                   "
    sSQL = sSQL & vbCrLf & "  FROM  ALL_INDEXES" & IIf(strReal = "A", "@LK_STG3", "") & " A                       "
    sSQL = sSQL & vbCrLf & "       ,ALL_IND_COLUMNS" & IIf(strReal = "A", "@LK_STG3", "") & " B                   "
    'sSQL = sSQL & vbCrLf & " WHERE A.TABLE_OWNER = '" & cboShema.Text & "'"
    'sSQL = sSQL & vbCrLf & "   AND A.TABLE_NAME =  '" & edtTable.Text & "' "
    sSQL = sSQL & vbCrLf & "  WHERE A.TABLE_NAME =  '" & edtTable.Text & "' "
    sSQL = sSQL & vbCrLf & "  AND A.TABLE_OWNER = '" & txtOwner.Text & "'"
    sSQL = sSQL & vbCrLf & "   AND A.TABLE_OWNER = B.TABLE_OWNER          "
    sSQL = sSQL & vbCrLf & "   AND A.INDEX_NAME = B.INDEX_NAME          "
    sSQL = sSQL & vbCrLf & "   AND A.TABLE_NAME = B.TABLE_NAME          "
    sSQL = sSQL & vbCrLf & "ORDER BY A.INDEX_NAME, B.COLUMN_POSITION    "



    txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If


    If rsTemp.RecordCount < 1 Then
        fpIndex.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpIndex, rsTemp)
    End If

    
    Dim sKeyVal As String
    
   


    '----------------------------------------------------------------------
    '파티션 조회
    '----------------------------------------------------------------------
    Set rsTemp = Nothing
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient

    sSQL = ""
    sSQL = sSQL & vbCrLf & "select  TABLE_NAME, PARTITION_NAME                     "
    sSQL = sSQL & vbCrLf & "       ,HIGH_VALUE                         "
    sSQL = sSQL & vbCrLf & "       ,TABLESPACE_NAME                    "
    sSQL = sSQL & vbCrLf & "       ,COMPRESSION                        "
    sSQL = sSQL & vbCrLf & "       ,NUM_ROWS                           "
    sSQL = sSQL & vbCrLf & "       ,BLOCKS                             "
    sSQL = sSQL & vbCrLf & "       ,LAST_ANALYZED                      "
    sSQL = sSQL & vbCrLf & "  from all_TAB_PARTITIONS" & IIf(strReal = "A", "@LK_STG3", "") & "                 "
    sSQL = sSQL & vbCrLf & " WHERE TABLE_NAME = '" & edtTable.Text & "' "
    sSQL = sSQL & vbCrLf & "   AND TABLE_OWNER      = '" & IIf(chkASIS.Value = 1, "IBKCDEV", UCase(cboShema.Text)) & "'"
    sSQL = sSQL & vbCrLf & " ORDER BY PARTITION_NAME                   "




    txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If


    fpPart.MaxRows = 0
    If rsTemp.RecordCount < 1 Then
        fpPart.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpPart, rsTemp)
    End If



    '----------------------------------------------------------------------
    '테이블정보 조회
    '----------------------------------------------------------------------
    Set rsTemp = Nothing
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient

    sSQL = ""
    sSQL = sSQL & vbCrLf & "select  TABLE_NAME, TABLESPACE_NAME,STATUS, NUM_ROWS, BLOCKS,                     "
    sSQL = sSQL & vbCrLf & "       TABLE_LOCK, SAMPLE_SIZE, LAST_ANALYZED, PARTITIONED                      "
    sSQL = sSQL & vbCrLf & "  from all_tables" & IIf(strReal = "A", "@LK_STG3", "") & "                 "
    sSQL = sSQL & vbCrLf & " WHERE TABLE_NAME = '" & edtTable.Text & "' "
    If Trim(cboShema.Text) <> "" Then
        sSQL = sSQL & vbCrLf & "   AND OWNER      = '" & IIf(chkASIS.Value = 1, "IBKCDEV", UCase(cboShema.Text)) & "'"
    Else
        sSQL = sSQL & vbCrLf & "   AND OWNER      NOT IN  " & C_NOTIN
    End If


    txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If


    fpInfo.MaxRows = 0
    If rsTemp.RecordCount < 1 Then
        fpInfo.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpInfo, rsTemp)
    End If



    '샘플데이터 조회
    subSelectData ("")

    '테이블상세정보 조회
    'Call subSelectTableDesc

    'group by col
    fpCol.Row = 1
    fpCol.Col = 1
'    txtGrpByCol.Text = fpCol.Text
'    txtGrpByCol2.Text = fpCol.Text
'    fpRecCount.MaxRows = 0



    '----------------------------------------------------------------------
    '최근 검색에 1줄 추가
    '----------------------------------------------------------------------
        
    '중복 최근 검색 삭제
    For i = 1 To fpNear.MaxRows
        fpNear.Row = i
        fpNear.Col = 1
        
        If fpNear.Value = edtTable.Text Then
            fpNear.DeleteRows i, 1
        End If
    
    Next i
    
    fpNear.MaxRows = fpNear.MaxRows + 1
    fpNear.InsertRows 1, 1
    fpNear.Row = 1
    fpNear.Col = 1
    fpNear.Value = edtTable.Text
    fpNear.Col = 2
    fpNear.Value = edtTablenm.Text
    fpNear.Col = 5
    fpNear.Value = txtOwner.Text
    
    
    '공백 삭제
    For i = 1 To fpNear.MaxRows
        fpNear.Row = i
        fpNear.Col = 1
        
        If Trim(fpNear.Value) = "" Then
            fpNear.DeleteRows i, 1
            fpNear.MaxRows = fpNear.MaxRows - 1
        End If
    
    Next i
    
    
    
    'po 관련 문장 생성 - Delete
    setPoDelete
    

    '검색내용 파일 저장
    Call saveNearFile
    

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


'------------------------------------------------------
' PO Delete 문 생성
'------------------------------------------------------
Private Sub setPoDelete()

On Error GoTo ErrHandler
    Dim sTemp As String
    txtPoDelete.Text = ""
    
    sTemp = ""
        
    Dim sDVO As String
        

    Dim sArr() As String
    Dim i As Integer
    
    sArr = Split(edtTable.Text, "_")
    sDVO = ""
    For i = LBound(sArr) To UBound(sArr)
        sDVO = sDVO & gfUpperFirst(sArr(i))
    Next i
    sDVO = sDVO & "001Dvo"
    
    
'여기

    sTemp = sTemp & "DataObjectMapper<" & sDVO & "> mapper = new DataObjectMapper<" & sDVO & ">( ""po7_app"", " & sDVO & ".class );"
    sTemp = sTemp & vbCrLf & "mapper.setQuery(" & sDVO & ".DEL001);"
       
    Dim sColName As String
    Dim sDBColName As String
    
    Dim j As Integer
    
    Dim sTempNotPk As String
    
    sTempNotPk = ""
    
    Dim sSizeData As String
    
    For i = 1 To fpCol.MaxRows
        With fpCol
            .Row = i
                        
            
            .Col = 1
            
            sArr = Split(.Value, "_")
              
            sColName = ""
            sDBColName = ""
            
            For j = LBound(sArr) To UBound(sArr)
                sColName = sColName & gfUpperFirst(sArr(j))
                
                If j = 0 Then
                    sDBColName = sDBColName & LCase(sArr(j))
                Else
                    sDBColName = sDBColName & gfUpperFirst(sArr(j))
                End If
                
            Next j
            
            .Col = 3
            If .Value = "NUMBER" Then
                sColName = "l" & sColName
            Else
                sColName = "str" & sColName
            End If
                        
            'PK인경무
            .Col = 5
            If Trim(.Value) <> "" Then
            
                .Col = 2
                sTemp = sTemp & vbCrLf & "mapper.setParameter(""" & sDBColName & """, " & sColName & ");          // " & .Value
            
            Else
                .Col = 2
                sTempNotPk = sTempNotPk & vbCrLf & "mapper.setParameter(""" & sDBColName & """, " & sColName & ");          // " & .Value
                
            End If
             
                
            
        End With
    Next i
    sTemp = sTemp & vbCrLf & "mapper.remove();"
    
                     
    
    
    txtPoDelete.Text = sTemp
    txtMapper.Text = sTempNotPk
    
    
    
    '주석 정렬
    
    Dim sOrgArr() As String
    Dim sSplit() As String
    Dim iMaxSize As Integer
    
    
    sOrgArr = Split(txtPoDelete.Text, vbCrLf)
    
    txtPoDelete.Text = ""
    iMaxSize = 0
    For i = LBound(sOrgArr) To UBound(sOrgArr)
                    
                    
        If InStr(1, UCase(sOrgArr(i)), " // ") Then
            sSplit = Split(sOrgArr(i), " // ")
            If iMaxSize < Len(sSplit(0)) Then
                iMaxSize = Len(sSplit(0))
            End If
        
        Else
            
        End If
    Next i
    
    For i = LBound(sOrgArr) To UBound(sOrgArr)
                    
                    
        If InStr(1, UCase(sOrgArr(i)), " // ") Then
            sSplit = Split(sOrgArr(i), " // ")
            
            txtPoDelete.Text = txtPoDelete.Text & sSplit(0) & gfMakeString(" ", iMaxSize - Len(sSplit(0)) + 5) & " // "
            txtPoDelete.Text = txtPoDelete.Text & sSplit(1) & vbCrLf
        Else
            txtPoDelete.Text = txtPoDelete.Text & sOrgArr(i) & vbCrLf
        End If
    Next i
    
    
    '주석 정렬 txtmapper
    sOrgArr = Split(txtMapper.Text, vbCrLf)
    
    txtMapper.Text = ""
    iMaxSize = 0
    For i = LBound(sOrgArr) To UBound(sOrgArr)
                    
                    
        If InStr(1, UCase(sOrgArr(i)), " // ") Then
            sSplit = Split(sOrgArr(i), " // ")
            If iMaxSize < Len(sSplit(0)) Then
                iMaxSize = Len(sSplit(0))
            End If
        
        Else
            
        End If
    Next i
    
    For i = LBound(sOrgArr) To UBound(sOrgArr)
                    
                    
        If InStr(1, UCase(sOrgArr(i)), " // ") Then
            sSplit = Split(sOrgArr(i), " // ")
            
            txtMapper.Text = txtMapper.Text & sSplit(0) & gfMakeString(" ", iMaxSize - Len(sSplit(0)) + 5) & " // "
            txtMapper.Text = txtMapper.Text & sSplit(1) & vbCrLf
        Else
            txtMapper.Text = txtMapper.Text & sOrgArr(i) & vbCrLf
        End If
    Next i
    
        
    

    
    Exit Sub
ErrHandler:
    gfMsgWrn Err.Description

End Sub




'------------------------------------------------------
' 샘플데이터 조회
'------------------------------------------------------
Private Sub subMakeSQL()
    Dim i As Integer
    
    
    Dim sCode As String
    Dim sName As String
    Dim sIQueryCode As String
    
    Dim iMaxLen As Integer
    iMaxLen = 0
    Dim iSpace As Integer
    iSpace = 10
    Dim sTy As String
    
    '최대 컬럼 사이즈 계산
    For i = 1 To fpCol.MaxRows
        fpCol.Row = i
        fpCol.Col = 1
        
        If Len(Trim(fpCol.Value)) > iMaxLen Then
            iMaxLen = Len(Trim(fpCol.Value))
        End If
        
    Next i
    
    
    Dim sISQL As String
    Dim sISQLValue As String
    sISQLValue = ""
    
    
    Dim sIQuerySQL As String
    Dim sIQuerySQLValue As String
    sIQuerySQLValue = ""
    
    Dim sUSQL As String
    Dim sUSQLWhere As String
    sUSQLWhere = ""
    
    Dim sSSQL As String
    Dim sSSQL2 As String
    
    
    Dim sGSQL As String
    
    sISQL = "INSERT INTO " & edtTable.Text & " ( " & "        /* " & edtTablenm.Text & "*/" & vbCrLf
    sISQLValue = "( "
    
    
    sIQuerySQL = "INSERT INTO " & edtTable.Text & " ( " & vbCrLf
    sIQuerySQLValue = "( "
    
    sUSQL = "UPDATE " & edtTable.Text & "        /* " & edtTablenm.Text & "*/" & vbCrLf & "   SET " & vbCrLf
    sUSQLWhere = " WHERE 1=1" & vbCrLf
    
    
    sSSQL = "SELECT "
    sSSQL2 = "SELECT "
    
    sGSQL = " "
    
    Dim sCodeArr() As String
    Dim x As Integer
    Dim sNvlWord As String
    
    For i = 1 To fpCol.MaxRows
        fpCol.Row = i
        fpCol.Col = 1
        sCode = Trim(fpCol.Value)
        fpCol.Col = 2
        sName = Trim(fpCol.Value)
        fpCol.Col = 3
        sTy = Left(fpCol.Value, 1)
        
        sIQueryCode = ""
        
        sCodeArr = Split(sCode, "_")
        
        'po명 변수명 만들기
        For x = LBound(sCodeArr) To UBound(sCodeArr)
            
            If x = 0 Then
                sIQueryCode = sIQueryCode & LCase(sCodeArr(x))
            Else
                sIQueryCode = sIQueryCode & gfUpperFirst(sCodeArr(x))
            End If
        Next x
        
        
        
        'insert field
        If i <> 1 Then
            sISQL = sISQL & "         ,"
            sIQuerySQL = sIQuerySQL & "         ,"
        Else
            sISQL = sISQL & "          "
            sIQuerySQL = sIQuerySQL & "          "
        End If
        
        
        sISQL = sISQL & "" & sCode
        sISQL = sISQL & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
        
        
        sIQuerySQL = sIQuerySQL & "" & sCode
        sIQuerySQL = sIQuerySQL & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)

        
        If i <> fpCol.MaxRows Then
            If chkComment.Value = "1" Then
                sISQL = sISQL & "   -- " & sName
                sIQuerySQL = sIQuerySQL & "   -- " & sName & " "
            End If
            
        Else
            If chkComment.Value = "1" Then
                sISQL = sISQL & "     -- " & sName
            
                sIQuerySQL = sIQuerySQL & "   -- " & sName & " "
            End If
        End If
        sISQL = sISQL & vbCrLf
        sIQuerySQL = sIQuerySQL & vbCrLf
        
        
        
        'insert values
        If i <> 1 Then
            sISQLValue = sISQLValue & "       ,"
            
            sIQuerySQLValue = sIQuerySQLValue & "       ,"
            
        Else
            sISQLValue = sISQLValue & "      "
            
            sIQuerySQLValue = sIQuerySQLValue & "      "
        End If
        
        
        
        
            If sTy = "N" Or sTy = "I" Then
                sISQLValue = sISQLValue & "XXX"
                sISQLValue = sISQLValue & "            -- " & sName
            Else
                sISQLValue = sISQLValue & "'XXX'"
                sISQLValue = sISQLValue & "          -- " & sName
            End If
        
        
        
        'IQuery의 경우
        
        If sCode = "FST_REG_DY" Then
            sIQuerySQLValue = sIQuerySQLValue & "SYSDATE"
            sNvlWord = "SYSDATE"
        'ElseIf sCode = "SYS_REG_BR_CD" Then
        '    sIQuerySQLValue = sIQuerySQLValue & "#SYS_BR_CD#  "
        
        'ElseIf sCode = "SYS_RGSR_EMNO" Then
        '    sIQuerySQLValue = sIQuerySQLValue & "#SYS_EMNO# "
        
        ElseIf sCode = "LAST_CHG_DY" Then
            sIQuerySQLValue = sIQuerySQLValue & "SYSDATE"
            sNvlWord = "SYSDATE"
        'ElseIf sCode = "SYS_CHG_BR_CD" Then
        '    sIQuerySQLValue = sIQuerySQLValue & "#SYS_BR_CD#"
        
        'ElseIf sCode = "SYS_EDIR_EMNO" Then
        '    sIQuerySQLValue = sIQuerySQLValue & "#SYS_EMNO#"
        
        ElseIf Right(sName, 2) = "코드" Or Right(sName, 2) = "여부" Then
        
            sIQuerySQLValue = sIQuerySQLValue & "NVL(:" & sIQueryCode & ",'-')"
            sNvlWord = "NVL(:" & sIQueryCode & ",'-')"
        
        ElseIf sTy = "N" Or sTy = "I" Then
           
            sIQuerySQLValue = sIQuerySQLValue & "NVL(:" & sIQueryCode & ",0)"
            sNvlWord = "NVL(:" & sIQueryCode & ",0)"
        Else
            
            sIQuerySQLValue = sIQuerySQLValue & ":" & sIQueryCode
            sNvlWord = ":" & sIQueryCode
        End If
        

        If chkComment.Value = "1" Then
            
            sIQuerySQLValue = sIQuerySQLValue & gfMakeString(" ", iMaxLen - Len(sNvlWord) + iSpace)
            sIQuerySQLValue = sIQuerySQLValue & "   -- " & sName
            
        End If
        
        
        
        sISQLValue = sISQLValue & vbCrLf
        
        sIQuerySQLValue = sIQuerySQLValue & vbCrLf
        
        
        'update field
        If i <> 1 Then
            sUSQL = sUSQL & "       ,"
        Else
            sUSQL = sUSQL & "        "
        End If
        
        sUSQL = sUSQL & "" & sCode
        sUSQL = sUSQL & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
        
        
        If sTy = "N" Or sTy = "I" Then
            sUSQL = sUSQL & "    = XXX"
            sUSQL = sUSQL & "            -- " & sName
        Else
            sUSQL = sUSQL & "    = 'XXX'"
            sUSQL = sUSQL & "          -- " & sName
        End If
        
        sUSQL = sUSQL & vbCrLf
        
        'where
        'If Trim(sUSQLWhere) = "" Then
        '    sUSQLWhere = sUSQLWhere & " WHERE "
        'Else
        '    sUSQLWhere = sUSQLWhere & "   AND "
        'End If
            sUSQLWhere = sUSQLWhere & "   AND "
        
        sUSQLWhere = sUSQLWhere & IIf(Trim(txtAlias.Text) <> "", Trim(txtAlias.Text) & ".", "") & sCode
        sUSQLWhere = sUSQLWhere & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
        If sTy = "N" Or sTy = "I" Then
            sUSQLWhere = sUSQLWhere & "    = XXX"
            sUSQLWhere = sUSQLWhere & "            -- " & sName
        Else
            sUSQLWhere = sUSQLWhere & "    = 'XXX'"
            sUSQLWhere = sUSQLWhere & "          -- " & sName
        End If
        
        sUSQLWhere = sUSQLWhere & vbCrLf
        
        
        'SELECT FIELD
        
        If i <> 1 Then
            sSSQL = sSSQL & "       ,"
            sSSQL2 = sSSQL2 & "       ,"
            
            fpCol.Col = 5
            If Trim(fpCol.Value) <> "" Then
                sGSQL = sGSQL & "       ,"
            End If
        End If
        
'-------------------------------------
        
        
        If i = 1 Then
            sSSQL = sSSQL & " " & IIf(Trim(txtAlias.Text) <> "", Trim(txtAlias.Text) & ".", "") & sCode
            sSSQL2 = sSSQL2 & " " & IIf(Trim(txtAlias.Text) <> "", Trim(txtAlias.Text) & ".", "") & sCode
            
            fpCol.Col = 5
            If Trim(fpCol.Value) <> "" Then
                sGSQL = sGSQL & " " & sCode
            End If
                
            
        Else
            sSSQL = sSSQL & "" & IIf(Trim(txtAlias.Text) <> "", Trim(txtAlias.Text) & ".", "") & sCode
            sSSQL2 = sSSQL2 & "" & IIf(Trim(txtAlias.Text) <> "", Trim(txtAlias.Text) & ".", "") & sCode
        
            fpCol.Col = 5
            If Trim(fpCol.Value) <> "" Then
                sGSQL = sGSQL & " " & sCode
            End If
        

        
        
        End If
        
        
        
        sSSQL = sSSQL & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
        sSSQL2 = sSSQL2 & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
        
        
        
        
        'as 논리명 처리
        If chkAS.Value = 1 Then
            sSSQL = sSSQL & "   AS """ & sName & """ "
            sSSQL2 = sSSQL2 & "   AS """ & sName & """ "
        Else
            sSSQL = sSSQL & "   -- " & sName
            sSSQL2 = sSSQL2 & "   AS " & sCode
            sSSQL2 = sSSQL2 & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
            sSSQL2 = sSSQL2 & "   -- " & sName
        End If
        
        
        
        
'-----------------------------------


  
                
                
        
        If Right(sName, 2) = "코드" Or Trim(sName) = "고객관리번호" Or Right(sName, 4) = "부점코드" Or Right(sName, 4) = "직원번호" Then
            
            sSSQL2 = sSSQL2 & vbCrLf
            sSSQL2 = sSSQL2 & "       ,"
            sSSQL2 = sSSQL2 & "" & IIf(Trim(txtAlias.Text) <> "", Trim(txtAlias.Text) & ".", "") & sCode
        
            sSSQL2 = sSSQL2 & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
            
            sSSQL2 = sSSQL2 & "   AS " & sCode & "_NM"
            sSSQL2 = sSSQL2 & gfMakeString(" ", iMaxLen - Len(sCode) - 3 + iSpace)
            sSSQL2 = sSSQL2 & "   -- " & sName & "명"
        End If

        
        
        
        
        sSSQL = sSSQL & vbCrLf
        sSSQL2 = sSSQL2 & vbCrLf
        
        fpCol.Col = 5
        If Trim(fpCol.Value) <> "" Then
            sGSQL = sGSQL & vbCrLf
        End If
        
        
    Next i
    
    txtIQuery.Text = ""
'    txtIQuery.Text = txtIQuery.Text & vbCrLf & "/*******************************"
'    txtIQuery.Text = txtIQuery.Text & vbCrLf & "@시스템명   : IBK캐피탈 계정계 시스템"
'    txtIQuery.Text = txtIQuery.Text & vbCrLf & "@업무분류   :"
'    txtIQuery.Text = txtIQuery.Text & vbCrLf & "@쿼리이름   : "
'    txtIQuery.Text = txtIQuery.Text & vbCrLf & "@쿼리개요   :"
'    txtIQuery.Text = txtIQuery.Text & vbCrLf & "@작성일자   : " & Now
'    txtIQuery.Text = txtIQuery.Text & vbCrLf & "@작성자 : 김경학"
'    txtIQuery.Text = txtIQuery.Text & vbCrLf & "@수정내역   :"
'    txtIQuery.Text = txtIQuery.Text & vbCrLf & "*******************************/"

    
    sIQuerySQL = sIQuerySQL & ")"
    sIQuerySQLValue = sIQuerySQLValue & ")"
    txtIQuery.Text = txtIQuery.Text & vbCrLf & sIQuerySQL & " VALUES " & vbCrLf & sIQuerySQLValue
    
    
    sISQL = sISQL & ")"
    sISQLValue = sISQLValue & ")"
    
    txtInsertSQL.Text = sISQL & " VALUES " & vbCrLf & sISQLValue
    
    
    txUpdateSQL.Text = sUSQL & sUSQLWhere
      
    'txtSelectSQL.Text = sSSQL & "  FROM " & edtTable.Text & IIf(Trim(txtAlias.Text) <> "", "  " & Trim(txtAlias.Text), "") & _
                                "        /* " & edtTablenm.Text & "*/" & vbCrLf & sUSQLWhere
                                
                                
    txtSelectSQL.Text = sSSQL & "  FROM " & edtTable.Text & IIf(strReal = "A", "@LK_STG3", "") & IIf(Trim(txtAlias.Text) <> "", "  " & Trim(txtAlias.Text), "") & _
                                "        /* " & edtTablenm.Text & "*/" & vbCrLf & sUSQLWhere
                                
    txtSQLcdnm.Text = sSSQL2 & "  FROM " & edtTable.Text & IIf(strReal = "A", "@LK_STG3", "") & IIf(Trim(txtAlias.Text) <> "", "  " & Trim(txtAlias.Text), "") & _
                                "        /* " & edtTablenm.Text & "*/" & vbCrLf & sUSQLWhere
    
    'txtInsSelect.Text = sISQL & vbCrLf & sSSQL & "  FROM " & edtTable.Text & vbCrLf & sUSQLWhere
    txtInsSelect.Text = sISQL & vbCrLf & sSSQL & "  FROM " & edtTable.Text & IIf(strReal = "A", "@LK_STG3", "") & vbCrLf & sUSQLWhere
    
    
'    txtGrpBySQL.Text = "SELECT " & sGSQL & "       , COUNT(*)  " & vbCrLf & " FROM " & edtTable.Text & vbCrLf & " GROUP BY " & sGSQL
End Sub




'------------------------------------------------------
' 샘플데이터 조회
'------------------------------------------------------
Private Sub subSelectData(ByVal sReqSQL As String)

    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
                    

                        

                            
    Dim sField As String
    Dim sWhere As String
    Dim i As Integer
    Dim sCol As String
    Dim sTemp As String
   
    Dim blnValue As Boolean
    
    
    Screen.MousePointer = vbHourglass
    
    If GCONN_USE = False Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = GetConnString
        conn.Open
    End If
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient


    blnValue = False
    
    Dim sType As String
    
    'SQL문 생성
    With fpCol
        
        sField = ""
        
        For i = 1 To .MaxRows
            
            .Row = i
            .Col = 1
            sCol = .Value
            
            .Col = 3
            sType = .Value
            
            If UCase(Left(sType, 4)) = "TIME" Then
                
                sField = sField & "to_char(" & sCol & ", 'YYYY-MM-DD HH24:mi:ss' )"
            Else
                sField = sField & sCol
            End If
            
            
            
            
            
            If i <> .MaxRows Then sField = sField & ","
            
            'VALUE 조회인경우
            sTemp = ""
            .Col = 10
            Dim sOP As String
            Dim sWhereValue As String
            If Trim(.Value) <> "" Then
                sWhereValue = Replace(.Value, "-", "")
                blnValue = True
                
                .Col = 9
                sOP = Trim(.Value)
                
                .Col = 3
                If Left(fpCol.Value, 1) = "N" Or Left(fpCol.Value, 1) = "I" Then
                    .Col = 10
                    If sOP = "%%" Then
                        sTemp = sCol & " LIKE '%" & sWhereValue & "%'"
                    ElseIf sOP = "%" Then
                        sTemp = sCol & " LIKE '" & sWhereValue & "%'"
                    Else
                        sTemp = sCol & " = " & sWhereValue
                    End If
                
                Else
                    .Col = 10
                    
                    If sOP = "%%" Then
                        sTemp = sCol & " LIKE '%" & sWhereValue & "%'"
                    ElseIf sOP = "%" Then
                        sTemp = sCol & " LIKE '" & sWhereValue & "%'"
                    Else
                        sTemp = sCol & " = '" & sWhereValue & "'"
                    End If
                    
                    
                End If
            
            End If
            If sTemp <> "" Then
                If Trim(sWhere) = "" Then
                    sWhere = sWhere & " WHERE " & sTemp
                Else
                    sWhere = sWhere & " AND " & sTemp
                End If
            End If
            
        Next i
    End With
    
    sField = sField & " , ROWID  AS ROW_ID "
    If Trim(sReqSQL) = "" Then
        sSQL = "SELECT " & sField & " FROM " & Trim(txtOwner.Text) & "." & edtTable.Text & IIf(strReal = "A", "@LK_STG3", "") & sWhere
                       
        'value 조회가 아닌경우 만 데이터 제한
        'If blnValue = False Then
            If sWhere = "" Then
                sSQL = sSQL & " WHERE ROWNUM < 100"
            Else
                sSQL = sSQL & " AND ROWNUM < 100"
            End If
        'End If
    Else
        sSQL = sReqSQL

    End If
    
    
 
    txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If
    
    
    fpSample.MaxCols = fpCol.MaxRows + 1
    
    fpSample.MaxRows = 0
    
    '헤더설정
    
    
    ReDim gColId(fpCol.MaxRows)
    ReDim gColName(fpCol.MaxRows)
    ReDim gPK(fpCol.MaxRows)
    ReDim gType(fpCol.MaxRows)
    
    
    For i = 1 To fpCol.MaxRows
        fpCol.Row = i
        fpCol.Col = 2
        sTemp = fpCol.Value
        
        If sTemp = "" Then
            fpCol.Col = 1
            sTemp = fpCol.Value
        End If
        
        
        fpSample.Row = 0
        fpSample.Col = i
        fpSample.Value = fpCol.Value
        
        fpCol.Col = 1: gColId(i) = fpCol.Value
        fpCol.Col = 2: gColName(i) = fpCol.Value
        fpCol.Col = 3: gType(i) = fpCol.Value
        fpCol.Col = 5: gPK(i) = fpCol.Value
        
        
    Next i
    
    fpSample.Row = 0
    fpSample.Col = i + 1
    fpSample.Value = "ROWID"
    
    
    If rsTemp.RecordCount < 1 Then
        fpSample.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpSample, rsTemp)
    End If
    
        
        
        
        
    '그리드 초기화
    With fpSample
        For i = 1 To fpSample.MaxCols
            .ColWidth(i) = 15
        Next i
        
        '.Col = -1
        '.Row = -1
        '.CellType = CellTypeEdit
        '.TypeEditMultiLine = True
        '.TypeMaxEditLen = 9999
        
        '.Col = 0
        '.Row = 0
        
    End With


    
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
'직원정보 검색
Private Sub cmdFindEmp_Click()
    'http://134.100.1.60/com.ibk.finehr.hrm.comm.EmpnoSearchEmpnoInfoSelAction.action?empno=20772&sejno=A97011&fromBr=N
    'http://134.100.1.60/com.ibk.finehr.hrm.comm.EmpnoSearchEmpnoListSelAction.action?mname=오성훈&sejno=A97011
    
    Dim sUrl As String
    
    
    
    '기업은행 조회
    If Trim(txtIBKBankEmp.Text) <> "" Then
        If IsNumeric(Trim(txtIBKBankEmp.Text)) Then
            sUrl = "http://134.100.1.60/com.ibk.finehr.hrm.comm.EmpnoSearchEmpnoInfoSelAction.action?empno=" & Trim(txtIBKBankEmp.Text)
        Else
            sUrl = "http://134.100.1.60/com.ibk.finehr.hrm.comm.EmpnoSearchEmpnoListSelAction.action?mname=" & Trim(txtIBKBankEmp.Text)
        End If
    End If
    
    
    WebBrowser1.Navigate (sUrl)
End Sub

Private Sub cmdFreeSQL_Click()
    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
                    

    Dim i As Integer
    

    If InStr(1, UCase(txtReqSQL), "INSERT") > 0 Then
        gfMsg ("INSERT, UPDATE, DELETE, CREATE 구문은 사용 할 수 없습니다.")
        Exit Sub
    End If
    
    
    If InStr(1, UCase(txtReqSQL), "UPDATE") > 0 Then
        gfMsg ("INSERT, UPDATE, DELETE, CREATE 구문은 사용 할 수 없습니다.")
        Exit Sub
    End If
    
    If InStr(1, UCase(txtReqSQL), "DELETE") > 0 Then
        gfMsg ("INSERT, UPDATE, DELETE, CREATE 구문은 사용 할 수 없습니다.")
        Exit Sub
    End If
    
    
    If InStr(1, UCase(txtReqSQL), "CREATE") > 0 Then
        gfMsg ("INSERT, UPDATE, DELETE, CREATE 구문은 사용 할 수 없습니다.")
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
   
    
    
    
    sSQL = txtReqSQL.Text
    
    txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If
    

    For i = 1 To rsTemp.Fields.Count
        fpSample.Row = 0
        fpSample.Col = i
        fpSample.Value = rsTemp.Fields(i - 1).Name
        
    Next i
    fpSample.MaxCols = rsTemp.Fields.Count


    fpSample.MaxRows = 0
    If rsTemp.RecordCount < 1 Then
        fpSample.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpSample, rsTemp)
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



Private Sub getClipinsert_Click()
    
    txtInsertData.SelStart = 1
    txtInsertData.SelLength = Len(txtInsertData.Text)

    Clipboard.SetText (txtInsertData.Text)
End Sub

Private Sub Image1_Click()
    frameCode.Visible = False
End Sub





Private Sub txtIBKSystemEmp_GotFocus()
    txtIBKBankEmp.Text = ""
End Sub

Private Sub Image2_Click()
    'frmSrc.Visible = False
End Sub

Private Sub Image3_Click()
    frmNeed.Visible = False
End Sub

Private Sub lblFavo_Click()
    If lblFavo.Caption = "☆" Then
        lblFavo.Caption = "★"
    Else
        lblFavo.Caption = "☆"
    End If
    
End Sub

Private Sub MSChart1_OLEStartDrag(Data As MSChart20Lib.DataObject, AllowedEffects As Long)

End Sub

Private Sub optMigAll_Click()
    txtMoveCnt.Text = ""
    txtMoveCnt.Locked = True
    txtMoveCnt.BackColor = &HECD9DD
End Sub


Private Sub optMigSome_Click()
    txtMoveCnt.Text = ""
    txtMoveCnt.Locked = False
    txtMoveCnt.BackColor = &H80000005
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 1 Then
'        txtTableSrchKor.SetFocus
    End If
    
    If SSTab1.Tab = 4 Then
'        TimerRefresh.Enabled = True
'        TimerRefresh.Interval = CDbl(txtRefrsh.Text) * 1000
'        Call cmdCloseSrch_Click
    Else
        'TimerRefresh.Enabled = False
    End If
    
    
End Sub

Private Sub Timer1_Timer()

    If Timer1.Tag = 1 Then
        If picMatch.Visible = True Then
            picMatch.Visible = False
            Timer1.Enabled = False
        End If
    End If
    
    Timer1.Tag = 1
    
    
End Sub


Private Sub txtCodeSrch_GotFocus()
    txtCodeSrch.SelStart = 0
    txtCodeSrch.SelLength = Len(txtCodeSrch.Text)
    
End Sub

Private Sub txtCodeSrch_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim iRow As Integer
    If KeyAscii <> 13 Then
        
        Exit Sub
        
    End If
    Dim sSearchKey As String
    sSearchKey = Trim(UCase(txtCodeSrch.Text))
    
    
    iRow = fpcode.SearchCol(4, 1, fpcode.MaxRows, sSearchKey, SearchFlagsValue)
    If iRow > 0 Then
        Call fpcode.ShowCell(1, IIf(iRow - 3 <= 1, 1, iRow - 3), PositionUpperCenter)
        fpcode.Col = -1
        fpcode.Row = iRow
        If optColor(0).Value = True Then fpcode.BackColor = S_CYAN
        If optColor(1).Value = True Then fpcode.BackColor = S_YELLOW
        If optColor(2).Value = True Then fpcode.BackColor = S_BROWN
        Exit Sub
    End If
    
    iRow = fpcode.SearchCol(5, 1, fpcode.MaxRows, sSearchKey, SearchFlagsValue)
    If iRow > 0 Then
        
        Call fpcode.ShowCell(1, IIf(iRow - 3 <= 1, 1, iRow - 3), PositionUpperCenter)
        fpcode.Col = -1
        fpcode.Row = iRow
        If optColor(0).Value = True Then fpcode.BackColor = S_CYAN
        If optColor(1).Value = True Then fpcode.BackColor = S_YELLOW
        If optColor(2).Value = True Then fpcode.BackColor = S_BROWN
        Exit Sub
    End If

    
    
   
    Dim iMatchCount As Integer
    Dim iFirstRow As Integer
    

    iMatchCount = 0
    iFirstRow = 0
    
    For i = 1 To fpcode.MaxRows
        fpcode.Row = i
        fpcode.Col = 4
        
        If InStr(1, Trim(fpcode.Value), sSearchKey) > 0 Then
        
            iMatchCount = iMatchCount + 1
            
            iRow = i
            
            If iFirstRow = 0 Then
                iFirstRow = iRow
            End If
            

            fpcode.Col = -1
            fpcode.Row = iRow
            fpcode.BackColor = S_LIGHT_GRAY
        
        End If
        
        fpcode.Col = 5

        If InStr(1, Trim(fpcode.Value), sSearchKey) > 0 Then
            
            iMatchCount = iMatchCount + 1
            
            iRow = i
            
            
            If iFirstRow = 0 Then
                iFirstRow = iRow
            End If
            
            
            
            fpcode.Col = -1
            fpcode.Row = iRow
            fpcode.BackColor = S_LIGHT_GRAY
        
        End If
    
    Next i
    
    

    Call fpcode.ShowCell(1, IIf(iFirstRow - 3 <= 1, 1, iFirstRow - 3), PositionUpperCenter)

    
    If iMatchCount > 1 Then
        lblMatch.Caption = "Matched " & iMatchCount & " columns...."
        picMatch.Visible = True
    End If

End Sub

Private Sub txtCodeSrchNm_GotFocus()
    txtCodeSrchNm.SelStart = 0
    txtCodeSrchNm.SelLength = Len(txtCodeSrchNm.Text)
End Sub

'코드 조회 버튼 클릭
Private Sub txtCodeSrchNm_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    
    '공통코드 조회
    Call subSelectCodeList("")
    frameCode.Visible = True
    
    
End Sub

Private Sub txtCodeWhere_GotFocus(Index As Integer)
    txtCodeWhere(Index).SelStart = 0
    txtCodeWhere(Index).SelLength = Len(txtCodeWhere(Index).Text)
    
End Sub

Private Sub txtDicSQL_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 And KeyCode = 65 Then
        txtDicSQL.SelStart = 0
        txtDicSQL.SelLength = Len(txtDicSQL.Text)
    End If
    
End Sub

Private Sub txtIBKBankEmp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call cmdFindEmp_Click
    End If
End Sub

Private Sub txtInsertSQL_KeyDown(KeyCode As Integer, Shift As Integer)

    If Shift = 2 And KeyCode = 65 Then
        txtInsertSQL.SelStart = 0
        txtInsertSQL.SelLength = Len(txtInsertSQL.Text)
    End If
    

End Sub

Private Sub txtInsSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 And KeyCode = 65 Then
        txtInsSelect.SelStart = 0
        txtInsSelect.SelLength = Len(txtInsSelect.Text)
    End If
    
End Sub

Private Sub txtLdgrNo_GotFocus()
    txtLdgrNo.SelStart = 0
    txtLdgrNo.SelLength = Len(txtLdgrNo.Text)

End Sub

Private Sub txtLdgrNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call cmdSearchLdgr_Click
    End If
End Sub

'여신원장 관련 데이터를 조회한다.
Private Sub cmdSearchLdgr_Click()


    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

'On Error GoTo ErrorHandler
                    
                            
    Dim sField As String
    Dim sWhere As String
    Dim i, j As Integer
    Dim sCol As String
    Dim sTemp As String
   
    Dim blnValue As Boolean
    
    '원장번호가 입력되지 않았으면 exit
    If Trim(txtLdgrNo.Text) = "" Then Exit Sub
    
    
    
    Screen.MousePointer = vbHourglass
    
    
    
    
    
    If GCONN_USE = False Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = GetConnString
        conn.Open
    End If
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
    
    Dim sTableName As String
    Dim objGrid As Object
    
    For i = 1 To 100
        sSQL = ""
        
        sTableName = ""
        '1 여신원장
        If i = 1 Then sTableName = "TLNI_LNLDGR"
        If i = 2 Then sTableName = "TLNE_LDGRADHIST"
        If i = 3 Then sTableName = "TLNG_DLY"
        If i = 4 Then sTableName = "TLNG_LDGRTRHIST"
        If i = 5 Then sTableName = "TLNG_TRAMTCALC"
        If i = 6 Then sTableName = "TLNE_ADAMTCALC"
        If i = 7 Then sTableName = "TLNS_RFND"
        If i = 8 Then sTableName = "TLNS_RFNDTRHIST"
        If i = 9 Then sTableName = "TLNS_RFAMTCALC"
        If i = 10 Then sTableName = "TLNS_SPCLLDGR"
        If i = 11 Then sTableName = "TLNS_SPCLTRHIST"
        If i = 12 Then sTableName = "TLNS_SPAMTCALC"
        
        sSQL = sSQL & " SELECT B.COLUMN_NAME, COMMENTS"
        sSQL = sSQL & "  FROM ALL_COL_COMMENTS A"
        sSQL = sSQL & "      ,ALL_TAB_COLUMNS B"
        sSQL = sSQL & " Where 1 = 1"
        sSQL = sSQL & "   AND A.TABLE_NAME = UPPER('" & sTableName & "')"
        sSQL = sSQL & "   AND A.OWNER(+) = B.OWNER"
        sSQL = sSQL & "   AND A.TABLE_NAME(+) = B.TABLE_NAME"
        sSQL = sSQL & "   AND A.COLUMN_NAME(+) = B.COLUMN_NAME"
        sSQL = sSQL & "  ORDER BY B.COLUMN_ID"
        
        If sTableName <> "" Then
        
            Set rsTemp = Nothing
            Set rsTemp = New ADODB.Recordset
            rsTemp.CursorLocation = adUseClient
        
            txtSQL.Text = sSQL
            If GCONN_USE = True Then
                rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
            Else
                rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
            End If
            
        End If
        
        
        '헤더 설정
        If i = 1 Then Set objGrid = fpTLNI_LNLDGR
        If i = 2 Then Set objGrid = fpTLNE_LDGRADHIST
        If i = 3 Then Set objGrid = fpTLNG_DLY
        If i = 4 Then Set objGrid = fpTLNG_LDGRTRHIST
        If i = 5 Then Set objGrid = fpTLNG_TRAMTCALC
        If i = 6 Then Set objGrid = fpTLNE_ADAMTCALC
        If i = 7 Then Set objGrid = fpTLNS_RFND
        If i = 8 Then Set objGrid = fpTLNS_RFNDTRHIST
        If i = 9 Then Set objGrid = fpTLNS_RFAMTCALC
        If i = 10 Then Set objGrid = fpTLNS_SPCLLDGR
        If i = 11 Then Set objGrid = fpTLNS_SPCLTRHIST
        If i = 12 Then Set objGrid = fpTLNS_SPAMTCALC
        
        
        
        
        If sTableName <> "" Then
        

            objGrid.MaxCols = rsTemp.RecordCount
            
            For j = 1 To rsTemp.RecordCount
                
                
                sField = IIf(IsNull(rsTemp.Fields("COMMENTS").Value), rsTemp.Fields("COLUMN_NAME"), rsTemp.Fields("COMMENTS"))
                
                
                
                objGrid.Row = 0
                objGrid.Col = j
                objGrid.Value = sField
                
                rsTemp.MoveNext
                
            Next j
        End If
        
        Set rsTemp = Nothing
        Set rsTemp = New ADODB.Recordset
        rsTemp.CursorLocation = adUseClient
        

         sSQL = "SELECT * FROM " & sTableName & " WHERE LDGR_NO = '" & Trim(txtLdgrNo.Text) & "' ORDER BY 1, 2, 3, 4"

        If sTableName <> "" Then
            txtSQL.Text = sSQL
            If GCONN_USE = True Then
                rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
            Else
                rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
            End If
            
            
            objGrid.MaxRows = 0
            If rsTemp.RecordCount < 1 Then
                objGrid.MaxRows = 1
            Else
                '조회된 데이타 디스플레이
                Call gfRsToFp(objGrid, rsTemp)
            End If
            
 
            
        End If
        
    Next i
    
    
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


Private Sub txtMapping_GotFocus()
    txtMapping.SelStart = 0
    txtMapping.SelLength = Len(txtMapping.Text)
End Sub


Private Sub txtSelectSQL_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
    If Shift = 2 And KeyCode = 65 Then
        txtSelectSQL.SelStart = 0
        txtSelectSQL.SelLength = Len(txtSelectSQL.Text)
    End If
    
    
End Sub



Private Sub txtSPconect_GotFocus()
    txtSPname.Text = ""
    txtSPconect.Text = ""
End Sub

Private Sub txtSPconect_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call cmdSPSearch_Click
    End If

End Sub

Private Sub txtSPname_GotFocus()
    txtSPname.Text = ""
    txtSPconect.Text = ""
End Sub

Private Sub txtSPname_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call cmdSPSearch_Click
    End If
End Sub

Private Sub txtSQL_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 And KeyCode = 65 Then
        txtSQL.SelStart = 0
        txtSQL.SelLength = Len(txtSQL.Text)
    End If
    
End Sub

Private Sub txtSrchId_()
    
End Sub

Private Sub txtSrchId_GotFocus()
    txtSrchId.SelStart = 0
    txtSrchId.SelLength = Len(txtSrchId.Text)
    
    
End Sub

'컬럼명을 그리드에서 찾기
Private Sub txtSrchId_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim iRow As Integer
    If KeyAscii <> 13 Then
        
        Exit Sub
        
    End If
    Dim sSearchKey As String
    
'엑셀에서 복사한 경우 enter key 제거
txtSrchId.Text = Replace(txtSrchId.Text, vbCrLf, "")

    sSearchKey = Trim(UCase(txtSrchId.Text))
    
    
    iRow = fpCol.SearchCol(1, 1, fpCol.MaxRows, sSearchKey, SearchFlagsValue)
    If iRow > 0 Then
        Call fpCol.ShowCell(1, IIf(iRow - 3 <= 1, 1, iRow - 3), PositionUpperCenter)
        fpCol.Col = -1
        fpCol.Row = iRow
        If optColor(0).Value = True Then fpCol.BackColor = S_CYAN
        If optColor(1).Value = True Then fpCol.BackColor = S_YELLOW
        If optColor(2).Value = True Then fpCol.BackColor = S_BROWN
        Exit Sub
    End If
    
    iRow = fpCol.SearchCol(2, 1, fpCol.MaxRows, sSearchKey, SearchFlagsValue)
    If iRow > 0 Then
        
        Call fpCol.ShowCell(1, IIf(iRow - 3 <= 1, 1, iRow - 3), PositionUpperCenter)
        fpCol.Col = -1
        fpCol.Row = iRow
        If optColor(0).Value = True Then fpCol.BackColor = S_CYAN
        If optColor(1).Value = True Then fpCol.BackColor = S_YELLOW
        If optColor(2).Value = True Then fpCol.BackColor = S_BROWN
        Exit Sub
    End If

    
    
   
    Dim iMatchCount As Integer
    Dim iFirstRow As Integer
    

    iMatchCount = 0
    iFirstRow = 0
    
    For i = 1 To fpCol.MaxRows
        fpCol.Row = i
        fpCol.Col = 1
        
        If InStr(1, Trim(fpCol.Value), sSearchKey) > 0 Then
        
            iMatchCount = iMatchCount + 1
            
            iRow = i
            
            If iFirstRow = 0 Then
                iFirstRow = iRow
            End If
            

            fpCol.Col = -1
            fpCol.Row = iRow
            fpCol.BackColor = S_LIGHT_GRAY
        
        End If
        
        
        '_ 없이 찾기
        If InStr(1, Replace(Trim(fpCol.Value), "_", ""), sSearchKey) > 0 Then
        
            iMatchCount = iMatchCount + 1
            
            iRow = i
            
            If iFirstRow = 0 Then
                iFirstRow = iRow
            End If
            

            fpCol.Col = -1
            fpCol.Row = iRow
            fpCol.BackColor = S_LIGHT_GRAY
        
        End If
        
        
        
        
        
        
        
        
        
        fpCol.Col = 2

        If InStr(1, Trim(fpCol.Value), sSearchKey) > 0 Then
            
            iMatchCount = iMatchCount + 1
            
            iRow = i
            
            
            If iFirstRow = 0 Then
                iFirstRow = iRow
            End If
            
            
            
            fpCol.Col = -1
            fpCol.Row = iRow
            fpCol.BackColor = S_LIGHT_GRAY
        
        End If
    
    Next i
    
    

    Call fpCol.ShowCell(1, IIf(iFirstRow - 3 <= 1, 1, iFirstRow - 3), PositionUpperCenter)

    
    If iMatchCount > 1 Then
        lblMatch.Caption = "Matched " & iMatchCount & " columns...."
        picMatch.Visible = True
    End If
    Timer1.Tag = 0
    Timer1.Enabled = True
End Sub

Private Sub txtTabMapping_GotFocus()
    txtTabMapping.SelStart = 0
    txtTabMapping.SelLength = Len(txtTabMapping.Text)
End Sub

Private Sub txUpdateSQL_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 And KeyCode = 65 Then
        txUpdateSQL.SelStart = 0
        txUpdateSQL.SelLength = Len(txUpdateSQL.Text)
    End If
End Sub





Private Sub txtTableSrchBigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call cmdTableSrch_Click
    End If
End Sub


Private Sub txtReqSQL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF5
            Call cmdFreeSQL_Click
    End Select
End Sub

Private Sub txtTableSrchEng_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call cmdTableSrch_Click
    End If
End Sub

Private Sub txtTableSrchKor_GotFocus()
    txtTableSrchEng.Text = ""
    txtTableSrchBigo.Text = ""
End Sub

Private Sub txtTableSrchEng_GotFocus()
    txtTableSrchKor.Text = ""
    txtTableSrchBigo.Text = ""
End Sub

Private Sub txtTableSrchBigo_GotFocus()
    txtTableSrchKor.Text = ""
    txtTableSrchEng.Text = ""
End Sub


Private Sub txtTableSrchKor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call cmdTableSrch_Click
    End If
End Sub





Private Sub makeCellBorderline()



    'Determines the section of the cell border displayed around the entire spreadsheet
    fpSample.SetCellBorder 1, 1, fpSample.MaxCols, fpSample.MaxRows, SS_BORDER_TYPE_LEFT, &HC0C0C0, SS_BORDER_STYLE_SOLID
    
    fpSample.SetCellBorder 1, 1, fpSample.MaxCols, fpSample.MaxRows, SS_BORDER_TYPE_TOP, &HC0C0C0, SS_BORDER_STYLE_SOLID
    fpSample.SetCellBorder 1, 1, fpSample.MaxCols, fpSample.MaxRows, SS_BORDER_TYPE_RIGHT, &HC0C0C0, SS_BORDER_STYLE_SOLID
    fpSample.SetCellBorder 1, 1, fpSample.MaxCols, fpSample.MaxRows, SS_BORDER_TYPE_BOTTOM, &HC0C0C0, SS_BORDER_STYLE_SOLID

    'Determines the section of the cell border displayed around the entire spreadsheet
    fpSample.SetCellBorder 1, 1, fpSample.MaxCols, fpSample.MaxRows, SS_BORDER_TYPE_LEFT, &HC0C0C0, SS_BORDER_STYLE_SOLID
    fpSample.SetCellBorder 1, 1, fpSample.MaxCols, fpSample.MaxRows, SS_BORDER_TYPE_TOP, &HC0C0C0, SS_BORDER_STYLE_SOLID
    fpSample.SetCellBorder 1, 1, fpSample.MaxCols, fpSample.MaxRows, SS_BORDER_TYPE_RIGHT, &HC0C0C0, SS_BORDER_STYLE_SOLID
    fpSample.SetCellBorder 1, 1, fpSample.MaxCols, fpSample.MaxRows, SS_BORDER_TYPE_BOTTOM, &HC0C0C0, SS_BORDER_STYLE_SOLID

    'Determines the section of the cell border displayed around the entire spreadsheet
    fpSample.SetCellBorder 1, 1, fpSample.MaxCols, fpSample.MaxRows, SS_BORDER_TYPE_LEFT, &HC0C0C0, SS_BORDER_STYLE_SOLID
    fpSample.SetCellBorder 1, 1, fpSample.MaxCols, fpSample.MaxRows, SS_BORDER_TYPE_TOP, &HC0C0C0, SS_BORDER_STYLE_SOLID
    fpSample.SetCellBorder 1, 1, fpSample.MaxCols, fpSample.MaxRows, SS_BORDER_TYPE_RIGHT, &HC0C0C0, SS_BORDER_STYLE_SOLID
    fpSample.SetCellBorder 1, 1, fpSample.MaxCols, fpSample.MaxRows, SS_BORDER_TYPE_BOTTOM, &HC0C0C0, SS_BORDER_STYLE_SOLID

    'Determines the section of the cell border displayed around the entire spreadsheet
    fpSample.SetCellBorder 1, 1, fpSample.MaxCols, fpSample.MaxRows, SS_BORDER_TYPE_LEFT, &HC0C0C0, SS_BORDER_STYLE_SOLID
    fpSample.SetCellBorder 1, 1, fpSample.MaxCols, fpSample.MaxRows, SS_BORDER_TYPE_TOP, &HC0C0C0, SS_BORDER_STYLE_SOLID
    fpSample.SetCellBorder 1, 1, fpSample.MaxCols, fpSample.MaxRows, SS_BORDER_TYPE_RIGHT, &HC0C0C0, SS_BORDER_STYLE_SOLID
    fpSample.SetCellBorder 1, 1, fpSample.MaxCols, fpSample.MaxRows, SS_BORDER_TYPE_BOTTOM, &HC0C0C0, SS_BORDER_STYLE_SOLID

    'Determines the section of the cell border displayed around the entire spreadsheet
    fpSample.SetCellBorder 1, 1, fpSample.MaxCols, fpSample.MaxRows, SS_BORDER_TYPE_LEFT, &HC0C0C0, SS_BORDER_STYLE_SOLID
    fpSample.SetCellBorder 1, 1, fpSample.MaxCols, fpSample.MaxRows, SS_BORDER_TYPE_TOP, &HC0C0C0, SS_BORDER_STYLE_SOLID
    fpSample.SetCellBorder 1, 1, fpSample.MaxCols, fpSample.MaxRows, SS_BORDER_TYPE_RIGHT, &HC0C0C0, SS_BORDER_STYLE_SOLID
    fpSample.SetCellBorder 1, 1, fpSample.MaxCols, fpSample.MaxRows, SS_BORDER_TYPE_BOTTOM, &HC0C0C0, SS_BORDER_STYLE_SOLID



End Sub


