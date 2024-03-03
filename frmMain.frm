VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "TabMan"
   ClientHeight    =   15240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15360
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   15240
   ScaleWidth      =   15360
   Begin TabDlg.SSTab SSTab1 
      Height          =   11115
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   19606
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   16777215
      TabCaption(0)   =   "Home"
      TabPicture(0)   =   "frmMain.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Command1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "EZFTP1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "SSTab4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Table Manager"
      TabPicture(1)   =   "frmMain.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(2)=   "fpTable"
      Tab(1).Control(3)=   "SSTab3"
      Tab(1).Control(4)=   "SSTab2"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "작업관리"
      TabPicture(2)   =   "frmMain.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(1)=   "SSTab5"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "ETC"
      TabPicture(3)   =   "frmMain.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtSQL"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame5 
         Height          =   735
         Left            =   -72360
         TabIndex        =   122
         Top             =   300
         Width           =   3615
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
            Left            =   120
            TabIndex        =   128
            Top             =   390
            Width           =   1455
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
            Left            =   1800
            TabIndex        =   127
            Top             =   390
            Width           =   1515
         End
         Begin VB.CommandButton Command8 
            Caption         =   "1D"
            Height          =   255
            Left            =   120
            TabIndex        =   126
            Top             =   120
            Width           =   435
         End
         Begin VB.CommandButton Command9 
            Caption         =   "1W"
            Height          =   255
            Left            =   600
            TabIndex        =   125
            Top             =   120
            Width           =   435
         End
         Begin VB.CommandButton Command10 
            Caption         =   "1M"
            Height          =   255
            Left            =   1080
            TabIndex        =   124
            Top             =   120
            Width           =   435
         End
         Begin VB.CommandButton Command12 
            Caption         =   "TODAY"
            Height          =   255
            Left            =   2400
            TabIndex        =   123
            Top             =   120
            Width           =   795
         End
      End
      Begin TabDlg.SSTab SSTab5 
         Height          =   10155
         Left            =   -74880
         TabIndex        =   79
         Top             =   840
         Width           =   14955
         _ExtentX        =   26379
         _ExtentY        =   17912
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "DEBUG LOG"
         TabPicture(0)   =   "frmMain.frx":037A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fpDebugLog"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "cmdDebugLog"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "txtPgmIDLog"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "ssTabeLog"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "JOB LOG"
         TabPicture(1)   =   "frmMain.frx":0396
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Command11"
         Tab(1).Control(1)=   "txtMessage"
         Tab(1).Control(2)=   "cboStat"
         Tab(1).Control(3)=   "txtPgmIDLog2"
         Tab(1).Control(4)=   "cmdJobLog"
         Tab(1).Control(5)=   "Command7"
         Tab(1).Control(6)=   "fpJobLog"
         Tab(1).Control(7)=   "Label16"
         Tab(1).Control(8)=   "Label15"
         Tab(1).ControlCount=   9
         Begin TabDlg.SSTab ssTabeLog 
            Height          =   4815
            Left            =   60
            TabIndex        =   119
            Top             =   5400
            Width           =   14775
            _ExtentX        =   26061
            _ExtentY        =   8493
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabHeight       =   520
            TabCaption(0)   =   "로그이력"
            TabPicture(0)   =   "frmMain.frx":03B2
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "fpDebugLogHist"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "과거로그 내역"
            TabPicture(1)   =   "frmMain.frx":03CE
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "fpDebugLogHistDtail"
            Tab(1).ControlCount=   1
            Begin FPSpreadADO.fpSpread fpDebugLogHist 
               Height          =   4275
               Left            =   120
               TabIndex        =   120
               Top             =   420
               Width           =   14415
               _Version        =   393216
               _ExtentX        =   25426
               _ExtentY        =   7541
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
               GrayAreaBackColor=   12640511
               MaxCols         =   4
               MaxRows         =   1
               RetainSelBlock  =   0   'False
               ScrollBarExtMode=   -1  'True
               ShadowColor     =   12648447
               SpreadDesigner  =   "frmMain.frx":03EA
               ScrollBarTrack  =   1
            End
            Begin FPSpreadADO.fpSpread fpDebugLogHistDtail 
               Height          =   4455
               Left            =   -74820
               TabIndex        =   121
               Top             =   420
               Width           =   14415
               _Version        =   393216
               _ExtentX        =   25426
               _ExtentY        =   7858
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
               GrayAreaBackColor=   12640511
               MaxCols         =   7
               MaxRows         =   1
               RetainSelBlock  =   0   'False
               ScrollBarExtMode=   -1  'True
               ShadowColor     =   12648447
               SpreadDesigner  =   "frmMain.frx":0886
               ScrollBarTrack  =   1
            End
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Clear"
            Height          =   375
            Left            =   -62040
            TabIndex        =   92
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtMessage 
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
            Left            =   -63780
            TabIndex        =   90
            Top             =   510
            Width           =   1515
         End
         Begin VB.ComboBox cboStat 
            Height          =   300
            ItemData        =   "frmMain.frx":0E2D
            Left            =   -69660
            List            =   "frmMain.frx":0E3D
            TabIndex        =   89
            Text            =   "Y"
            Top             =   570
            Width           =   675
         End
         Begin VB.TextBox txtPgmIDLog2 
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
            Left            =   -74760
            TabIndex        =   86
            Top             =   570
            Width           =   3315
         End
         Begin VB.CommandButton cmdJobLog 
            Caption         =   "조회"
            Height          =   315
            Left            =   -71220
            Style           =   1  '그래픽
            TabIndex        =   85
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton Command7 
            Caption         =   "조회"
            Height          =   495
            Left            =   -55980
            Style           =   1  '그래픽
            TabIndex        =   84
            Top             =   13080
            Width           =   855
         End
         Begin VB.TextBox txtPgmIDLog 
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
            Left            =   300
            TabIndex        =   82
            Top             =   360
            Width           =   3315
         End
         Begin VB.CommandButton cmdDebugLog 
            Caption         =   "조회"
            Height          =   315
            Left            =   3660
            Style           =   1  '그래픽
            TabIndex        =   81
            Top             =   360
            Width           =   855
         End
         Begin FPSpreadADO.fpSpread fpDebugLog 
            Height          =   4635
            Left            =   300
            TabIndex        =   80
            Top             =   720
            Width           =   14415
            _Version        =   393216
            _ExtentX        =   25426
            _ExtentY        =   8176
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
            GrayAreaBackColor=   12640511
            MaxCols         =   6
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            ShadowColor     =   12648447
            SpreadDesigner  =   "frmMain.frx":0E51
            ScrollBarTrack  =   1
         End
         Begin FPSpreadADO.fpSpread fpJobLog 
            Height          =   8955
            Left            =   -74880
            TabIndex        =   83
            Top             =   1080
            Width           =   14775
            _Version        =   393216
            _ExtentX        =   26061
            _ExtentY        =   15796
            _StockProps     =   64
            ColsFrozen      =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "굴림체"
               Size            =   9
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   12640511
            MaxCols         =   14
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            ShadowColor     =   12648447
            SpreadDesigner  =   "frmMain.frx":13BC
            ScrollBarTrack  =   1
         End
         Begin VB.Label Label16 
            Caption         =   "메세지"
            Height          =   255
            Left            =   -64500
            TabIndex        =   91
            Top             =   600
            Width           =   555
         End
         Begin VB.Label Label15 
            Caption         =   "상태"
            Height          =   255
            Left            =   -70140
            TabIndex        =   88
            Top             =   600
            Width           =   555
         End
      End
      Begin VB.TextBox txtSQL 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9255
         IMEMode         =   8  '영문
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  '수직
         TabIndex        =   49
         Top             =   840
         Width           =   13695
      End
      Begin TabDlg.SSTab SSTab4 
         Height          =   10695
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   14955
         _ExtentX        =   26379
         _ExtentY        =   18865
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "직원정보 조회"
         TabPicture(0)   =   "frmMain.frx":1C88
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "WebBrowser1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "cmdIBKSGroup"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdIBKsystem"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "cmdFineNetemp"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cmdFindDept"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "txtFindKey"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "cmdFindEmp"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "EDW 메타조회"
         TabPicture(1)   =   "frmMain.frx":1CA4
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "WebBrowser2"
         Tab(1).Control(1)=   "cmdEDW"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "DB2 InfoCenter"
         TabPicture(2)   =   "frmMain.frx":1CC0
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "WebBrowser3"
         Tab(2).Control(1)=   "cmdInfoCenter"
         Tab(2).Control(2)=   "cmdOnline"
         Tab(2).ControlCount=   3
         TabCaption(3)   =   "DB설정"
         TabPicture(3)   =   "frmMain.frx":1CDC
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "txtDBA"
         Tab(3).Control(1)=   "txtDBApass"
         Tab(3).Control(2)=   "cmdConnChange"
         Tab(3).Control(3)=   "txtPWD"
         Tab(3).Control(4)=   "txtID"
         Tab(3).Control(5)=   "txtDSN"
         Tab(3).Control(6)=   "txtProvider"
         Tab(3).Control(7)=   "Label12"
         Tab(3).Control(8)=   "Label11"
         Tab(3).Control(9)=   "Label10"
         Tab(3).Control(10)=   "Label9"
         Tab(3).Control(11)=   "Label8"
         Tab(3).Control(12)=   "Label7"
         Tab(3).ControlCount=   13
         Begin VB.CommandButton cmdOnline 
            Caption         =   "Online Info Center"
            Height          =   435
            Left            =   -71220
            TabIndex        =   87
            Top             =   420
            Width           =   2535
         End
         Begin VB.TextBox txtDBA 
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
            Left            =   -73200
            TabIndex        =   72
            Top             =   2640
            Width           =   2355
         End
         Begin VB.TextBox txtDBApass 
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
            Left            =   -73200
            TabIndex        =   71
            Top             =   3060
            Width           =   2355
         End
         Begin VB.CommandButton cmdConnChange 
            Caption         =   "변경"
            Height          =   375
            Left            =   -72840
            Style           =   1  '그래픽
            TabIndex        =   70
            Top             =   3780
            Width           =   1335
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
            IMEMode         =   10  '한글 
            Left            =   -73200
            TabIndex        =   68
            Top             =   2220
            Width           =   2355
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
            IMEMode         =   10  '한글 
            Left            =   -73200
            TabIndex        =   66
            Top             =   1800
            Width           =   2355
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
            IMEMode         =   10  '한글 
            Left            =   -73200
            TabIndex        =   64
            Top             =   1380
            Width           =   2355
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
            IMEMode         =   10  '한글 
            Left            =   -73200
            TabIndex        =   62
            Top             =   1020
            Width           =   2355
         End
         Begin VB.CommandButton cmdInfoCenter 
            Caption         =   "ReFresh"
            Height          =   495
            Left            =   -74640
            Style           =   1  '그래픽
            TabIndex        =   45
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton cmdEDW 
            Caption         =   "ReFresh"
            Height          =   375
            Left            =   -74760
            Style           =   1  '그래픽
            TabIndex        =   44
            Top             =   360
            Width           =   1755
         End
         Begin VB.CommandButton cmdFindEmp 
            Caption         =   "직원조회"
            Height          =   375
            Left            =   2640
            Style           =   1  '그래픽
            TabIndex        =   41
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtFindKey 
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
            Left            =   240
            TabIndex        =   40
            Top             =   600
            Width           =   2355
         End
         Begin VB.CommandButton cmdFindDept 
            Caption         =   "부서조회"
            Height          =   375
            Left            =   5820
            Style           =   1  '그래픽
            TabIndex        =   39
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton cmdFineNetemp 
            Caption         =   "직원찾기홈"
            Height          =   375
            Left            =   7260
            Style           =   1  '그래픽
            TabIndex        =   38
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton cmdIBKsystem 
            Caption         =   "IBKS직원찾기"
            Height          =   375
            Left            =   4140
            Style           =   1  '그래픽
            TabIndex        =   37
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton cmdIBKSGroup 
            Caption         =   "그룹웨어"
            Height          =   375
            Left            =   8940
            Style           =   1  '그래픽
            TabIndex        =   36
            Top             =   600
            Width           =   1335
         End
         Begin SHDocVwCtl.WebBrowser WebBrowser1 
            Height          =   9435
            Left            =   60
            TabIndex        =   35
            Top             =   1080
            Width           =   14775
            ExtentX         =   26061
            ExtentY         =   16642
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
            Height          =   9795
            Left            =   -74880
            TabIndex        =   42
            Top             =   780
            Width           =   14715
            ExtentX         =   25956
            ExtentY         =   17277
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
            Height          =   9675
            Left            =   -74940
            TabIndex        =   43
            Top             =   900
            Width           =   14775
            ExtentX         =   26061
            ExtentY         =   17066
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
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "DBA USER"
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
            Left            =   -74355
            TabIndex        =   74
            Top             =   2760
            Width           =   975
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "DBA USER PASS"
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
            TabIndex        =   73
            Top             =   3120
            Width           =   1560
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
            Left            =   -73830
            TabIndex        =   69
            Top             =   2340
            Width           =   450
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
            Left            =   -73560
            TabIndex        =   67
            Top             =   1920
            Width           =   180
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
            Left            =   -73785
            TabIndex        =   65
            Top             =   1500
            Width           =   405
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
            Left            =   -74160
            TabIndex        =   63
            Top             =   1080
            Width           =   780
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ftptest"
         Height          =   435
         Left            =   12360
         TabIndex        =   28
         Top             =   480
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.PictureBox EZFTP1 
         Height          =   480
         Left            =   13980
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   129
         Top             =   480
         Width           =   1200
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   13080
         TabIndex        =   25
         Top             =   420
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   1275
         Left            =   -69780
         TabIndex        =   8
         Top             =   7800
         Width           =   10095
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
            Height          =   735
            IMEMode         =   8  '영문
            Left            =   60
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   30
            Top             =   480
            Width           =   8115
         End
         Begin VB.CommandButton cmdValueSelect 
            Caption         =   "값조회"
            Height          =   495
            Left            =   9180
            Style           =   1  '그래픽
            TabIndex        =   21
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton cmdColMultiple 
            Caption         =   "컬럼등록"
            Height          =   555
            Left            =   9180
            Style           =   1  '그래픽
            TabIndex        =   20
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdTableUpdate 
            Caption         =   "수정"
            Height          =   675
            Left            =   8220
            Style           =   1  '그래픽
            TabIndex        =   13
            Top             =   540
            Width           =   915
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
            Left            =   7560
            TabIndex        =   12
            Top             =   180
            Width           =   1575
         End
         Begin VB.TextBox edtTablenm 
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
            Left            =   4560
            TabIndex        =   11
            Top             =   180
            Width           =   2955
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
            Left            =   2160
            TabIndex        =   10
            Top             =   180
            Width           =   2415
         End
         Begin VB.ComboBox cboShema 
            Height          =   300
            Left            =   60
            TabIndex        =   9
            Top             =   180
            Width           =   2055
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Height          =   1095
         Left            =   -74880
         TabIndex        =   2
         Top             =   360
         Width           =   4755
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
            IMEMode         =   8  '영문
            Left            =   960
            TabIndex        =   32
            Top             =   780
            Width           =   2355
         End
         Begin VB.CommandButton cmdTableSrch 
            Caption         =   "조회"
            Height          =   915
            Left            =   3360
            Style           =   1  '그래픽
            TabIndex        =   7
            Top             =   180
            Width           =   1335
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
            Left            =   960
            TabIndex        =   6
            Top             =   480
            Width           =   2355
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
            Left            =   960
            TabIndex        =   5
            Top             =   180
            Width           =   2355
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
            TabIndex        =   33
            Top             =   840
            Width           =   390
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "테이블ID"
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
            TabIndex        =   4
            Top             =   540
            Width           =   765
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
            Left            =   60
            TabIndex        =   3
            Top             =   240
            Width           =   780
         End
      End
      Begin FPSpreadADO.fpSpread fpTable 
         Height          =   7275
         Left            =   -74880
         TabIndex        =   1
         Top             =   1500
         Width           =   4755
         _Version        =   393216
         _ExtentX        =   8387
         _ExtentY        =   12832
         _StockProps     =   64
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
         GrayAreaBackColor=   12640511
         MaxCols         =   9
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         ShadowColor     =   12648447
         ShadowDark      =   12632256
         SpreadDesigner  =   "frmMain.frx":1CF8
         ScrollBarTrack  =   1
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   7155
         Left            =   -70080
         TabIndex        =   14
         Top             =   540
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   12621
         _Version        =   393216
         Style           =   1
         Tabs            =   7
         TabsPerRow      =   7
         TabHeight       =   520
         BackColor       =   16777215
         TabCaption(0)   =   "테이블정보"
         TabPicture(0)   =   "frmMain.frx":2329
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fpCol"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtColNm"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "txtColId"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "txtCodeDomain"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cmdColUpdate"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "frameCode"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "txtColBigo"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "SQL문 실행"
         TabPicture(1)   =   "frmMain.frx":2345
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtReqSQL"
         Tab(1).Control(1)=   "cmdFreeSQL"
         Tab(1).Control(2)=   "Label3"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "INSERT SQL"
         TabPicture(2)   =   "frmMain.frx":2361
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtAlias"
         Tab(2).Control(1)=   "Command6"
         Tab(2).Control(2)=   "txtGrpBySQL"
         Tab(2).Control(3)=   "txtSelectSQL"
         Tab(2).Control(4)=   "txUpdateSQL"
         Tab(2).Control(5)=   "txtInsertSQL"
         Tab(2).Control(6)=   "Label4"
         Tab(2).ControlCount=   7
         TabCaption(3)   =   "상세정보"
         TabPicture(3)   =   "frmMain.frx":237D
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame4"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Tab 4"
         TabPicture(4)   =   "frmMain.frx":2399
         Tab(4).ControlEnabled=   0   'False
         Tab(4).ControlCount=   0
         TabCaption(5)   =   "기타정보"
         TabPicture(5)   =   "frmMain.frx":23B5
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Frame3"
         Tab(5).Control(1)=   "txtGrpByCol"
         Tab(5).Control(2)=   "cmdRecCountSrch"
         Tab(5).Control(3)=   "fpRecCount"
         Tab(5).ControlCount=   4
         TabCaption(6)   =   "Table Space"
         TabPicture(6)   =   "frmMain.frx":23D1
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "cboDanwi"
         Tab(6).Control(1)=   "cmdTableSpaceSrch"
         Tab(6).Control(2)=   "fpTableSPace"
         Tab(6).ControlCount=   3
         Begin VB.Frame Frame4 
            BackColor       =   &H00C2EFDC&
            Height          =   6795
            Left            =   -74940
            TabIndex        =   103
            Top             =   420
            Width           =   10035
            Begin TabDlg.SSTab SSTab6 
               Height          =   2235
               Left            =   3480
               TabIndex        =   116
               Top             =   180
               Width           =   6495
               _ExtentX        =   11456
               _ExtentY        =   3942
               _Version        =   393216
               Style           =   1
               Tabs            =   2
               TabsPerRow      =   2
               TabHeight       =   520
               TabCaption(0)   =   "데이터 가공에 사용된"
               TabPicture(0)   =   "frmMain.frx":23ED
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "fpInTab"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).ControlCount=   1
               TabCaption(1)   =   "나를 사용하는"
               TabPicture(1)   =   "frmMain.frx":2409
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "fpOutTab"
               Tab(1).ControlCount=   1
               Begin FPSpreadADO.fpSpread fpInTab 
                  Height          =   1815
                  Left            =   60
                  TabIndex        =   117
                  Top             =   360
                  Width           =   6375
                  _Version        =   393216
                  _ExtentX        =   11245
                  _ExtentY        =   3201
                  _StockProps     =   64
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
                  GrayAreaBackColor=   12640511
                  MaxCols         =   2
                  MaxRows         =   1
                  RetainSelBlock  =   0   'False
                  ScrollBarExtMode=   -1  'True
                  SelectBlockOptions=   0
                  ShadowColor     =   12648447
                  ShadowDark      =   12632256
                  SpreadDesigner  =   "frmMain.frx":2425
                  ScrollBarTrack  =   1
               End
               Begin FPSpreadADO.fpSpread fpOutTab 
                  Height          =   1815
                  Left            =   -74940
                  TabIndex        =   118
                  Top             =   360
                  Width           =   6375
                  _Version        =   393216
                  _ExtentX        =   11245
                  _ExtentY        =   3201
                  _StockProps     =   64
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
                  MaxCols         =   2
                  MaxRows         =   1
                  RetainSelBlock  =   0   'False
                  ScrollBarExtMode=   -1  'True
                  SelectBlockOptions=   0
                  ShadowColor     =   14737632
                  ShadowDark      =   12632256
                  SpreadDesigner  =   "frmMain.frx":2816
                  ScrollBarTrack  =   1
               End
            End
            Begin VB.CommandButton cmdAddRow 
               Caption         =   "+"
               Height          =   375
               Left            =   2940
               Style           =   1  '그래픽
               TabIndex        =   115
               Top             =   2040
               Width           =   495
            End
            Begin VB.CommandButton cmdDescReg 
               Caption         =   "등록"
               Height          =   495
               Left            =   60
               Style           =   1  '그래픽
               TabIndex        =   114
               Top             =   180
               Width           =   3315
            End
            Begin VB.TextBox txtCR_LOG 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   12
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2295
               IMEMode         =   8  '영문
               Left            =   60
               MultiLine       =   -1  'True
               ScrollBars      =   2  '수직
               TabIndex        =   112
               Top             =   4440
               Width           =   9855
            End
            Begin VB.TextBox txtIMP_DESC 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   12
                  Charset         =   129
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1695
               IMEMode         =   8  '영문
               Left            =   60
               MultiLine       =   -1  'True
               ScrollBars      =   2  '수직
               TabIndex        =   110
               Top             =   2460
               Width           =   9855
            End
            Begin VB.TextBox txtUSE_YN 
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
               TabIndex        =   108
               Top             =   1440
               Width           =   2355
            End
            Begin VB.TextBox txtCYCLE 
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
               TabIndex        =   106
               Top             =   1080
               Width           =   2355
            End
            Begin VB.TextBox txtORGTAB 
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
               TabIndex        =   104
               Top             =   720
               Width           =   2355
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "생성로직"
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
               TabIndex        =   113
               Top             =   4200
               Width           =   780
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "주요사항"
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
               TabIndex        =   111
               Top             =   2040
               Width           =   780
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "사용여부"
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
               TabIndex        =   109
               Top             =   1500
               Width           =   780
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "갱신주기"
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
               TabIndex        =   107
               Top             =   1140
               Width           =   780
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  '투명
               Caption         =   "원천Table"
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
               TabIndex        =   105
               Top             =   780
               Width           =   915
            End
         End
         Begin VB.Frame Frame3 
            Height          =   6435
            Left            =   -70980
            TabIndex        =   99
            Top             =   480
            Visible         =   0   'False
            Width           =   5775
            Begin VB.CommandButton cmdRecCountSrch2 
               Caption         =   "조회"
               Height          =   435
               Left            =   420
               Style           =   1  '그래픽
               TabIndex        =   101
               Top             =   540
               Width           =   3555
            End
            Begin VB.TextBox txtGrpByCol2 
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
               Left            =   420
               TabIndex        =   100
               Top             =   180
               Width           =   3615
            End
            Begin FPSpreadADO.fpSpread fpRecCount2 
               Height          =   5295
               Left            =   480
               TabIndex        =   102
               Top             =   1080
               Width           =   3615
               _Version        =   393216
               _ExtentX        =   6376
               _ExtentY        =   9340
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
               GrayAreaBackColor=   14737632
               MaxCols         =   1
               MaxRows         =   1
               RetainSelBlock  =   0   'False
               ScrollBarExtMode=   -1  'True
               SelectBlockOptions=   0
               SpreadDesigner  =   "frmMain.frx":2C07
               ScrollBarTrack  =   1
            End
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
            Left            =   -66960
            TabIndex        =   97
            Top             =   6735
            Width           =   495
         End
         Begin VB.CommandButton Command6 
            Caption         =   "SQL 재생성"
            Height          =   375
            Left            =   -66420
            TabIndex        =   96
            Top             =   6675
            Width           =   1395
         End
         Begin VB.TextBox txtGrpBySQL 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   12
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1155
            IMEMode         =   8  '영문
            Left            =   -74940
            MultiLine       =   -1  'True
            ScrollBars      =   3  '양방향
            TabIndex        =   95
            Top             =   5460
            Width           =   9855
         End
         Begin VB.TextBox txtSelectSQL 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   12
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2475
            IMEMode         =   8  '영문
            Left            =   -74940
            MultiLine       =   -1  'True
            ScrollBars      =   3  '양방향
            TabIndex        =   94
            Top             =   2940
            Width           =   9855
         End
         Begin VB.TextBox txUpdateSQL 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   12
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            IMEMode         =   8  '영문
            Left            =   -74940
            MultiLine       =   -1  'True
            ScrollBars      =   3  '양방향
            TabIndex        =   93
            Top             =   1680
            Width           =   9855
         End
         Begin VB.ComboBox cboDanwi 
            Height          =   300
            ItemData        =   "frmMain.frx":2F9F
            Left            =   -74880
            List            =   "frmMain.frx":2FA9
            TabIndex        =   51
            Text            =   "MB"
            Top             =   540
            Width           =   1635
         End
         Begin VB.CommandButton cmdTableSpaceSrch 
            Caption         =   "조회"
            Height          =   435
            Left            =   -72780
            Style           =   1  '그래픽
            TabIndex        =   50
            Top             =   420
            Width           =   1695
         End
         Begin VB.TextBox txtGrpByCol 
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
            Left            =   -74760
            TabIndex        =   48
            Top             =   420
            Width           =   3615
         End
         Begin VB.CommandButton cmdRecCountSrch 
            Caption         =   "조회"
            Height          =   435
            Left            =   -74760
            Style           =   1  '그래픽
            TabIndex        =   47
            Top             =   780
            Width           =   3555
         End
         Begin VB.TextBox txtInsertSQL 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   12
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            IMEMode         =   8  '영문
            Left            =   -74940
            MultiLine       =   -1  'True
            ScrollBars      =   3  '양방향
            TabIndex        =   31
            Top             =   460
            Width           =   9855
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
            Height          =   855
            IMEMode         =   8  '영문
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   29
            Top             =   6280
            Width           =   9795
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
            Height          =   6075
            IMEMode         =   8  '영문
            Left            =   -74940
            MultiLine       =   -1  'True
            ScrollBars      =   2  '수직
            TabIndex        =   26
            Top             =   1000
            Width           =   9975
         End
         Begin VB.CommandButton cmdFreeSQL 
            Caption         =   "조회"
            Height          =   495
            Left            =   -65940
            Style           =   1  '그래픽
            TabIndex        =   24
            Top             =   460
            Width           =   855
         End
         Begin VB.Frame frameCode 
            BackColor       =   &H00FFFFFF&
            Height          =   4635
            Left            =   240
            TabIndex        =   22
            Top             =   1440
            Width           =   9315
            Begin FPSpreadADO.fpSpread fpcode 
               Height          =   3855
               Left            =   60
               TabIndex        =   23
               Top             =   540
               Width           =   9135
               _Version        =   393216
               _ExtentX        =   16113
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
               MaxCols         =   5
               MaxRows         =   1
               RetainSelBlock  =   0   'False
               ScrollBarExtMode=   -1  'True
               SelectBlockOptions=   0
               SpreadDesigner  =   "frmMain.frx":2FB5
               ScrollBarTrack  =   1
            End
            Begin VB.Image Image1 
               Height          =   405
               Left            =   8580
               Picture         =   "frmMain.frx":346D
               Stretch         =   -1  'True
               Top             =   120
               Width           =   510
            End
         End
         Begin VB.CommandButton cmdColUpdate 
            Caption         =   "수정"
            Height          =   310
            Left            =   8160
            Style           =   1  '그래픽
            TabIndex        =   19
            Top             =   5980
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
            Left            =   6840
            TabIndex        =   18
            Top             =   5980
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
            Left            =   120
            TabIndex        =   17
            Top             =   5980
            Width           =   3135
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
            Left            =   3300
            TabIndex        =   16
            Top             =   5980
            Width           =   3495
         End
         Begin FPSpreadADO.fpSpread fpCol 
            Height          =   5535
            Left            =   420
            TabIndex        =   15
            Top             =   480
            Width           =   9795
            _Version        =   393216
            _ExtentX        =   17277
            _ExtentY        =   9763
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
            GrayAreaBackColor=   12640511
            MaxCols         =   9
            MaxRows         =   50
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            ShadowColor     =   12648447
            SpreadDesigner  =   "frmMain.frx":5E35
            ScrollBarTrack  =   1
         End
         Begin FPSpreadADO.fpSpread fpRecCount 
            Height          =   5595
            Left            =   -74760
            TabIndex        =   46
            Top             =   1320
            Width           =   3615
            _Version        =   393216
            _ExtentX        =   6376
            _ExtentY        =   9869
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
            GrayAreaBackColor=   14737632
            MaxCols         =   2
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            SelectBlockOptions=   0
            SpreadDesigner  =   "frmMain.frx":666E
            ScrollBarTrack  =   1
         End
         Begin FPSpreadADO.fpSpread fpTableSPace 
            Height          =   5775
            Left            =   -74940
            TabIndex        =   52
            Top             =   1140
            Width           =   9915
            _Version        =   393216
            _ExtentX        =   17489
            _ExtentY        =   10186
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
            MaxCols         =   7
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            SelectBlockOptions=   0
            SpreadDesigner  =   "frmMain.frx":6A83
            ScrollBarTrack  =   1
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
            Left            =   -67500
            TabIndex        =   98
            Top             =   6795
            Width           =   450
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
            Left            =   -74820
            TabIndex        =   27
            Top             =   700
            Width           =   1035
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   2265
         Left            =   -74940
         TabIndex        =   53
         Top             =   8820
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   3995
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         Tab             =   2
         TabsPerRow      =   7
         TabHeight       =   520
         TabCaption(0)   =   "sample Data"
         TabPicture(0)   =   "frmMain.frx":7092
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "fpSample"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "INDEX"
         TabPicture(1)   =   "frmMain.frx":70AE
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fpIndex"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Table 상태"
         TabPicture(2)   =   "frmMain.frx":70CA
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "Label5"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Label13"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Label14"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "fpTableStat"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "cboDanwi2"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "txtRowCnt"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "txtTBSpace"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "txtIndexTbspace"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).ControlCount=   8
         TabCaption(3)   =   "Partition"
         TabPicture(3)   =   "frmMain.frx":70E6
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "ftPartition"
         Tab(3).Control(1)=   "cmdPartitionSearch"
         Tab(3).ControlCount=   2
         Begin VB.TextBox txtIndexTbspace 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
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
            Left            =   10800
            TabIndex        =   77
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox txtTBSpace 
            Alignment       =   2  '가운데 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
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
            Left            =   7080
            TabIndex        =   75
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton cmdPartitionSearch 
            Caption         =   "조회"
            Height          =   1455
            Left            =   -74940
            Style           =   1  '그래픽
            TabIndex        =   60
            Top             =   540
            Width           =   2835
         End
         Begin VB.TextBox txtRowCnt 
            Alignment       =   1  '오른쪽 맞춤
            Appearance      =   0  '평면
            BackColor       =   &H00E0E0E0&
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
            Left            =   3060
            TabIndex        =   57
            Top             =   360
            Width           =   2595
         End
         Begin VB.ComboBox cboDanwi2 
            Height          =   300
            ItemData        =   "frmMain.frx":7102
            Left            =   240
            List            =   "frmMain.frx":710F
            TabIndex        =   56
            Text            =   "MB"
            Top             =   420
            Width           =   1635
         End
         Begin FPSpreadADO.fpSpread fpIndex 
            Height          =   1755
            Left            =   -74940
            TabIndex        =   54
            Top             =   400
            Width           =   14415
            _Version        =   393216
            _ExtentX        =   25426
            _ExtentY        =   3096
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
            MaxCols         =   3
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            SelectBlockOptions=   0
            SpreadDesigner  =   "frmMain.frx":711F
            ScrollBarTrack  =   1
         End
         Begin FPSpreadADO.fpSpread fpSample 
            Height          =   1815
            Left            =   -75000
            TabIndex        =   55
            Top             =   400
            Width           =   14715
            _Version        =   393216
            _ExtentX        =   25956
            _ExtentY        =   3201
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
            SpreadDesigner  =   "frmMain.frx":74F6
            ScrollBarTrack  =   1
         End
         Begin FPSpreadADO.fpSpread fpTableStat 
            Height          =   1395
            Left            =   60
            TabIndex        =   58
            Top             =   780
            Width           =   14715
            _Version        =   393216
            _ExtentX        =   25956
            _ExtentY        =   2461
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
            SpreadDesigner  =   "frmMain.frx":97FD
            ScrollBarTrack  =   1
         End
         Begin FPSpreadADO.fpSpread ftPartition 
            Height          =   1755
            Left            =   -71640
            TabIndex        =   61
            Top             =   420
            Width           =   9495
            _Version        =   393216
            _ExtentX        =   16748
            _ExtentY        =   3096
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
            MaxCols         =   4
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            SelectBlockOptions=   0
            SpreadDesigner  =   "frmMain.frx":BB04
            ScrollBarTrack  =   1
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "INDEX_TBSPACE :"
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
            Left            =   9060
            TabIndex        =   78
            Top             =   450
            Width           =   1680
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "TBSPACE :"
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
            Left            =   5940
            TabIndex        =   76
            Top             =   450
            Width           =   1020
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  '투명
            Caption         =   "Rows 수 :"
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
            Left            =   2100
            TabIndex        =   59
            Top             =   450
            Width           =   900
         End
      End
   End
   Begin VB.Menu mnu_sample 
      Caption         =   "샘플데이타메뉴"
      Visible         =   0   'False
      Begin VB.Menu mnu_exec 
         Caption         =   "Excel"
      End
      Begin VB.Menu mnu_up 
         Caption         =   "Update"
      End
      Begin VB.Menu mnu_in 
         Caption         =   "Insert"
      End
      Begin VB.Menu mnu_Del 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu mnu_excel1 
      Caption         =   "엑셀"
      Visible         =   0   'False
      Begin VB.Menu mnu_excel2 
         Caption         =   "Excel"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public gConn As ADODB.Connection
Public gSelectGrid As String
Const GCONN_USE = False
Dim gaCollist() As String
Dim gaPklist() As String
Dim gaTylist() As String
Dim gbTableSelect As Boolean
'====================================
' 테이블 목록 조회
'====================================

Private Sub cboDanwi2_Click()
  '테이블상태 정보 조회
     Call subTableStat
End Sub

Private Sub cmdAddRow_Click()
     fpInTab.MaxRows = fpInTab.MaxRows + 1
End Sub

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
    
    For i = 1 To fpCol.MaxRows
        
        fpCol.Row = i
        fpCol.Col = 1: sColid = fpCol.Value
        fpCol.Col = 2: sColnm = fpCol.Value
        fpCol.Col = 6: sCodeDomain = fpCol.Value
        fpCol.Col = 7: sColBigo = fpCol.Value
        sColBigo2 = ""
    
    
        sSQL = ""
        sSQL = sSQL & vbCrLf & " DELETE FROM DIC_COL "
        sSQL = sSQL & vbCrLf & "  WHERE OWNER  = '" & cboShema.Text & "'"
        sSQL = sSQL & vbCrLf & "    AND TABID  = '" & edtTable.Text & "' "
        sSQL = sSQL & vbCrLf & "    AND COLID  = '" & sColid & "'"
        
        
        txtSQL.Text = sSQL
        If GCONN_USE = True Then
            gConn.Execute sSQL
        Else
            conn.Execute sSQL
        End If
        
    
        
        sSQL = ""
        sSQL = sSQL & vbCrLf & " INSERT INTO DIC_COL (OWNER, TABID, COLID, COLNM, CDDOM, BIGO1, BIGO2) "
        sSQL = sSQL & vbCrLf & "  VALUES ('" & cboShema.Text & "' ,"
        sSQL = sSQL & vbCrLf & "          '" & edtTable.Text & "' ,"
        sSQL = sSQL & vbCrLf & "          '" & Trim(sColid) & "' ,"
        sSQL = sSQL & vbCrLf & "          '" & Trim(sColnm) & "' ,"
        sSQL = sSQL & vbCrLf & "          '" & Trim(sCodeDomain) & "' ,"
        sSQL = sSQL & vbCrLf & "          '" & Trim(sColBigo) & "' ,"
        sSQL = sSQL & vbCrLf & "          '" & " " & "' )"
        
        
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
    Call fpTable_DblClick(1, fpTable.Row)
    
    
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

Private Sub cmdConnChange_Click()
    'DB설정 정보 레지스트리에서 읽어오기
    Call SaveSetting("DB2_TAB_MAN_BY_HAK", "CONNECT", "PROVIDER", txtProvider.Text)
    Call SaveSetting("DB2_TAB_MAN_BY_HAK", "CONNECT", "DSN", txtDSN.Text)
    Call SaveSetting("DB2_TAB_MAN_BY_HAK", "CONNECT", "USER", txtID.Text)
    Call SaveSetting("DB2_TAB_MAN_BY_HAK", "CONNECT", "PASS", txtPWD.Text)
    Call SaveSetting("DB2_TAB_MAN_BY_HAK", "CONNECT", "DBA", txtDBA.Text)
    Call SaveSetting("DB2_TAB_MAN_BY_HAK", "CONNECT", "DBAPASS", txtDBApass.Text)
    
'    txtProvider.Text = strProvider
'    txtDSN.Text = strDSN
'    txtID.Text = strUser
'    txtPWD.Text = strPass
'    txtDBA.Text = strDBA
'    txtDBApass.Text = strDBAUPass
         
    
    strProvider = txtProvider.Text
    strDSN = txtDSN.Text
    strUser = txtID.Text
    strPass = txtPWD.Text
    strDBA = txtDBA.Text
    strDBAUPass = txtDBApass.Text
    
    MsgBox ("변경완료")

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
    
    If Trim(txtPgmIDLog.Text) = "" Then
        MsgBox "Input Program Id "
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass

    sSQL = ""
    sSQL = sSQL & vbCrLf & "select prog_id                                                                                                      "
    sSQL = sSQL & vbCrLf & "       ,no                                                                                                          "
    sSQL = sSQL & vbCrLf & "       ,msg                                                                                                         "
    sSQL = sSQL & vbCrLf & "       ,to_char(bf_time, 'yyyy-mm-dd hh24:mi:ss')                                                                   "
    sSQL = sSQL & vbCrLf & "       ,to_char(af_time, 'yyyy-mm-dd hh24:mi:ss')                                                                   "
    sSQL = sSQL & vbCrLf & "       ,(DAYS(af_time) - DAYS(bf_time)) * 86400 + (MIDNIGHT_SECONDS(af_time) - MIDNIGHT_SECONDS(bf_time))  / 60     "
    sSQL = sSQL & vbCrLf & "from (                                                                                                              "
    sSQL = sSQL & vbCrLf & "            SELECT prog_id                                                                                          "
    sSQL = sSQL & vbCrLf & "                        ,no                                                                                         "
    sSQL = sSQL & vbCrLf & "                        ,msg                                                                                        "
    sSQL = sSQL & vbCrLf & "                        ,lag(lst_mdfc_dttm,1) over (order by no) bf_time                                            "
    sSQL = sSQL & vbCrLf & "                        ,lst_mdfc_dttm  as af_time                                                                  "
    sSQL = sSQL & vbCrLf & "                    from (select prog_id                                                                            "
    sSQL = sSQL & vbCrLf & "                                            ,no                                                                     "
    sSQL = sSQL & vbCrLf & "                                            ,msg                                                                    "
    sSQL = sSQL & vbCrLf & "                                            ,lst_mdfc_dttm                                                          "
    sSQL = sSQL & vbCrLf & "                                from tb_rsm_debug_l_a                                                               "
    sSQL = sSQL & vbCrLf & "                                where prog_id like '%" & UCase(Trim(txtPgmIDLog.Text)) & "%'                                              "
    sSQL = sSQL & vbCrLf & "                                UNION ALL                                                                           "
    sSQL = sSQL & vbCrLf & "                                select 'ZZZZZ'                                                                       "
    sSQL = sSQL & vbCrLf & "                                            ,9999                                                                   "
    sSQL = sSQL & vbCrLf & "                                            ,'현재step'                                                             "
    sSQL = sSQL & vbCrLf & "                                            , current date                                                          "
    sSQL = sSQL & vbCrLf & "                                from dual                                                                           "
    sSQL = sSQL & vbCrLf & "                            )                                                                                       "
    sSQL = sSQL & vbCrLf & ")                                                                                                                   "
    sSQL = sSQL & vbCrLf & "order by prog_id, no            "
    sSQL = sSQL & vbCrLf & " WITH UR "
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
    
    If Trim(txtPgmIDLog.Text) = "" Then
        MsgBox "Input Program Id "
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    
    
    Dim sStDay As String
    Dim sEdDay As String
    
    sSQL = ""
    sSQL = sSQL & vbCrLf & "SELECT PROG_ID, SEQ, MIN(LST_MDFC_DTTM), MAX(LST_MDFC_DTTM)"
    sSQL = sSQL & vbCrLf & "  From TB_RSM_DEBUG_H_A"
    sSQL = sSQL & vbCrLf & " WHERE prog_id like '%" & UCase(Trim(txtPgmIDLog.Text)) & "%'  "
    sSQL = sSQL & vbCrLf & "   AND LST_MDFC_DTTM BETWEEN FN_RSM_TMSTMP('" & Trim(txtStDay.Text) & "') - 1 day AND FN_RSM_TMSTMP ('" & Trim(txtEdDay.Text) & "') + 1 day"
    sSQL = sSQL & vbCrLf & " GROUP BY PROG_ID, SEQ"
    sSQL = sSQL & vbCrLf & "  ORDER BY PROG_ID, SEQ DESC"

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

Private Sub cmdDescReg_Click()
    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    If Trim(edtTable.Text) = "" Then
        MsgBox ("테이블 선택하세요")
        Exit Sub
    End If
    
    Dim i As Integer
    Dim j As Integer
    Dim sComp As String
    For i = 1 To fpInTab.MaxRows
        
        fpInTab.Row = i
        fpInTab.Col = 1
        sComp = UCase(Trim(fpInTab.Value))
        
        For j = i + 1 To fpInTab.MaxRows
            fpInTab.Row = j
            fpInTab.Col = 1
                    
            If sComp <> "" Then
                If sComp = UCase(Trim(fpInTab.Value)) Then
                    MsgBox j & "번 row의 " & sComp & " 테이블명이 중복 등록되었습니다."
                    Exit Sub
                End If
            End If
        Next j
    Next i
    
    
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
    sSQL = sSQL & vbCrLf & " DELETE FROM DIC_TAB_DESC "
    sSQL = sSQL & vbCrLf & "  WHERE TABID     = '" & edtTable.Text & "'                                                "
    sSQL = sSQL & vbCrLf & "    AND OWNER  = '" & cboShema.Text & "'                                                "
    
    If GCONN_USE = True Then
        gConn.Execute sSQL
    Else
        conn.Execute sSQL
    End If
    

    
    sSQL = ""
    sSQL = sSQL & vbCrLf & " INSERT INTO DIC_TAB_DESC (OWNER, TABID, ORGTAB, CYCLE, USE_YN,IMP_DESC, CR_LOG) "
    sSQL = sSQL & vbCrLf & "  VALUES ('" & cboShema.Text & "' ,"
    sSQL = sSQL & vbCrLf & "          '" & edtTable.Text & "' ,"
    sSQL = sSQL & vbCrLf & "          '" & Trim(txtORGTAB.Text) & "' ,"
    sSQL = sSQL & vbCrLf & "          '" & Trim(txtCYCLE.Text) & "' ,"
    sSQL = sSQL & vbCrLf & "          '" & Trim(txtUSE_YN.Text) & "' ,"
    sSQL = sSQL & vbCrLf & "          '" & Trim(txtIMP_DESC.Text) & "' ,"
    sSQL = sSQL & vbCrLf & "          '" & Trim(txtCR_LOG.Text) & "' )"
    
    

    If GCONN_USE = True Then
        gConn.Execute sSQL
    Else
        conn.Execute sSQL
    End If
    
    
    
    
    
    sSQL = ""
    sSQL = sSQL & vbCrLf & " DELETE FROM DIC_TAB_IN "
    sSQL = sSQL & vbCrLf & "  WHERE TABID     = '" & edtTable.Text & "'                                                "
    sSQL = sSQL & vbCrLf & "    AND OWNER  = '" & cboShema.Text & "'                                                "
    
    If GCONN_USE = True Then
        gConn.Execute sSQL
    Else
        conn.Execute sSQL
    End If
    
    
    
    
    
    For i = 1 To fpInTab.MaxRows
        fpInTab.Row = i
        fpInTab.Col = 1
        
        If Trim(fpInTab.Value) <> "" Then
            sSQL = ""
            sSQL = sSQL & vbCrLf & " INSERT INTO DIC_TAB_IN (OWNER, TABID, INTAB) "
            sSQL = sSQL & vbCrLf & "  VALUES ('" & cboShema.Text & "' ,"
            sSQL = sSQL & vbCrLf & "          '" & edtTable.Text & "' ,"
            sSQL = sSQL & vbCrLf & "          '" & UCase(Trim(fpInTab.Value)) & "') "
        
            If GCONN_USE = True Then
                gConn.Execute sSQL
            Else
                conn.Execute sSQL
            End If
        
        End If
    
    Next i
    
    
    
    
    
    
    If GCONN_USE = True Then
        gConn.CommitTrans
    Else
        conn.CommitTrans
    End If
    
    
    gfMsg ("등록되었습니다.")
    

    '테이블상세정보 조회
    Call subSelectTableDesc
    
    
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

Private Sub cmdEDW_Click()
    
    Dim sUrl As String
    
    sUrl = "http://134.100.207.40/fine_db/EDW/table_group.asp"
    
    
    WebBrowser2.Navigate (sUrl)
End Sub

'부서정보 검색
Private Sub cmdFindDept_Click()
    'http://134.100.1.60/com.ibk.finehr.hrm.comm.EmpnoSearchBranchInfoSelAction.action?brcd=001&sejno=A97011
    'http://134.100.1.60/com.ibk.finehr.hrm.comm.EmpnoSearchBranchListSelAction.action?brnm=영업부&sejno=A97011

    Dim sUrl As String
    
    If IsNumeric(Trim(txtFindKey.Text)) Then
        sUrl = "http://134.100.1.60/com.ibk.finehr.hrm.comm.EmpnoSearchBranchInfoSelAction.action?brcd=" & Trim(txtFindKey.Text)
    Else
        sUrl = "http://134.100.1.60/com.ibk.finehr.hrm.comm.EmpnoSearchBranchListSelAction.action?brnm=" & Trim(txtFindKey.Text)
    End If
    WebBrowser1.Navigate (sUrl)
End Sub

'직원정보 검색
Private Sub cmdFindEmp_Click()
    'http://134.100.1.60/com.ibk.finehr.hrm.comm.EmpnoSearchEmpnoInfoSelAction.action?empno=20772&sejno=A97011&fromBr=N
    'http://134.100.1.60/com.ibk.finehr.hrm.comm.EmpnoSearchEmpnoListSelAction.action?mname=오성훈&sejno=A97011
    
    Dim sUrl As String
    
    If IsNumeric(Trim(txtFindKey.Text)) Then
        sUrl = "http://134.100.1.60/com.ibk.finehr.hrm.comm.EmpnoSearchEmpnoInfoSelAction.action?empno=" & Trim(txtFindKey.Text)
    Else
        sUrl = "http://134.100.1.60/com.ibk.finehr.hrm.comm.EmpnoSearchEmpnoListSelAction.action?mname=" & Trim(txtFindKey.Text)
    End If
    WebBrowser1.Navigate (sUrl)
End Sub

'지캠프 직원찾기
Private Sub cmdFineNetemp_Click()
    WebBrowser1.Navigate ("http://134.100.1.60/hrm/comm/empno_search_sel_frm.jsp")
End Sub

Private Sub cmdFrameClose_Click()
    frameCode.Visible = False
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
    
    sSQL = sSQL & vbCrLf & " WITH UR "
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

'회사그룹웨어
Private Sub cmdIBKSGroup_Click()
    WebBrowser1.Navigate ("http://iworld.ibksystem.co.kr/iFrame/out/main.xml")
End Sub


'IBKS직원찾기
Private Sub cmdIBKsystem_Click()
    WebBrowser1.Navigate ("https://ehr.ibksystem.co.kr/SMSearch_popup.jsp")
End Sub

Private Sub cmdInfoCenter_Click()
    Dim sUrl As String
    
    sUrl = "http://localhost:51000/help/index.jsp"

    WebBrowser3.Navigate (sUrl)
End Sub

Private Sub cmdJobLog_Click()
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
    
    If Trim(txtPgmIDLog2.Text) = "" Then
        MsgBox "Input Program Id "
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass

    sSQL = ""
    sSQL = sSQL & vbCrLf & "select PRFR_YMD                  -- 집행년월일                                                                                                                                                                                                          "
    sSQL = sSQL & vbCrLf & "       ,JOB_GRP_NM                -- 작업그룹명                                                                                                                                                                                                         "
    sSQL = sSQL & vbCrLf & "       ,JOB_STTM                  -- 작업시작시각                                                                                                                                                                                                       "
    sSQL = sSQL & vbCrLf & "       ,JOB_FNSH_TMD              -- 작업종료시각                                                                                                                                                                                                       "
    sSQL = sSQL & vbCrLf & "             --,(DAYS(JOB_FNSH_TMD) - DAYS(JOB_STTM)) * 86400 + (MIDNIGHT_SECONDS(JOB_FNSH_TMD) - MIDNIGHT_SECONDS(JOB_STTM))  / 60                                                                                                                     "
    sSQL = sSQL & vbCrLf & "            ,(                                                                                                                                                                                                                                          "
    sSQL = sSQL & vbCrLf & "                DAYS( CASE WHEN  JOB_FNSH_TMD < JOB_STTM THEN '2009-01-02' ELSE '2009-01-01' END || ' '  ||SUBSTR(JOB_FNSH_TMD,1,2) || ':' || SUBSTR(JOB_FNSH_TMD,3,2) || ':' ||SUBSTR(JOB_FNSH_TMD,5,2) )                                              "
    sSQL = sSQL & vbCrLf & "            - DAYS( '2009-01-01'  || ' '  ||SUBSTR(JOB_STTM,1,2) || ':' || SUBSTR(JOB_STTM,3,2) || ':' ||SUBSTR(JOB_STTM,5,2) )                                                                                                                         "
    sSQL = sSQL & vbCrLf & "            )                                                                                                                                                                                                                                           "
    sSQL = sSQL & vbCrLf & "            * 86400                                                                                                                                                                                                                                     "
    sSQL = sSQL & vbCrLf & "            + (                                                                                                                                                                                                                                         "
    sSQL = sSQL & vbCrLf & "                    MIDNIGHT_SECONDS(CASE WHEN  JOB_FNSH_TMD < JOB_STTM THEN '2009-01-02' ELSE '2009-01-01' END || ' '  ||SUBSTR(JOB_FNSH_TMD,1,2) || ':' || SUBSTR(JOB_FNSH_TMD,3,2) || ':' ||SUBSTR(JOB_FNSH_TMD,5,2) )                               "
    sSQL = sSQL & vbCrLf & "            -                                                                                                                                                                                                                                           "
    sSQL = sSQL & vbCrLf & "                        MIDNIGHT_SECONDS('2009-01-01'  || ' '  ||SUBSTR(JOB_STTM,1,2) || ':' || SUBSTR(JOB_STTM,3,2) || ':' ||SUBSTR(JOB_STTM,5,2) )                                                                                                    "
    sSQL = sSQL & vbCrLf & "            )  / 60                                                                                                                                                                                                                                     "
    sSQL = sSQL & vbCrLf & "                                                                                                                                                                                                                                                        "
    sSQL = sSQL & vbCrLf & "       ,PCSN_RSLT_YN              -- 처리결과여부                                                                                                                                                                                                       "
    sSQL = sSQL & vbCrLf & "       ,TBL_ID                    -- 테이블ID                                                                                                                                                                                                           "
    sSQL = sSQL & vbCrLf & "       ,IQCN                      -- 조회건수                                                                                                                                                                                                           "
    sSQL = sSQL & vbCrLf & "       ,ENIN_DLTN_CNT             -- 기업정보삭제건수                                                                                                                                                                                                   "
    sSQL = sSQL & vbCrLf & "       ,NRMT_PCSN_CNT             -- 정상처리건수                                                                                                                                                                                                       "
    sSQL = sSQL & vbCrLf & "       ,ENIN_MDFC_CNT             -- 기업정보변경건수                                                                                                                                                                                                   "
    sSQL = sSQL & vbCrLf & "       ,ERR_CNT                   -- 오류건수                                                                                                                                                                                                           "
    sSQL = sSQL & vbCrLf & "       ,ERR_CNTN                  -- 오류내용                                                                                                                                                                                                           "
    sSQL = sSQL & vbCrLf & "       ,MSG_CNTN                    -- 메시지내용                                                                                                                                                                                                       "
    sSQL = sSQL & vbCrLf & "       from TB_RSM_CO004_C_A                                                                                                                                                                                                                            "
    sSQL = sSQL & vbCrLf & "where job_grp_nm like '%" & UCase(Trim(txtPgmIDLog2.Text)) & "%' --and PRFR_YMD >= '20091029' --and JOB_STTM >= '090000'정상 완료[-]                                                                                                                                           "
    sSQL = sSQL & vbCrLf & "  and  PRFR_YMD between '" & Trim(txtStDay.Text) & "' and '" & Trim(txtEdDay.Text) & "'"
    If Trim(cboStat.Text) <> "" Then
        sSQL = sSQL & vbCrLf & "  and  PCSN_RSLT_YN = '" & cboStat.Text & "'"
    End If
    If Trim(txtMessage.Text) <> "" Then
        sSQL = sSQL & vbCrLf & "  and  MSG_CNTN LIKE '" & txtMessage.Text & "%'"
    End If
    
    sSQL = sSQL & vbCrLf & "order by JOB_GRP_NM ASC, PRFR_YMD  DESC  ,JOB_STTM desc                                                                                                                                                                                                                 "
    sSQL = sSQL & vbCrLf & "fetch first 2000 rows only                                                                                                                                                                                                                              "

    sSQL = sSQL & vbCrLf & " WITH UR "
    txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If

    fpJobLog.MaxRows = 0
    If rsTemp.RecordCount < 1 Then
        fpJobLog.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpJobLog, rsTemp)
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

Private Sub cmdOnline_Click()
    WebBrowser1.Navigate ("https://publib.boulder.ibm.com/infocenter/db2luw/v9r7/index.jsp")
End Sub

Private Sub cmdPartitionSearch_Click()
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

    sSQL = sSQL & vbCrLf & "    SELECT  PARTITION_NAME, HIGH_VALUE, TABLESPACE_NAME, PARTITION_POSITION "
    sSQL = sSQL & vbCrLf & " From SYSIBMADM.ALL_TAB_PARTITIONS "
    sSQL = sSQL & vbCrLf & "WHERE TABLE_NAME = '" & edtTable.Text & "' "
    sSQL = sSQL & vbCrLf & "ORDER BY HIGH_VALUE "
    
    
    txtSQL.Text = sSQL
    sSQL = sSQL & vbCrLf & " WITH UR "
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If

    ftPartition.MaxRows = 0
    If rsTemp.RecordCount < 1 Then
        ftPartition.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(ftPartition, rsTemp)
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

'기준년월별 count 조회
Private Sub cmdRecCountSrch_Click()
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
    sSQL = sSQL & vbCrLf & "select " & txtGrpByCol.Text & ", count(*) "
    sSQL = sSQL & vbCrLf & "  from " & edtTable.Text
    sSQL = sSQL & vbCrLf & " group by " & txtGrpByCol.Text
    sSQL = sSQL & vbCrLf & " order by  " & txtGrpByCol.Text
    
    sSQL = sSQL & vbCrLf & " WITH UR "
    txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If

    fpRecCount.MaxRows = 0
    If rsTemp.RecordCount < 1 Then
        fpRecCount.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpRecCount, rsTemp)
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


Private Sub cmdRecCountSrch2_Click()
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
    
    
    sSQL = sSQL & vbCrLf & " SELECT *  FROM (  "
    
    
    Dim i As Integer
    Dim j As Integer
    Dim baseym As String
    
    For i = 1995 To 2011
        For j = 1 To 12
            If j <= 9 Then
                baseym = CStr(i) & "0" & CStr(j)
            Else
                baseym = CStr(i) & CStr(j)
            End If
            
            sSQL = sSQL & vbCrLf & " SELECT " & txtGrpByCol2.Text & " FROM " & edtTable.Text & " WHERE " & txtGrpByCol2.Text & " = '" & baseym & "' and rownum <= 1 union all"
        Next j
    Next i
    sSQL = sSQL & vbCrLf & "  select '9999-12-31' from dual     ) "
    sSQL = sSQL & vbCrLf & "  ORDER BY 1 "
    sSQL = sSQL & vbCrLf & " WITH UR "
    
    
    txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If

    fpRecCount2.MaxRows = 0
    If rsTemp.RecordCount < 1 Then
        fpRecCount2.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpRecCount2, rsTemp)
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

'테이블 스페이스 조회
Private Sub cmdTableSpaceSrch_Click()
    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    
    If GCONN_USE = False Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = GetConnStringDBA
    
        conn.Open
    End If
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
    Dim sDanwi  As String
    If cboDanwi.Text = "MB" Then
        sDanwi = "/1024/1024"
    End If

    If cboDanwi.Text = "GB" Then
        sDanwi = "/1024/1024/1024"
    End If

    sSQL = ""
    sSQL = sSQL & vbCrLf & "select value(substr(tablespace_name,1,20),'TOTAL') as tablespace,   "
    sSQL = sSQL & vbCrLf & "       case when TABLESPACE_TYPE = 0 then 'DMS'           "
    sSQL = sSQL & vbCrLf & "            when TABLESPACE_TYPE = 1 then 'SMS' end as type,             "
    sSQL = sSQL & vbCrLf & "       max(page_size) as page_size,"
    sSQL = sSQL & vbCrLf & "       sum(total_pages*page_size) " & sDanwi & " as total_size_MB,"
    sSQL = sSQL & vbCrLf & "       sum(used_pages*page_size)" & sDanwi & " as used_size_MB,"
    sSQL = sSQL & vbCrLf & "       sum(free_pages*page_size)" & sDanwi & "  as free_size_MB,"
    sSQL = sSQL & vbCrLf & "       case when sum(total_pages*page_size)/1024/1024 <> 0"
    sSQL = sSQL & vbCrLf & "       then dec((dec(sum(used_pages*page_size)/1024/1024 ) / dec(sum(total_pages*page_size)/1024/1024 )) * 100,5,2)"
    sSQL = sSQL & vbCrLf & "       else 0 end USED_PCT"
    sSQL = sSQL & vbCrLf & "  from table(snapshot_tbs_cfg('', -1)) as snapshot_tbs_cfg"
    sSQL = sSQL & vbCrLf & " where tablespace_type = 0 "
    sSQL = sSQL & vbCrLf & " group by rollup(substr(tablespace_name,1,20)),TABLESPACE_TYPE "
    
    
    sSQL = sSQL & vbCrLf & " WITH UR "
    txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If

    fpTableSPace.MaxRows = 0
    If rsTemp.RecordCount < 1 Then
        fpTableSPace.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpTableSPace, rsTemp)
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




'테이블상태 조회
Private Sub subTableStat()
    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
Exit Sub
    Screen.MousePointer = vbHourglass
    
    If GCONN_USE = False Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = GetConnStringDBA
    
        conn.Open
    End If
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
    Dim sDanwi  As String
    If cboDanwi2.Text = "KB" Then
        sDanwi = "/1"
    End If
    
    If cboDanwi2.Text = "MB" Then
        sDanwi = "/1024"
    End If

    If cboDanwi2.Text = "GB" Then
        sDanwi = "/1024/1024"
    End If

    sSQL = ""
    sSQL = sSQL & vbCrLf & "SELECT CASE AVAILABLE WHEN 'N' THEN '사용불가' ELSE '사용가' END 상태, "
    sSQL = sSQL & vbCrLf & "       DOUBLE(DATA_OBJECT_L_SIZE) " & sDanwi & " AS 논리사이즈, "
    sSQL = sSQL & vbCrLf & "       DOUBLE(DATA_OBJECT_P_SIZE) " & sDanwi & " AS 실제사이즈, "
    sSQL = sSQL & vbCrLf & "       DOUBLE(INDEX_OBJECT_L_SIZE) " & sDanwi & " AS INDEX논리사이즈, "
    sSQL = sSQL & vbCrLf & "       DOUBLE(INDEX_OBJECT_P_SIZE) " & sDanwi & " AS INDEX물리사이즈, "
    sSQL = sSQL & vbCrLf & "       CASE REORG_PENDING WHEN 'Y' THEN 'REORG권장' ELSE REORG_PENDING END REORG, "
    sSQL = sSQL & vbCrLf & "       INPLACE_REORG_STATUS AS inplace상태, "
    sSQL = sSQL & vbCrLf & "       LOAD_STATUS AS LOAD상태, "
    sSQL = sSQL & vbCrLf & "       READ_ACCESS_ONLY AS 읽기전용상태,  "
    sSQL = sSQL & vbCrLf & "       NO_LOAD_RESTART AS LOAD재시작, "
    sSQL = sSQL & vbCrLf & "       NUM_REORG_REC_ALTERS AS REORG권장변경조작, "
    sSQL = sSQL & vbCrLf & "       INDEXES_REQUIRE_REBUILD AS INDEX리빌드 "
    sSQL = sSQL & vbCrLf & "  FROM TABLE (SYSPROC.ADMIN_GET_TAB_INFO('" & cboShema.Text & "', '" & edtTable.Text & "'))  AS T "

    
    sSQL = sSQL & vbCrLf & " WITH UR "
    txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If

    Dim i As Integer
    
    For i = 1 To rsTemp.Fields.Count
        fpTableStat.Row = 0
        fpTableStat.Col = i
        fpTableStat.Value = rsTemp.Fields(i - 1).Name
        
    Next i
    fpTableStat.MaxCols = rsTemp.Fields.Count


    fpTableStat.MaxRows = 0
    If rsTemp.RecordCount < 1 Then
        fpTableStat.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpTableStat, rsTemp)
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


    sSQL = ""
    sSQL = sSQL & vbCrLf & "SELECT A.NAME, B.TABNM, B.BIGO1,  A.CTIME,     "
    sSQL = sSQL & vbCrLf & "       B.BIGO2 , A.CREATOR, A.CARD, A.TBSPACE , A.INDEX_TBSPACE       "
    sSQL = sSQL & vbCrLf & "  FROM SYSIBM.SYSTABLES A,               "
    sSQL = sSQL & vbCrLf & "      DIC_TAB B                          "
    sSQL = sSQL & vbCrLf & " WHERE A.TYPE = 'T'                      "
    sSQL = sSQL & vbCrLf & "   AND A.CREATOR = '" & cboShema.Text & "'"
    sSQL = sSQL & vbCrLf & "   AND A.CREATOR = B.OWNER (+)           "
    sSQL = sSQL & vbCrLf & "   AND A.NAME = B.TABID (+)              "
        
    If Trim(txtTableSrchEng.Text) <> "" Then
        sSQL = sSQL & vbCrLf & " AND UPPER(A.NAME) LIKE '%" + UCase(Trim(txtTableSrchEng.Text)) + "%'"
    ElseIf Trim(txtTableSrchKor.Text) <> "" Then
        sSQL = sSQL & vbCrLf & " AND  B.TABNM  LIKE '%" + Trim(txtTableSrchKor.Text) + "%'"
    ElseIf Trim(txtTableSrchBigo.Text) <> "" Then
        sSQL = sSQL & vbCrLf & " AND  UPPER(B.BIGO1)  LIKE '%" + UCase(Trim(txtTableSrchBigo.Text)) + "%'"
    End If
    
    sSQL = sSQL & vbCrLf & " ORDER BY A.NAME"
    
    sSQL = sSQL & vbCrLf & " WITH UR "
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
    sSQL = sSQL & vbCrLf & " DELETE FROM DIC_TAB "
    sSQL = sSQL & vbCrLf & "  WHERE TABID     = '" & edtTable.Text & "'                                                "
    sSQL = sSQL & vbCrLf & "    AND OWNER  = '" & cboShema.Text & "'                                                "
    
    If GCONN_USE = True Then
        gConn.Execute sSQL
    Else
        conn.Execute sSQL
    End If
    

    
    sSQL = ""
    sSQL = sSQL & vbCrLf & " INSERT INTO DIC_TAB (OWNER, TABID, TABNM, BIGO1, BIGO2) "
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

Private Sub Command10_Click()
    Dim tempDay As String
    tempDay = DateAdd("m", -1, Now)
    txtStDay.Text = Mid(tempDay, 1, 4) & Mid(tempDay, 6, 2) & Mid(tempDay, 9, 2)
End Sub

Private Sub Command11_Click()
    txtMessage.Text = ""
End Sub

Private Sub Command12_Click()
    Dim tempDay As String
    tempDay = Now
    txtEdDay.Text = Mid(tempDay, 1, 4) & Mid(tempDay, 6, 2) & Mid(tempDay, 9, 2)
End Sub

Private Sub Command2_Click()
    'fpSpread1.MaxRows = 0
    'EZFTP1.RemoteAddress = "134.100.204.72"
    'EZFTP1.UserName = "tftp"
    'EZFTP1.Password = "tftp1234"
    'EZFTP1.Connect
    'EZFTP1.GetDirectory ("*.*")
    'EZFTP1.GetDirectory
    
    
    
End Sub



Private Sub Command1_Click()
    Clipboard.SetText "김경학"
End Sub

Private Sub Command4_Click()
     Clipboard.SetText txUpdateSQL.Text
End Sub

Private Sub cmdValueSelect_Click()
    Call subSelectData("")
End Sub

Private Sub Command3_Click()
    Clipboard.SetText txtInsertSQL.Text
End Sub

Private Sub Command5_Click()
    Clipboard.SetText txtSelectSQL.Text
End Sub

Private Sub Command6_Click()
    'SQL display
    Call subMakeSQL
End Sub

Private Sub Command8_Click()
    Dim tempDay As String
    tempDay = DateAdd("d", -1, Now)
    txtStDay.Text = Mid(tempDay, 1, 4) & Mid(tempDay, 6, 2) & Mid(tempDay, 9, 2)
End Sub

Private Sub Command9_Click()
    Dim tempDay As String
    tempDay = DateAdd("d", -7, Now)
    txtStDay.Text = Mid(tempDay, 1, 4) & Mid(tempDay, 6, 2) & Mid(tempDay, 9, 2)
End Sub

Private Sub EZFTP1_NextDirectoryEntry(ByVal FileName As String, ByVal Attributes As Long, ByVal Length As Double)

'
'    fpSpread1.MaxRows = fpSpread1.MaxRows + 1
'    fpSpread1.Row = fpSpread1.MaxRows
'    fpSpread1.Col = 1
'    fpSpread1.Value = FileName
'    fpSpread1.Col = 2
'    fpSpread1.Value = Attributes
'    fpSpread1.Col = 3
'    fpSpread1.Value = Length
    
    
End Sub

'====================================
' Form Load 이벤트
'====================================
Private Sub Form_Load()
    '콤보박스 셋팅
    cboShema.AddItem ("RSMUSER")
    cboShema.ListIndex = 0
    
    If GCONN_USE = True Then
        If Not fConn Then
            gfMsg ("DB 접속오류")
        End If
    End If
    
    WebBrowser1.Navigate ("http://134.100.1.60/hrm/comm/empno_search_sel_frm.jsp")
    
    frameCode.Visible = False
    
    SSTab1.Tab = 1
    SSTab2.Tab = 0
    SSTab3.Tab = 0
    SSTab4.Tab = 0
    
    
    frmMain.Left = 200
    frmMain.Top = 0
    
    'edw URL go
    cmdEDW_Click
    
    'infocenter URL go
    cmdInfoCenter_Click
        
    
    Dim tempDay As String
    tempDay = DateAdd("m", -1, Now)
    
    txtStDay.Text = Mid(tempDay, 1, 4) & Mid(tempDay, 6, 2) & Mid(tempDay, 9, 2)
    txtEdDay.Text = Mid(Now, 1, 4) & Mid(Now, 6, 2) & Mid(Now, 9, 2)
End Sub

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

Private Sub Form_Unload(Cancel As Integer)
    Set gConn = Nothing
End Sub


Private Sub fpCol_Click(ByVal Col As Long, ByVal Row As Long)
    If Col <> 6 Then Exit Sub
    fpCol.Row = Row
    fpCol.Col = Col
    
    
    Dim sCode As String
    
    If Trim(fpCol.Value) = "" Then Exit Sub
    
    If fpCol.Value = "전행코드" Then
        '전행공통코드 조회
        fpCol.Col = 1
        sCode = fpCol.Value
        
        subSelectCodeListBankAll (sCode)
    Else
        'RSM 공통코드 조회
        sCode = fpCol.Value
        
        subSelectCodeList (sCode)
        frameCode.Visible = True
    End If
    
    
End Sub


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


  

    sSQL = ""
    sSQL = sSQL & vbCrLf & "SELECT 'RSM' || CD_DMN_NO, CD_DMN_KRN_NM, STND_CD, CD_NM, CD_DTL_DSCP     "
    sSQL = sSQL & vbCrLf & "  FROM TB_RSM_CO002_C_A         "
    sSQL = sSQL & vbCrLf & " WHERE CD_DMN_NO = '" & sCode & "'                      "
    sSQL = sSQL & vbCrLf & " ORDER BY  STND_CD"
    
    sSQL = sSQL & vbCrLf & " WITH UR "
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


Private Sub subSelectCodeListBankAll(ByVal sCode As String)
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
    sSQL = sSQL & vbCrLf & "SELECT STND_TDCT_CD || '-' || CD_DMN_NO, CD_DMN_KRN_NM, STND_CD, CD_NM, CD_DTL_DSCP     "
    sSQL = sSQL & vbCrLf & "  FROM TB_RSM_OD236_C_D         "
    sSQL = sSQL & vbCrLf & " WHERE CD_DMN_ENSN_NM = '" & sCode & "'                      "
    sSQL = sSQL & vbCrLf & " ORDER BY  STND_TDCT_CD, CD_DMN_NO,  STND_CD "
    
    sSQL = sSQL & vbCrLf & " WITH UR "
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
        frameCode.Visible = True
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



'**********************************************************
'컬럼 정보 더블 클릭 이벤트 처리
'**********************************************************

Private Sub fpCol_DblClick(ByVal Col As Long, ByVal Row As Long)

    

On Error GoTo ErrorHandler
            
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


Private Sub fpDebugLog_DblClick(ByVal Col As Long, ByVal Row As Long)
    Call gfGrid_Sort(fpDebugLog, Col, Row)
                        
    If Row < 1 Then Exit Sub
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
    sSQL = sSQL & vbCrLf & "select prog_id   ,seq                                                                                                 "
    sSQL = sSQL & vbCrLf & "       ,no                                                                                                          "
    sSQL = sSQL & vbCrLf & "       ,msg                                                                                                         "
    sSQL = sSQL & vbCrLf & "       ,to_char(bf_time, 'yyyy-mm-dd hh24:mi:ss')                                                                   "
    sSQL = sSQL & vbCrLf & "       ,to_char(af_time, 'yyyy-mm-dd hh24:mi:ss')                                                                   "
    sSQL = sSQL & vbCrLf & "       ,(DAYS(af_time) - DAYS(bf_time)) * 86400 + (MIDNIGHT_SECONDS(af_time) - MIDNIGHT_SECONDS(bf_time))  / 60     "
    sSQL = sSQL & vbCrLf & "from (                                                                                                              "
    sSQL = sSQL & vbCrLf & "            SELECT prog_id   ,seq                                                                                   "
    sSQL = sSQL & vbCrLf & "                        ,no                                                                                         "
    sSQL = sSQL & vbCrLf & "                        ,msg                                                                                        "
    sSQL = sSQL & vbCrLf & "                        ,lag(lst_mdfc_dttm,1) over (order by no) bf_time                                            "
    sSQL = sSQL & vbCrLf & "                        ,lst_mdfc_dttm  as af_time                                                                  "
    sSQL = sSQL & vbCrLf & "                    from (select prog_id      ,seq                                                                     "
    sSQL = sSQL & vbCrLf & "                                            ,no                                                                      "
    sSQL = sSQL & vbCrLf & "                                            ,msg                                                                    "
    sSQL = sSQL & vbCrLf & "                                            ,lst_mdfc_dttm                                                          "
    sSQL = sSQL & vbCrLf & "                                from TB_RSM_DEBUG_H_A                                                               "
    sSQL = sSQL & vbCrLf & "                                where prog_id = '" & sPgm & "'                                              "
    sSQL = sSQL & vbCrLf & "                                AND SEQ = '" & sSeq & "'                                              "
    sSQL = sSQL & vbCrLf & "                                UNION ALL                                                                           "
    sSQL = sSQL & vbCrLf & "                                select 'ZZZZZ'  , " & sSeq
    sSQL = sSQL & vbCrLf & "                                            ,9999                                                                   "
    sSQL = sSQL & vbCrLf & "                                            ,'현재step'                                                             "
    sSQL = sSQL & vbCrLf & "                                            , current date                                                          "
    sSQL = sSQL & vbCrLf & "                                from dual                                                                           "
    sSQL = sSQL & vbCrLf & "                            )                                                                                       "
    sSQL = sSQL & vbCrLf & ")                                                                                                                   "
    sSQL = sSQL & vbCrLf & "order by prog_id, no            "
    sSQL = sSQL & vbCrLf & " WITH UR "
    txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If

    fpDebugLogHistDtail.MaxRows = 0
    If rsTemp.RecordCount < 1 Then
        fpDebugLogHistDtail.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpDebugLogHistDtail, rsTemp)
    End If
    ssTabeLog.Tab = 1
    
    
    
    
    
    
    
    
    
    
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

Private Sub fpJobLog_DblClick(ByVal Col As Long, ByVal Row As Long)
    Call gfGrid_Sort(fpJobLog, Col, Row)
                        
    If Row < 1 Then Exit Sub
End Sub

Private Sub fpSample_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    fpSample.Row = 1
    fpSample.Col = 1
    If fpSample.Value <> "" Then
        PopupMenu Me.mnu_sample, , SSTab2.Left + fpSample.Left + MouseX + 100, SSTab2.Top + fpSample.Top + MouseY
        gSelectGrid = "fpSample"
    End If
End Sub

'------------------------------------------------------
' 테이블 더블클릭 시 테이블 정의서 조회
'------------------------------------------------------
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
        .Col = 7: txtRowCnt.Text = .Text
        .Col = 8: txtTBSpace.Text = .Text
        .Col = 9: txtIndexTbspace.Text = .Text
    End With
        
    '컬럼정보CLEAR
    txtColId.Text = ""
    txtColNm.Text = ""
    txtCodeDomain.Text = ""
    txtColBigo.Text = ""
        
    ftPartition.MaxRows = 0
    
    Screen.MousePointer = vbHourglass
    
    
    '테이블상태 정보 조회
     Call subTableStat
    
    If GCONN_USE = False Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = GetConnString
        conn.Open
    End If
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient




    sSQL = ""
    sSQL = sSQL & vbCrLf & "SELECT A.NAME AS COLID, B.COLNM,  A.COLTYPE,                                                 "
    sSQL = sSQL & vbCrLf & "       CASE A.COLTYPE WHEN 'DECIMAL' THEN TRIM(CHAR(A.LENGTH)) || ',' || TRIM(CHAR(A.SCALE)) "
    sSQL = sSQL & vbCrLf & "                    WHEN 'CHAR' THEN TRIM(CHAR(A.LENGTH))                                    "
    sSQL = sSQL & vbCrLf & "                    WHEN 'VARCHAR' THEN TRIM(CHAR(A.LENGTH))                                 "
    sSQL = sSQL & vbCrLf & "                    WHEN 'SMALLINT' THEN TRIM(CHAR(A.LENGTH))                                "
    sSQL = sSQL & vbCrLf & "                    WHEN 'TIMESTMP' THEN TRIM(CHAR(A.LENGTH))                                "
    sSQL = sSQL & vbCrLf & "                    WHEN 'CLOB' THEN TRIM(CHAR(A.LENGTH))                                    "
    sSQL = sSQL & vbCrLf & "                    WHEN 'INTEGER' THEN TRIM(CHAR(A.LENGTH))                                 "
    sSQL = sSQL & vbCrLf & "                    WHEN 'DATE' THEN TRIM(CHAR(A.LENGTH))                                    "
    sSQL = sSQL & vbCrLf & "                    ELSE 'ERR'                                                               "
    sSQL = sSQL & vbCrLf & "     END LENGTH,                                                                             "
    sSQL = sSQL & vbCrLf & "     A.KEYSEQ,                                                                               "
    sSQL = sSQL & vbCrLf & "     CASE WHEN B.CDDOM IS NULL OR  TRIM(B.CDDOM) = '' THEN                                   "
    sSQL = sSQL & vbCrLf & "               CASE WHEN B.COLNM LIKE '%코드%'  THEN                                         "
    sSQL = sSQL & vbCrLf & "                                                       '전행코드'                            "
'    sSQL = sSQL & vbCrLf & "                          ( SELECT CASE WHEN COUNT(*) > 1 THEN '전행코드' ELSE '' END        "
'    sSQL = sSQL & vbCrLf & "                             FROM TB_RSM_OD236_C_D                                           "
'    sSQL = sSQL & vbCrLf & "                            WHERE CD_DMN_ENSN_NM = A.NAME    )                               "
    sSQL = sSQL & vbCrLf & "                END                                                                          "
    sSQL = sSQL & vbCrLf & "          ELSE B.CDDOM                                                                       "
    sSQL = sSQL & vbCrLf & "     END ,                                                                                   "
    sSQL = sSQL & vbCrLf & "     B.BIGO1,                                                                                "
    sSQL = sSQL & vbCrLf & "     A.COLNO,                                                                                "
    sSQL = sSQL & vbCrLf & "     '' AS VALU                                                                              "
    sSQL = sSQL & vbCrLf & " FROM SYSIBM.SYSCOLUMNS A,                                                                   "
    sSQL = sSQL & vbCrLf & "      DIC_COL B                                                                              "
    sSQL = sSQL & vbCrLf & " WHERE A.TBNAME     = '" & edtTable.Text & "'                                                "
    sSQL = sSQL & vbCrLf & "   AND A.TBCREATOR  = '" & cboShema.Text & "'                                                "
    sSQL = sSQL & vbCrLf & "   AND A.TBCREATOR = B.OWNER(+)                                                              "
    sSQL = sSQL & vbCrLf & "   AND A.TBNAME    = B.TABID(+)                                                              "
    sSQL = sSQL & vbCrLf & "   AND A.NAME      = B.COLID(+)                                                              "
    sSQL = sSQL & vbCrLf & "ORDER BY COLNO                                                                               "
    
    
    sSQL = sSQL & vbCrLf & " WITH UR "
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
    
    Dim i As Integer
    
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
    sSQL = sSQL & vbCrLf & " SELECT NAME, COLNAMES, CREATE_TIME "
    sSQL = sSQL & vbCrLf & "   FROM SYSIBM.SYSINDEXES "
    sSQL = sSQL & vbCrLf & " WHERE TBNAME     = '" & edtTable.Text & "'                                                "
    sSQL = sSQL & vbCrLf & "   AND TBCREATOR  = '" & cboShema.Text & "'                                                "
    
    
    sSQL = sSQL & vbCrLf & " WITH UR "
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
    
    
    '샘플데이터 조회
    subSelectData ("")
    
    '테이블상세정보 조회
    Call subSelectTableDesc
        
    'group by col
    fpCol.Row = 1
    fpCol.Col = 1
    txtGrpByCol.Text = fpCol.Text
    txtGrpByCol2.Text = fpCol.Text
    fpRecCount.MaxRows = 0
    

    
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

'------------------------------------------------------
' 샘플데이터 조회
'------------------------------------------------------
Private Sub subMakeSQL()
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
    
    Dim sUSQL As String
    Dim sUSQLWhere As String
    sUSQLWhere = ""
    
    Dim sSSQL As String
    
    
    Dim sGSQL As String
    
    sISQL = "INSERT INTO " & edtTable.Text & " ( " & "        /* " & edtTablenm.Text & "*/" & vbCrLf
    sISQLValue = "( "
    
    sUSQL = "UPDATE " & edtTable.Text & "        /* " & edtTablenm.Text & "*/" & vbCrLf & "   SET " & vbCrLf
    sUSQLWhere = ""
    
    
    sSSQL = "SELECT "
    
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
        Else
            sISQL = sISQL & "          "
        End If
        
        
        sISQL = sISQL & "" & sCode
        sISQL = sISQL & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
        
        If i <> fpCol.MaxRows Then
            sISQL = sISQL & "   -- " & sName
        Else
            sISQL = sISQL & "     -- " & sName
        End If
        sISQL = sISQL & vbCrLf
        
        'insert values
        If i <> 1 Then
            sISQLValue = sISQLValue & "       ,"
        Else
            sISQLValue = sISQLValue & "      "
        End If
        
        
        If sTy = "D" Or sTy = "I" Then
            sISQLValue = sISQLValue & "XXX"
        Else
            sISQLValue = sISQLValue & "'XXX'"
        End If
        
        sISQLValue = sISQLValue & "          -- " & sName
        sISQLValue = sISQLValue & vbCrLf
        
        
        'update field
        If i <> 1 Then
            sUSQL = sUSQL & "       ,"
        Else
            sUSQL = sUSQL & "        "
        End If
        
        sUSQL = sUSQL & "" & sCode
        sUSQL = sUSQL & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
        
        
        If sTy = "D" Or sTy = "I" Then
            sUSQL = sUSQL & "    = XXX"
        Else
            sUSQL = sUSQL & "    = 'XXX'"
        End If
        sUSQL = sUSQL & "          -- " & sName
        sUSQL = sUSQL & vbCrLf
        
        'where
        If Trim(sUSQLWhere) = "" Then
            sUSQLWhere = sUSQLWhere & " WHERE "
        Else
            sUSQLWhere = sUSQLWhere & "   AND "
        End If
        
        sUSQLWhere = sUSQLWhere & IIf(Trim(txtAlias.Text) <> "", Trim(txtAlias.Text) & ".", "") & sCode
        sUSQLWhere = sUSQLWhere & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
        If sTy = "D" Or sTy = "I" Then
            sUSQLWhere = sUSQLWhere & "    = XXX"
        Else
            sUSQLWhere = sUSQLWhere & "    = 'XXX'"
        End If
        sUSQLWhere = sUSQLWhere & "          -- " & sName
        sUSQLWhere = sUSQLWhere & vbCrLf
        
        
        'SELECT FIELD
        
        If i <> 1 Then
            sSSQL = sSSQL & "       ,"
            
            fpCol.Col = 5
            If Trim(fpCol.Value) <> "" Then
                sGSQL = sGSQL & "       ,"
            End If
        End If
        
        
        
        If i = 1 Then
            sSSQL = sSSQL & " " & IIf(Trim(txtAlias.Text) <> "", Trim(txtAlias.Text) & ".", "") & sCode
            
            fpCol.Col = 5
            If Trim(fpCol.Value) <> "" Then
                sGSQL = sGSQL & " " & sCode
            End If
            
            
        Else
            sSSQL = sSSQL & "" & IIf(Trim(txtAlias.Text) <> "", Trim(txtAlias.Text) & ".", "") & sCode
        
            fpCol.Col = 5
            If Trim(fpCol.Value) <> "" Then
                sGSQL = sGSQL & " " & sCode
            End If
        
        
        End If
        sSSQL = sSSQL & gfMakeString(" ", iMaxLen - Len(sCode) + iSpace)
        
        If i <> fpCol.MaxRows Then
            sSSQL = sSSQL & "   -- " & sName
        Else
            sSSQL = sSSQL & "     -- " & sName
        End If
        sSSQL = sSSQL & vbCrLf
        
        fpCol.Col = 5
        If Trim(fpCol.Value) <> "" Then
            sGSQL = sGSQL & vbCrLf
        End If
        
        
    Next i
    sISQL = sISQL & ")"
    sISQLValue = sISQLValue & ")"
    
    txtInsertSQL.Text = sISQL & " VALUES " & vbCrLf & sISQLValue
    
    
    txUpdateSQL.Text = sUSQL & sUSQLWhere
    
    txtSelectSQL.Text = sSSQL & "  FROM " & edtTable.Text & IIf(Trim(txtAlias.Text) <> "", "  " & Trim(txtAlias.Text), "") & "        /* " & edtTablenm.Text & "*/" & vbCrLf & sUSQLWhere
    
    txtGrpBySQL.Text = "SELECT " & sGSQL & "       , COUNT(*)  " & vbCrLf & " FROM " & edtTable.Text & vbCrLf & " GROUP BY " & sGSQL
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
   

    
    Screen.MousePointer = vbHourglass
    
    If GCONN_USE = False Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = GetConnString
        conn.Open
    End If
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient



 'SQL문 생성
    With fpCol
        
        sField = ""
        
        For i = 1 To .MaxRows
            
            .Row = i
            .Col = 1
            sCol = .Value
            sField = sField & sCol
            
            If i <> .MaxRows Then sField = sField & ","
            
            'VALUE 조회인경우
            sTemp = ""
            .Col = 9
            If Trim(.Value) <> "" Then
                
                .Col = 3
                If (.Value = "CHAR") Or (.Value = "VARCHAR") Or (.Value = "TIMESTMP") Then
                    .Col = 9
                    sTemp = sCol & " = '" & .Value & "'"
                Else
                    .Col = 9
                    sTemp = sCol & " = " & .Value
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
    
    If Trim(sReqSQL) = "" Then
        sSQL = "SELECT " & sField & " FROM " & edtTable.Text & sWhere
               
        If sWhere = "" Then
            sSQL = sSQL & " WHERE ROWNUM < 100"
        Else
            sSQL = sSQL & " AND ROWNUM < 100"
        End If
    Else
        sSQL = sReqSQL

    End If
    
    
 
    sSQL = sSQL & vbCrLf & " WITH UR "
    txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If
    
    
    fpSample.MaxCols = fpCol.MaxRows
    
    fpSample.MaxRows = 0
    
    '헤더설정
    
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
    Next i
    
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

End Sub












'------------------------------------------------------
' 테이블 상세정보 조회
'------------------------------------------------------
Private Sub subSelectTableDesc()

    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
                    

                        

                            
    Dim sField As String
    Dim sWhere As String
    Dim i As Integer
    Dim sCol As String
    Dim sTemp As String
   

    
    Screen.MousePointer = vbHourglass
    
    If GCONN_USE = False Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = GetConnString
        conn.Open
    End If
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient


    txtORGTAB.Text = ""
    txtCYCLE.Text = ""
    txtUSE_YN.Text = ""
    txtIMP_DESC.Text = ""
    txtCR_LOG.Text = ""
    fpInTab.MaxRows = 0
    fpOutTab.MaxRows = 0


    sSQL = ""
    sSQL = sSQL & vbCrLf & " SELECT NVL(ORGTAB, ' ') ORGTAB, NVL(CYCLE, ' ') CYCLE, NVL(USE_YN, ' ') USE_YN, NVL(IMP_DESC, ' ') IMP_DESC, NVL(CR_LOG, ' ')  CR_LOG"
    sSQL = sSQL & vbCrLf & "   FROM DIC_TAB_DESC "
    sSQL = sSQL & vbCrLf & "  WHERE OWNER  = '" & cboShema.Text & "'"
    sSQL = sSQL & vbCrLf & "    AND TABID  = '" & edtTable.Text & "' "

    txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If
    
    
    If Not rsTemp.EOF Then
        txtORGTAB.Text = rsTemp.Fields("ORGTAB").Value
        txtCYCLE.Text = rsTemp.Fields("CYCLE").Value
        txtUSE_YN.Text = rsTemp.Fields("USE_YN").Value
        txtIMP_DESC.Text = rsTemp.Fields("IMP_DESC").Value
        txtCR_LOG.Text = rsTemp.Fields("CR_LOG").Value
    End If
        
        
    
    
    
    sSQL = ""
    sSQL = sSQL & vbCrLf & " SELECT A.INTAB, B.TABNM "
    sSQL = sSQL & vbCrLf & "   FROM DIC_TAB_IN A LEFT OUTER JOIN DIC_TAB B "
    sSQL = sSQL & vbCrLf & "     ON A.INTAB  = B.TABID  "
    sSQL = sSQL & vbCrLf & "    AND A.OWNER  = B.OWNER  "
    sSQL = sSQL & vbCrLf & "  WHERE A.OWNER  = '" & cboShema.Text & "'"
    sSQL = sSQL & vbCrLf & "    AND A.TABID  = '" & edtTable.Text & "' "
    
    
    rsTemp.Close
    txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If
    
    txtSQL.Text = sSQL
    
    
    If rsTemp.RecordCount < 1 Then
        fpInTab.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpInTab, rsTemp)
    End If
    
        
        




    
    
    'OUT TAB 사용되어지는 테이블
    sSQL = ""
    sSQL = sSQL & vbCrLf & " SELECT A.TABID, B.TABNM"
    sSQL = sSQL & vbCrLf & "   FROM DIC_TAB_IN A LEFT OUTER JOIN DIC_TAB B "
    sSQL = sSQL & vbCrLf & "    ON A.TABID  = B.TABID  "
    sSQL = sSQL & vbCrLf & "    AND A.OWNER  = B.OWNER  "
    sSQL = sSQL & vbCrLf & "  WHERE A.OWNER  = '" & cboShema.Text & "'"
    sSQL = sSQL & vbCrLf & "    AND A.INTAB  = '" & edtTable.Text & "' "
    
    
    rsTemp.Close
    txtSQL.Text = sSQL
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If
    
    txtSQL.Text = sSQL
    
    
    If rsTemp.RecordCount < 1 Then
        fpOutTab.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpOutTab, rsTemp)
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
    sSQL = sSQL & vbCrLf & " DELETE FROM DIC_COL "
    sSQL = sSQL & vbCrLf & "  WHERE OWNER  = '" & cboShema.Text & "'                                                "
    sSQL = sSQL & vbCrLf & "    AND TABID  = '" & edtTable.Text & "'                                                "
    sSQL = sSQL & vbCrLf & "    AND COLID  = '" & txtColId.Text & "'                                                "
    
    If GCONN_USE = True Then
        gConn.Execute sSQL
    Else
        conn.Execute sSQL
    End If
    

    
    sSQL = ""
    sSQL = sSQL & vbCrLf & " INSERT INTO DIC_COL (OWNER, TABID, COLID, COLNM, CDDOM, BIGO1, BIGO2) "
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
    
    '테이블 레이아웃 재조회
    Call fpTable_DblClick(1, fpTable.Row)
    
    
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
'우클릭시 엑셀로 전환
Private Sub fpTable_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    fpTable.Row = 1
    fpTable.Col = 1
    If fpTable.Value <> "" Then
        PopupMenu Me.mnu_excel1, , fpTable.Left + MouseX + 100, fpTable.Top + MouseY
        gSelectGrid = "fpTable"
    End If
End Sub
'우클릭시 엑셀로 전환
Private Sub fpCol_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    fpCol.Row = 1
    fpCol.Col = 1
    If fpCol.Value <> "" Then
        PopupMenu Me.mnu_excel1, , SSTab3.Left + fpCol.Left + MouseX + 100, SSTab3.Top + fpCol.Top + MouseY
        gSelectGrid = "fpCol"
    End If
End Sub


Private Sub Image1_Click()
    frameCode.Visible = False
End Sub

Private Sub mnu_excel2_Click()
    If gSelectGrid = "fpTable" Then
        Call gfFpToExcel(fpTable)
    End If
    If gSelectGrid = "fpCol" Then
        Call gfFpToExcel(fpCol)
    End If
    If gSelectGrid = "fpSample" Then
        Call gfFpToExcel(fpSample)
    End If
    
End Sub

'***************************************************
' Insert sql 생성
'***************************************************
Private Sub mnu_in_Click()

End Sub

Private Sub SSTab3_DblClick()
    If UCase(Trim(txtGrpByCol.Text)) = "O" Then
        Frame3.Visible = True
    Else
        Frame3.Visible = False
    End If
End Sub

Private Sub SSTab5_Click(PreviousTab As Integer)

    If SSTab5.Tab = 0 Then txtPgmIDLog.SetFocus
    If SSTab5.Tab = 1 Then txtPgmIDLog2.SetFocus
    
End Sub

Private Sub txtEdDay_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call cmdJobLog_Click
    End If
End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call cmdJobLog_Click
    End If
End Sub

Private Sub txtPgmIDLog_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call cmdDebugLog_Click
    End If
End Sub

Private Sub txtPgmIDLog2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call cmdJobLog_Click
    End If
End Sub

Private Sub txtStDay_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call cmdJobLog_Click
    End If
End Sub

Private Sub txtTableSrchBigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call cmdTableSrch_Click
    End If
End Sub

Private Sub txtFindKey_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call cmdFindEmp_Click
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
