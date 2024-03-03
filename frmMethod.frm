VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMethod 
   BackColor       =   &H00ECD9DD&
   Caption         =   "AS-IS TO-BE Method 매핑"
   ClientHeight    =   13380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   25635
   Icon            =   "frmMethod.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13380
   ScaleWidth      =   25635
   Begin TabDlg.SSTab SSTab1 
      Height          =   13215
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   25515
      _ExtentX        =   45006
      _ExtentY        =   23310
      _Version        =   393216
      Tabs            =   7
      Tab             =   3
      TabsPerRow      =   10
      TabHeight       =   520
      TabCaption(0)   =   "AS-IS Method 매핑"
      TabPicture(0)   =   "frmMethod.frx":048A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtCondition"
      Tab(0).Control(1)=   "cmdDel"
      Tab(0).Control(2)=   "cmdSave"
      Tab(0).Control(3)=   "cmdAddRow"
      Tab(0).Control(4)=   "cmdSearch"
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(6)=   "fpMethod"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "결함관리"
      TabPicture(1)   =   "frmMethod.frx":04A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSTab2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "SQL 변환"
      TabPicture(2)   =   "frmMethod.frx":04C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "DMB매핑정의서 조회"
      TabPicture(3)   =   "frmMethod.frx":04DE
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label1(25)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label1(30)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "fpDMB"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cmdDMBSearch"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cmdSrchTable"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Command1"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Frame4"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "cmdSrchColumn"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "VO생성"
      TabPicture(4)   =   "frmMethod.frx":04FA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fpSrc"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "cmdExcelUpload"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "cmdMakeSrc"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "txtDirPath"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Frame5"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).ControlCount=   5
      TabCaption(5)   =   "보증번호 단위 이관"
      TabPicture(5)   =   "frmMethod.frx":0516
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label1(31)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Label1(32)"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Label1(33)"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "fpMig"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "txtGrntNo"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "cmdMig"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "optMig1"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "optMig2"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "txtColId"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "txtColVal"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).ControlCount=   10
      TabCaption(6)   =   "json"
      TabPicture(6)   =   "frmMethod.frx":0532
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Command9"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "txtFind"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Command8"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "Command7"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "fpJson"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).ControlCount=   5
      Begin VB.CommandButton Command9 
         BackColor       =   &H80000016&
         Caption         =   "BookMark Reset"
         Height          =   375
         Left            =   -64080
         Style           =   1  '그래픽
         TabIndex        =   149
         Top             =   1560
         Width           =   1875
      End
      Begin VB.TextBox txtFind 
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
         Left            =   -69720
         TabIndex        =   148
         Top             =   1680
         Width           =   5295
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H008080FF&
         Caption         =   "매핑"
         Height          =   960
         Left            =   -72000
         TabIndex        =   147
         Top             =   1080
         Width           =   1635
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H008080FF&
         Caption         =   "초기화"
         Height          =   960
         Left            =   -73800
         TabIndex        =   146
         Top             =   1080
         Width           =   1635
      End
      Begin VB.TextBox txtColVal 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   10  '한글 
         Left            =   -69720
         TabIndex        =   141
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtColId 
         Appearance      =   0  '평면
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   10  '한글 
         Left            =   -72960
         TabIndex        =   140
         Top             =   1320
         Width           =   2655
      End
      Begin VB.OptionButton optMig2 
         Height          =   255
         Left            =   -73560
         TabIndex        =   137
         Top             =   1080
         Width           =   495
      End
      Begin VB.OptionButton optMig1 
         Height          =   255
         Left            =   -73560
         TabIndex        =   136
         Top             =   720
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.CommandButton cmdMig 
         BackColor       =   &H008080FF&
         Caption         =   "이관"
         Height          =   960
         Left            =   -66240
         TabIndex        =   134
         Top             =   615
         Width           =   1635
      End
      Begin VB.TextBox txtGrntNo 
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
         Height          =   285
         IMEMode         =   10  '한글 
         Left            =   -72120
         TabIndex        =   133
         Top             =   690
         Width           =   5415
      End
      Begin VB.Frame Frame5 
         Height          =   12735
         Left            =   -74760
         TabIndex        =   123
         Top             =   380
         Width           =   24735
         Begin TabDlg.SSTab SSTab4 
            Height          =   6135
            Left            =   120
            TabIndex        =   127
            Top             =   6600
            Width           =   23775
            _ExtentX        =   41936
            _ExtentY        =   10821
            _Version        =   393216
            Tabs            =   2
            TabHeight       =   520
            TabCaption(0)   =   "VO"
            TabPicture(0)   =   "frmMethod.frx":054E
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "fpSQLVO"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Command5"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "DataCollection"
            TabPicture(1)   =   "frmMethod.frx":056A
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Command6"
            Tab(1).Control(1)=   "fpDataCollection"
            Tab(1).ControlCount=   2
            Begin VB.CommandButton Command6 
               Caption         =   "dataCollection copy"
               Height          =   420
               Left            =   -74640
               TabIndex        =   131
               Top             =   360
               Width           =   3015
            End
            Begin VB.CommandButton Command5 
               Caption         =   "Vo copy"
               Height          =   420
               Left            =   600
               TabIndex        =   129
               Top             =   120
               Width           =   1455
            End
            Begin FPSpreadADO.fpSpread fpSQLVO 
               Height          =   5355
               Left            =   120
               TabIndex        =   128
               Top             =   600
               Width           =   23115
               _Version        =   393216
               _ExtentX        =   40772
               _ExtentY        =   9446
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
               MaxCols         =   18
               MaxRows         =   1
               Protect         =   0   'False
               RetainSelBlock  =   0   'False
               ScrollBarExtMode=   -1  'True
               ShadowColor     =   12648447
               ShadowDark      =   12632256
               SpreadDesigner  =   "frmMethod.frx":0586
               ScrollBarTrack  =   1
            End
            Begin FPSpreadADO.fpSpread fpDataCollection 
               Height          =   4755
               Left            =   -74880
               TabIndex        =   130
               Top             =   840
               Width           =   23115
               _Version        =   393216
               _ExtentX        =   40772
               _ExtentY        =   8387
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
               MaxCols         =   3
               MaxRows         =   1
               Protect         =   0   'False
               RetainSelBlock  =   0   'False
               ScrollBarExtMode=   -1  'True
               ShadowColor     =   12648447
               ShadowDark      =   12632256
               SpreadDesigner  =   "frmMethod.frx":0B91
               ScrollBarTrack  =   1
            End
         End
         Begin VB.CommandButton cmdSQlClear 
            BackColor       =   &H008080FF&
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   20.25
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   15840
            TabIndex        =   126
            Top             =   360
            Width           =   3555
         End
         Begin VB.CommandButton cmdMakeVo 
            BackColor       =   &H008080FF&
            Caption         =   "VO생성"
            BeginProperty Font 
               Name            =   "굴림"
               Size            =   36
               Charset         =   129
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3375
            Left            =   15840
            TabIndex        =   125
            Top             =   2880
            Width           =   5475
         End
         Begin RichTextLib.RichTextBox txtVOSQL 
            Height          =   6195
            Left            =   240
            TabIndex        =   124
            Top             =   240
            Width           =   15435
            _ExtentX        =   27226
            _ExtentY        =   10927
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"frmMethod.frx":0F40
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
      Begin VB.TextBox cmdSrchColumn 
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
         Left            =   1440
         TabIndex        =   121
         Top             =   800
         Width           =   5295
      End
      Begin VB.TextBox txtDirPath 
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
         Left            =   -55800
         TabIndex        =   99
         Text            =   "C:\TEMP"
         Top             =   11180
         Visible         =   0   'False
         Width           =   6795
      End
      Begin VB.CommandButton cmdMakeSrc 
         Caption         =   "소스 생성"
         Height          =   855
         Left            =   -53280
         TabIndex        =   98
         Top             =   6620
         Visible         =   0   'False
         Width           =   3315
      End
      Begin VB.CommandButton cmdExcelUpload 
         Caption         =   "Excel 읽기"
         Height          =   855
         Left            =   -56880
         TabIndex        =   96
         Top             =   6620
         Visible         =   0   'False
         Width           =   3315
      End
      Begin VB.Frame Frame4 
         Caption         =   "결과 내 검색"
         Height          =   675
         Left            =   360
         TabIndex        =   82
         Top             =   1340
         Width           =   10935
         Begin VB.CommandButton Command3 
            BackColor       =   &H80000016&
            Caption         =   "BookMark Reset"
            Height          =   375
            Left            =   5760
            Style           =   1  '그래픽
            TabIndex        =   84
            Top             =   180
            Width           =   1875
         End
         Begin VB.TextBox txtColumnSrch 
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
            Left            =   180
            TabIndex        =   83
            Top             =   240
            Width           =   5295
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Excel"
         Height          =   675
         Left            =   8640
         TabIndex        =   81
         Top             =   500
         Width           =   1935
      End
      Begin VB.TextBox cmdSrchTable 
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
         Left            =   1500
         TabIndex        =   80
         Top             =   440
         Width           =   5295
      End
      Begin VB.CommandButton cmdDMBSearch 
         BackColor       =   &H008080FF&
         Caption         =   "조회"
         Height          =   735
         Left            =   6960
         TabIndex        =   79
         Top             =   440
         Width           =   1635
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   12675
         Left            =   -74940
         TabIndex        =   68
         Top             =   440
         Width           =   25275
         _ExtentX        =   44582
         _ExtentY        =   22357
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "SQL 변환"
         TabPicture(0)   =   "frmMethod.frx":0FE5
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "txtNewSQL"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtOldSQL"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdClear"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "cmdChange"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cmdTxtClear"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cmdMinMax(1)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "cmdMinMax(0)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Command4"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "fpDup"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).ControlCount=   9
         TabCaption(1)   =   "매핑정의"
         TabPicture(1)   =   "frmMethod.frx":1001
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fpMap"
         Tab(1).Control(1)=   "txtCodeSrch"
         Tab(1).Control(2)=   "cmdExcel"
         Tab(1).Control(3)=   "cmdBookMakr"
         Tab(1).ControlCount=   4
         TabCaption(2)   =   "TableMaing"
         TabPicture(2)   =   "frmMethod.frx":101D
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtTableMap"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "주석적용"
         TabPicture(3)   =   "frmMethod.frx":1039
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "txtNewComment"
         Tab(3).Control(1)=   "txtOldComment"
         Tab(3).Control(2)=   "cmdCommentChange"
         Tab(3).ControlCount=   3
         Begin VB.CommandButton cmdCommentChange 
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
            Height          =   2895
            Left            =   -62640
            TabIndex        =   144
            Top             =   4440
            Width           =   675
         End
         Begin FPSpreadADO.fpSpread fpDup 
            Height          =   11175
            Left            =   11880
            TabIndex        =   92
            Top             =   11220
            Width           =   13275
            _Version        =   393216
            _ExtentX        =   23416
            _ExtentY        =   19711
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
            MaxCols         =   8
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            SelectBlockOptions=   0
            ShadowColor     =   12648447
            ShadowDark      =   12632256
            SpreadDesigner  =   "frmMethod.frx":1055
            ScrollBarTrack  =   1
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Excel"
            Height          =   375
            Left            =   10620
            TabIndex        =   95
            Top             =   12180
            Width           =   1155
         End
         Begin VB.CommandButton cmdMinMax 
            Caption         =   "↑"
            Height          =   315
            Index           =   0
            Left            =   11040
            Style           =   1  '그래픽
            TabIndex        =   94
            Top             =   11280
            Width           =   735
         End
         Begin VB.CommandButton cmdMinMax 
            Caption         =   "↓"
            Height          =   315
            Index           =   1
            Left            =   11040
            Style           =   1  '그래픽
            TabIndex        =   93
            Top             =   11700
            Width           =   735
         End
         Begin VB.CommandButton cmdBookMakr 
            BackColor       =   &H80000016&
            Caption         =   "BookMark Reset"
            Height          =   555
            Left            =   -66960
            Style           =   1  '그래픽
            TabIndex        =   77
            Top             =   660
            Width           =   1875
         End
         Begin VB.CommandButton cmdExcel 
            Caption         =   "Excel"
            Height          =   555
            Left            =   -69000
            TabIndex        =   76
            Top             =   660
            Width           =   1935
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
            IMEMode         =   8  '영문
            Left            =   -74520
            TabIndex        =   75
            Top             =   780
            Width           =   5175
         End
         Begin VB.CommandButton cmdTxtClear 
            Caption         =   "Celar"
            Height          =   615
            Left            =   1320
            TabIndex        =   74
            Top             =   11220
            Width           =   2955
         End
         Begin VB.CommandButton cmdChange 
            Caption         =   "2.변환"
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
            Left            =   7740
            TabIndex        =   72
            Top             =   11220
            Width           =   2955
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "1.문자열 정리"
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
            Left            =   4500
            TabIndex        =   71
            Top             =   11220
            Width           =   2955
         End
         Begin RichTextLib.RichTextBox txtOldSQL 
            Height          =   10635
            Left            =   120
            TabIndex        =   69
            Top             =   480
            Width           =   12555
            _ExtentX        =   22146
            _ExtentY        =   18759
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"frmMethod.frx":15B4
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
         Begin RichTextLib.RichTextBox txtNewSQL 
            Height          =   10635
            Left            =   12780
            TabIndex        =   70
            Top             =   480
            Width           =   12435
            _ExtentX        =   21934
            _ExtentY        =   18759
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"frmMethod.frx":1659
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
         Begin FPSpreadADO.fpSpread fpMap 
            Height          =   11115
            Left            =   -74700
            TabIndex        =   73
            Top             =   1320
            Width           =   24315
            _Version        =   393216
            _ExtentX        =   42889
            _ExtentY        =   19606
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
            MaxCols         =   8
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            SelectBlockOptions=   0
            ShadowColor     =   12648447
            ShadowDark      =   12632256
            SpreadDesigner  =   "frmMethod.frx":16FE
            ScrollBarTrack  =   1
         End
         Begin RichTextLib.RichTextBox txtTableMap 
            Height          =   7635
            Left            =   -74880
            TabIndex        =   116
            Top             =   480
            Width           =   8355
            _ExtentX        =   14737
            _ExtentY        =   13467
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"frmMethod.frx":1C5D
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
         Begin RichTextLib.RichTextBox txtOldComment 
            Height          =   11595
            Left            =   -74880
            TabIndex        =   142
            Top             =   360
            Width           =   12195
            _ExtentX        =   21511
            _ExtentY        =   20452
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"frmMethod.frx":1D02
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
         Begin RichTextLib.RichTextBox txtNewComment 
            Height          =   11595
            Left            =   -61920
            TabIndex        =   143
            Top             =   360
            Width           =   12075
            _ExtentX        =   21299
            _ExtentY        =   20452
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"frmMethod.frx":1DA7
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
         Caption         =   "조치이력"
         Height          =   8295
         Left            =   -65340
         TabIndex        =   51
         Top             =   4820
         Width           =   15675
         Begin VB.CommandButton cmdNameChange 
            BackColor       =   &H008080FF&
            Caption         =   "조치자변경"
            Height          =   495
            Left            =   13800
            TabIndex        =   119
            Top             =   7680
            Width           =   1635
         End
         Begin RichTextLib.RichTextBox rtxtBug 
            Height          =   3420
            Index           =   7
            Left            =   240
            TabIndex        =   114
            Top             =   4200
            Width           =   15375
            _ExtentX        =   27120
            _ExtentY        =   6033
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmMethod.frx":1E4C
         End
         Begin VB.CommandButton cmdTempSave 
            BackColor       =   &H008080FF&
            Caption         =   "임시저장"
            Height          =   495
            Left            =   9240
            TabIndex        =   105
            Top             =   7680
            Width           =   1635
         End
         Begin VB.TextBox txtKey 
            Height          =   510
            Left            =   14640
            Locked          =   -1  'True
            TabIndex        =   104
            Top             =   6840
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cmdConfirm 
            BackColor       =   &H008080FF&
            Caption         =   "확인완료"
            Enabled         =   0   'False
            Height          =   495
            Left            =   5520
            TabIndex        =   59
            Top             =   7740
            Width           =   1635
         End
         Begin VB.CommandButton cmdReject 
            BackColor       =   &H008080FF&
            Caption         =   "반려"
            Enabled         =   0   'False
            Height          =   495
            Left            =   3780
            TabIndex        =   56
            Top             =   7740
            Width           =   1635
         End
         Begin VB.CommandButton cmdProcComp 
            BackColor       =   &H008080FF&
            Caption         =   "조치완료"
            Enabled         =   0   'False
            Height          =   495
            Left            =   2040
            TabIndex        =   55
            Top             =   7740
            Width           =   1635
         End
         Begin VB.CommandButton cmdNewInput 
            BackColor       =   &H008080FF&
            Caption         =   "신규입력"
            Height          =   495
            Left            =   300
            TabIndex        =   54
            Top             =   7740
            Width           =   1635
         End
         Begin FPSpreadADO.fpSpread fpProcHist 
            Height          =   3555
            Left            =   180
            TabIndex        =   52
            Top             =   300
            Width           =   15375
            _Version        =   393216
            _ExtentX        =   27120
            _ExtentY        =   6271
            _StockProps     =   64
            ColsFrozen      =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   15522269
            MaxCols         =   5
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "frmMethod.frx":1ED0
         End
         Begin RichTextLib.RichTextBox txtName 
            Height          =   300
            Left            =   12480
            TabIndex        =   120
            Top             =   7800
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393217
            Enabled         =   -1  'True
            MultiLine       =   0   'False
            Appearance      =   0
            TextRTF         =   $"frmMethod.frx":2331
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "조치의견"
            Height          =   195
            Index           =   21
            Left            =   240
            TabIndex        =   53
            Top             =   3960
            Width           =   4395
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "결함 등록"
         Height          =   4155
         Left            =   -65340
         TabIndex        =   42
         Top             =   560
         Width           =   15675
         Begin VB.CheckBox chkUI 
            Caption         =   "UI검토의견"
            Height          =   255
            Left            =   4080
            TabIndex        =   117
            Top             =   200
            Width           =   2655
         End
         Begin RichTextLib.RichTextBox rtxtBug 
            Height          =   300
            Index           =   8
            Left            =   1200
            TabIndex        =   115
            Top             =   3480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393217
            BackColor       =   14737632
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            Appearance      =   0
            TextRTF         =   $"frmMethod.frx":23CE
         End
         Begin RichTextLib.RichTextBox rtxtBug 
            Height          =   2580
            Index           =   6
            Left            =   3960
            TabIndex        =   113
            Top             =   480
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   4551
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmMethod.frx":2452
         End
         Begin RichTextLib.RichTextBox rtxtBug 
            Height          =   300
            Index           =   5
            Left            =   1200
            TabIndex        =   112
            Top             =   3120
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   529
            _Version        =   393217
            Enabled         =   -1  'True
            MultiLine       =   0   'False
            Appearance      =   0
            TextRTF         =   $"frmMethod.frx":24D6
         End
         Begin RichTextLib.RichTextBox rtxtBug 
            Height          =   300
            Index           =   4
            Left            =   1200
            TabIndex        =   111
            Top             =   1680
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   529
            _Version        =   393217
            BackColor       =   14737632
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            Appearance      =   0
            TextRTF         =   $"frmMethod.frx":255A
         End
         Begin RichTextLib.RichTextBox rtxtBug 
            Height          =   300
            Index           =   3
            Left            =   1200
            TabIndex        =   110
            Top             =   1320
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   529
            _Version        =   393217
            BackColor       =   14737632
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            Appearance      =   0
            TextRTF         =   $"frmMethod.frx":25DE
         End
         Begin RichTextLib.RichTextBox rtxtBug 
            Height          =   300
            Index           =   2
            Left            =   1200
            TabIndex        =   109
            Top             =   960
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   529
            _Version        =   393217
            Enabled         =   -1  'True
            MultiLine       =   0   'False
            Appearance      =   0
            TextRTF         =   $"frmMethod.frx":2662
         End
         Begin RichTextLib.RichTextBox rtxtBug 
            Height          =   300
            Index           =   1
            Left            =   1200
            TabIndex        =   108
            Top             =   600
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   529
            _Version        =   393217
            Enabled         =   -1  'True
            MultiLine       =   0   'False
            Appearance      =   0
            TextRTF         =   $"frmMethod.frx":26E6
         End
         Begin VB.TextBox txtFileServer 
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
            Height          =   480
            IMEMode         =   10  '한글 
            Index           =   9
            Left            =   2580
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   67
            Text            =   "frmMethod.frx":276A
            Top             =   3480
            Width           =   7875
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H008080FF&
            Caption         =   "신규입력"
            Height          =   735
            Left            =   10920
            TabIndex        =   58
            Top             =   3180
            Width           =   1635
         End
         Begin VB.CommandButton cmdBugReg 
            BackColor       =   &H008080FF&
            Caption         =   "결함등록"
            Height          =   795
            Left            =   12660
            TabIndex        =   50
            Top             =   3120
            Width           =   2895
         End
         Begin RichTextLib.RichTextBox rtxtBug 
            Height          =   300
            Index           =   0
            Left            =   1200
            TabIndex        =   107
            Top             =   240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   529
            _Version        =   393217
            Enabled         =   -1  'True
            MultiLine       =   0   'False
            Appearance      =   0
            TextRTF         =   $"frmMethod.frx":27F1
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "No"
            Height          =   195
            Index           =   22
            Left            =   240
            TabIndex        =   57
            Top             =   3540
            Width           =   4395
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "관련파일"
            Height          =   195
            Index           =   20
            Left            =   240
            TabIndex        =   49
            Top             =   3180
            Width           =   4395
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "상태"
            Height          =   195
            Index           =   19
            Left            =   240
            TabIndex        =   48
            Top             =   1740
            Width           =   4395
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "결함내용"
            Height          =   195
            Index           =   18
            Left            =   3120
            TabIndex        =   47
            Top             =   2160
            Width           =   1155
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "등록자"
            Height          =   195
            Index           =   17
            Left            =   240
            TabIndex        =   46
            Top             =   1380
            Width           =   4395
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "개발자"
            Height          =   195
            Index           =   16
            Left            =   240
            TabIndex        =   45
            Top             =   1020
            Width           =   4395
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "프로그램명"
            Height          =   195
            Index           =   15
            Left            =   240
            TabIndex        =   44
            Top             =   660
            Width           =   4395
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "프로그램ID"
            Height          =   195
            Index           =   14
            Left            =   240
            TabIndex        =   43
            Top             =   300
            Width           =   4395
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   12675
         Left            =   -74880
         TabIndex        =   40
         Top             =   440
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   22357
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "결함목록"
         TabPicture(0)   =   "frmMethod.frx":2875
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1(23)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label1(24)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label1(28)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label1(29)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "fpBugList"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cmdSearch2"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "txtUserNm"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "optStatus(0)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "optStatus(1)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "optStatus(2)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "txtPGMid"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txtPGMNM"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "optStatus(3)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "chkSelUI"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).ControlCount=   14
         TabCaption(1)   =   "개발목록"
         TabPicture(1)   =   "frmMethod.frx":2891
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cdmDevSearch"
         Tab(1).Control(1)=   "txtDevPgm"
         Tab(1).Control(2)=   "txtDevName"
         Tab(1).Control(3)=   "fpDevPgm"
         Tab(1).Control(4)=   "Label1(27)"
         Tab(1).Control(5)=   "Label1(26)"
         Tab(1).ControlCount=   6
         Begin VB.CheckBox chkSelUI 
            Caption         =   "UI검토의견"
            Height          =   255
            Left            =   5040
            TabIndex        =   118
            Top             =   1080
            Width           =   2655
         End
         Begin VB.OptionButton optStatus 
            Caption         =   "조치완료"
            Height          =   195
            Index           =   3
            Left            =   2160
            TabIndex        =   106
            Top             =   1080
            Width           =   1155
         End
         Begin VB.TextBox txtPGMNM 
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
            Left            =   4080
            TabIndex        =   103
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox txtPGMid 
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
            Left            =   4080
            TabIndex        =   101
            Top             =   360
            Width           =   2895
         End
         Begin VB.CommandButton cdmDevSearch 
            BackColor       =   &H008080FF&
            Caption         =   "조회"
            Height          =   495
            Left            =   -67980
            TabIndex        =   91
            Top             =   600
            Width           =   1635
         End
         Begin VB.TextBox txtDevPgm 
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
            Left            =   -73380
            TabIndex        =   90
            Top             =   780
            Width           =   5115
         End
         Begin VB.TextBox txtDevName 
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
            Left            =   -73380
            TabIndex        =   87
            Top             =   420
            Width           =   1575
         End
         Begin VB.OptionButton optStatus 
            Caption         =   "확인완료"
            Height          =   195
            Index           =   2
            Left            =   3300
            TabIndex        =   66
            Top             =   1080
            Width           =   1155
         End
         Begin VB.OptionButton optStatus 
            Caption         =   "진행중"
            Height          =   195
            Index           =   1
            Left            =   1200
            TabIndex        =   65
            Top             =   1080
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton optStatus 
            Caption         =   "전체"
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   64
            Top             =   1080
            Width           =   690
         End
         Begin VB.TextBox txtUserNm 
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
            Left            =   1260
            TabIndex        =   61
            Top             =   480
            Width           =   1575
         End
         Begin VB.CommandButton cmdSearch2 
            BackColor       =   &H008080FF&
            Caption         =   "조회"
            Height          =   855
            Left            =   7740
            TabIndex        =   60
            Top             =   420
            Width           =   1395
         End
         Begin FPSpreadADO.fpSpread fpBugList 
            Height          =   11115
            Left            =   60
            TabIndex        =   41
            Top             =   1380
            Width           =   9315
            _Version        =   393216
            _ExtentX        =   16431
            _ExtentY        =   19606
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ColsFrozen      =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GrayAreaBackColor=   15522269
            MaxCols         =   9
            MaxRows         =   1
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "frmMethod.frx":28AD
         End
         Begin FPSpreadADO.fpSpread fpDevPgm 
            Height          =   11115
            Left            =   -74880
            TabIndex        =   86
            Top             =   1260
            Width           =   8955
            _Version        =   393216
            _ExtentX        =   15796
            _ExtentY        =   19606
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
            MaxCols         =   15
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            SelectBlockOptions=   0
            ShadowColor     =   12648447
            ShadowDark      =   12632256
            SpreadDesigner  =   "frmMethod.frx":2EC6
            ScrollBarTrack  =   1
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "프로그램명"
            Height          =   195
            Index           =   29
            Left            =   3120
            TabIndex        =   102
            Top             =   780
            Width           =   1515
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "프로그램ID"
            Height          =   195
            Index           =   28
            Left            =   3120
            TabIndex        =   100
            Top             =   420
            Width           =   1515
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "프로그램"
            Height          =   195
            Index           =   27
            Left            =   -74820
            TabIndex        =   89
            Top             =   780
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "개발자"
            Height          =   195
            Index           =   26
            Left            =   -74820
            TabIndex        =   88
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "상태"
            Height          =   195
            Index           =   24
            Left            =   60
            TabIndex        =   63
            Top             =   1080
            Width           =   1515
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "등록자/개발자"
            Height          =   195
            Index           =   23
            Left            =   60
            TabIndex        =   62
            Top             =   540
            Width           =   1515
         End
      End
      Begin VB.TextBox txtCondition 
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
         Left            =   -74820
         TabIndex        =   22
         Top             =   800
         Width           =   9315
      End
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H008080FF&
         Caption         =   "삭제"
         Height          =   495
         Left            =   -58380
         TabIndex        =   6
         Top             =   680
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H008080FF&
         Caption         =   "저장"
         Height          =   495
         Left            =   -59880
         TabIndex        =   5
         Top             =   680
         Width           =   1455
      End
      Begin VB.CommandButton cmdAddRow 
         BackColor       =   &H008080FF&
         Caption         =   "행추가"
         Height          =   495
         Left            =   -61560
         TabIndex        =   4
         Top             =   680
         Width           =   1635
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H008080FF&
         Caption         =   "조회"
         Height          =   495
         Left            =   -65400
         TabIndex        =   3
         Top             =   680
         Width           =   1635
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00ECD9DD&
         Height          =   12015
         Left            =   -56400
         TabIndex        =   2
         Top             =   1100
         Width           =   6795
         Begin VB.CommandButton cmdInsert 
            BackColor       =   &H008080FF&
            Caption         =   "신규등록"
            Height          =   375
            Left            =   1320
            TabIndex        =   39
            Top             =   240
            Width           =   1395
         End
         Begin VB.CommandButton cmdNew 
            BackColor       =   &H008080FF&
            Caption         =   "신규 입력"
            Height          =   495
            Left            =   3600
            TabIndex        =   38
            Top             =   10620
            Width           =   1455
         End
         Begin VB.CommandButton cmdSingleSave 
            BackColor       =   &H008080FF&
            Caption         =   "단건 저장"
            Height          =   495
            Left            =   5100
            TabIndex        =   37
            Top             =   10620
            Width           =   1455
         End
         Begin VB.TextBox txtData 
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
            Index           =   14
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   21
            Top             =   9840
            Width           =   6615
         End
         Begin VB.TextBox txtData 
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
            Height          =   540
            IMEMode         =   10  '한글 
            Index           =   13
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   20
            Top             =   9000
            Width           =   6615
         End
         Begin VB.TextBox txtData 
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
            Index           =   12
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   19
            Top             =   8400
            Width           =   6615
         End
         Begin VB.TextBox txtData 
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
            Index           =   11
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   18
            Top             =   7740
            Width           =   6615
         End
         Begin VB.TextBox txtData 
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
            Height          =   345
            IMEMode         =   10  '한글 
            Index           =   10
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   17
            Top             =   7080
            Width           =   6615
         End
         Begin VB.TextBox txtData 
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
            Index           =   9
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   16
            Top             =   6480
            Width           =   6615
         End
         Begin VB.TextBox txtData 
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
            Index           =   8
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   15
            Top             =   5880
            Width           =   6615
         End
         Begin VB.TextBox txtData 
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
            Index           =   7
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   14
            Top             =   5220
            Width           =   6615
         End
         Begin VB.TextBox txtData 
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
            Height          =   600
            IMEMode         =   10  '한글 
            Index           =   6
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   13
            Top             =   4260
            Width           =   6615
         End
         Begin VB.TextBox txtData 
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
            Index           =   5
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   3540
            Width           =   6615
         End
         Begin VB.TextBox txtData 
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
            Height          =   540
            IMEMode         =   10  '한글 
            Index           =   4
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   2640
            Width           =   6615
         End
         Begin VB.TextBox txtData 
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
            Index           =   3
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   10
            Top             =   2040
            Width           =   6615
         End
         Begin VB.TextBox txtData 
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
            Index           =   2
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   9
            Top             =   1440
            Width           =   6615
         End
         Begin VB.TextBox txtData 
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
            Index           =   1
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   840
            Width           =   6615
         End
         Begin VB.TextBox txtData 
            Appearance      =   0  '평면
            BackColor       =   &H8000000A&
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
            Index           =   0
            Left            =   60
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   7
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "이관구분"
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   36
            Top             =   9660
            Width           =   4395
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "비고"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   35
            Top             =   8760
            Width           =   4395
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "TO-BE Biz Method 명"
            Height          =   195
            Index           =   11
            Left            =   60
            TabIndex        =   34
            Top             =   8220
            Width           =   4395
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "TO-BE Biz Method"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   33
            Top             =   7560
            Width           =   4395
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "TO-BE Biz 업무구분"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   32
            Top             =   6885
            Width           =   4395
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "TO-BE Biz 한글명"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   31
            Top             =   6240
            Width           =   4395
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "TOBE Biz 영문명"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   30
            Top             =   5640
            Width           =   4395
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "프로그램 구분"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   29
            Top             =   4980
            Width           =   4395
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "AS-IS METHOD 설명"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   28
            Top             =   4020
            Width           =   4395
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "AS-IS METHOD"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   27
            Top             =   3300
            Width           =   4395
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "BEAN 클래스설명"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   26
            Top             =   2460
            Width           =   4395
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "AS-IS Bean"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   25
            Top             =   1860
            Width           =   4395
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "AS-IS 하위그룹"
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   24
            Top             =   1260
            Width           =   4395
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '투명
            Caption         =   "As-Is Package"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   23
            Top             =   660
            Width           =   4395
         End
      End
      Begin FPSpreadADO.fpSpread fpMethod 
         Height          =   11895
         Left            =   -74880
         TabIndex        =   1
         Top             =   1215
         Width           =   18375
         _Version        =   393216
         _ExtentX        =   32411
         _ExtentY        =   20981
         _StockProps     =   64
         ColsFrozen      =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   15522269
         MaxCols         =   15
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "frmMethod.frx":3600
      End
      Begin FPSpreadADO.fpSpread fpDMB 
         Height          =   10395
         Left            =   240
         TabIndex        =   78
         Top             =   2120
         Width           =   24315
         _Version        =   393216
         _ExtentX        =   42889
         _ExtentY        =   18336
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ColsFrozen      =   6
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
         MaxCols         =   24
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ShadowColor     =   12648447
         ShadowDark      =   12632256
         SpreadDesigner  =   "frmMethod.frx":3D7B
         ScrollBarTrack  =   1
      End
      Begin FPSpreadADO.fpSpread fpSrc 
         Height          =   2415
         Left            =   -56400
         TabIndex        =   97
         Top             =   8180
         Visible         =   0   'False
         Width           =   4515
         _Version        =   393216
         _ExtentX        =   7964
         _ExtentY        =   4260
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxRows         =   10000
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "frmMethod.frx":46E0
      End
      Begin FPSpreadADO.fpSpread fpMig 
         Height          =   11355
         Left            =   -74760
         TabIndex        =   132
         Top             =   1695
         Width           =   20955
         _Version        =   393216
         _ExtentX        =   36962
         _ExtentY        =   20029
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
         MaxCols         =   3
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ShadowColor     =   12648447
         ShadowDark      =   12632256
         SpreadDesigner  =   "frmMethod.frx":223D6
         ScrollBarTrack  =   1
      End
      Begin FPSpreadADO.fpSpread fpJson 
         Height          =   8835
         Left            =   -74760
         TabIndex        =   145
         Top             =   2160
         Width           =   13155
         _Version        =   393216
         _ExtentX        =   23204
         _ExtentY        =   15584
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
         MaxCols         =   3
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ShadowColor     =   12648447
         ShadowDark      =   12632256
         SpreadDesigner  =   "frmMethod.frx":22785
         ScrollBarTrack  =   1
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "이행값"
         Height          =   195
         Index           =   33
         Left            =   -69600
         TabIndex        =   139
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "컬럼명"
         Height          =   195
         Index           =   32
         Left            =   -72960
         TabIndex        =   138
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "보증번호"
         Height          =   195
         Index           =   31
         Left            =   -72960
         TabIndex        =   135
         Top             =   735
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "컬럼ID"
         Height          =   195
         Index           =   30
         Left            =   480
         TabIndex        =   122
         Top             =   860
         Width           =   1515
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '투명
         Caption         =   "테이블명/id"
         Height          =   195
         Index           =   25
         Left            =   420
         TabIndex        =   85
         Top             =   500
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cdmDevSearch_Click()
    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    

    Set conn = New ADODB.Connection
    conn.ConnectionString = GetConnStringHgDev

    conn.Open

    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
            
        sSQL = sSQL & vbCrLf & "SELECT 업무구분         "
        sSQL = sSQL & vbCrLf & "        ,화면ID          "
        sSQL = sSQL & vbCrLf & "        ,화면명          "
        sSQL = sSQL & vbCrLf & "        ,개발자          "
        sSQL = sSQL & vbCrLf & "        ,보고서ID        "
        sSQL = sSQL & vbCrLf & "        ,보고서명        "
        sSQL = sSQL & vbCrLf & "        ,프로그램ID      "
        sSQL = sSQL & vbCrLf & "        ,프로그램명      "
        sSQL = sSQL & vbCrLf & "        ,프로그램구분    "
        sSQL = sSQL & vbCrLf & "        ,단위테스트ID    "
        sSQL = sSQL & vbCrLf & "        ,프로그램유형    "
        sSQL = sSQL & vbCrLf & "        ,리스크구분      "
        sSQL = sSQL & vbCrLf & "        ,난이도          "
        sSQL = sSQL & vbCrLf & "        ,개발계획시작일  "
        sSQL = sSQL & vbCrLf & "        ,개발계획종료일  "
        sSQL = sSQL & vbCrLf & "  FROM HGDEV_T008     "
        sSQL = sSQL & vbCrLf & " WHERE 1=1"

        If Trim(txtDevName.Text) <> "" Then
            sSQL = sSQL & vbCrLf & "   AND 개발자 LIKE  '%" & Trim(txtDevName.Text) & "%'          "
        End If
        
        If Trim(txtDevPgm.Text) <> "" Then
            sSQL = sSQL & vbCrLf & "   AND (  화면ID LIKE  '%" & Trim(txtDevPgm.Text) & "%'          "
            sSQL = sSQL & vbCrLf & "       OR 화면명 LIKE  '%" & Trim(txtDevPgm.Text) & "%'          "
            sSQL = sSQL & vbCrLf & "       OR 보고서ID LIKE  '%" & Trim(txtDevPgm.Text) & "%'          "
            sSQL = sSQL & vbCrLf & "       OR 보고서명 LIKE  '%" & Trim(txtDevPgm.Text) & "%'          "
            sSQL = sSQL & vbCrLf & "       OR 프로그램ID LIKE  '%" & Trim(txtDevPgm.Text) & "%'          "
            sSQL = sSQL & vbCrLf & "       OR 프로그램명 LIKE  '%" & Trim(txtDevPgm.Text) & "%' )          "
        End If
        
        


    rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

    fpDevPgm.MaxRows = 0
    If rsTemp.RecordCount < 1 Then
        fpDevPgm.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpDevPgm, rsTemp)
    End If
    
    
    Dim i As Integer
    For i = 1 To fpDevPgm.MaxRows
        fpDevPgm.RowHeight(i) = 15
    Next i
    
    
    fpDevPgm.Row = -1
    fpDevPgm.Col = -1
    fpDevPgm.Lock = True
    
    
    'makeCellBorderline
    
        
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

Private Sub chkUI_Click()
    If chkUI.Value = "1" Then
        rtxtBug(6).Text = "<< UI 검토의견 >> " & vbCrLf
        rtxtBug(8).Text = ""
    End If
    
End Sub

Private Sub cmdAddRow_Click()
    fpMethod.MaxRows = fpMethod.MaxRows + 1
    fpMethod.InsertRows fpMethod.ActiveRow + 1, 1
    fpMethod.Col = -1
    fpMethod.Row = fpMethod.ActiveRow + 1
    fpMethod.BackColor = C_ADD_COLOR
    

    makeCellBorderline
    
End Sub

Private Sub cmdBookMakr_Click()
    Dim i As Integer
    For i = 1 To fpMap.MaxRows
        fpMap.Col = -1
        fpMap.Row = i
        fpMap.BackColor = S_WHITE
    Next i
End Sub

Private Sub cmdBugReg_Click()

    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    
    

    Set conn = New ADODB.Connection
    conn.ConnectionString = GetConnStringHgDev

    conn.Open
    conn.BeginTrans
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
    Dim i As Integer
    
    If gfMsgQst("저장하시겠습니까?") = vbNo Then Exit Sub

                
    Screen.MousePointer = vbHourglass
    
                '삭제하고 insert
                If Trim(rtxtBug(8).Text) <> "" Then
                    
                    sSQL = ""
                    sSQL = sSQL & vbCrLf & " DELETE FROM HGDEV.HGDEV_T003 "
                    sSQL = sSQL & vbCrLf & "  WHERE NO  = " & rtxtBug(8).Text & ""
        
                    conn.Execute sSQL
                                
                '채번하기
                Else
                    sSQL = ""
                    sSQL = sSQL & vbCrLf & " SELECT nvl(MAX(NO),0) + 1 FROM  HGDEV.HGDEV_T003  "
                    
                    rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
                    
                    
                    Dim sNo As String
                    sNo = rsTemp.Fields(0).Value
                    rtxtBug(8).Text = sNo
                    
                    
                End If
                
                '등록하기
                sSQL = ""
                sSQL = sSQL & vbCrLf & "INSERT INTO HGDEV_T003 (         /* */"
                sSQL = sSQL & vbCrLf & "          NO                 -- "
                sSQL = sSQL & vbCrLf & "         ,프로그램ID             -- "
                sSQL = sSQL & vbCrLf & "         ,프로그램명              -- "
                sSQL = sSQL & vbCrLf & "         ,등록자                -- "
                sSQL = sSQL & vbCrLf & "         ,개발자                -- "
                sSQL = sSQL & vbCrLf & "         ,상태                 -- "
                sSQL = sSQL & vbCrLf & "         ,결함내용               -- "
                sSQL = sSQL & vbCrLf & "         ,파일명                -- "
                sSQL = sSQL & vbCrLf & "         ,등록일시                 -- "
                sSQL = sSQL & vbCrLf & "         ,UI여부                 -- "
                sSQL = sSQL & vbCrLf & ") VALUES "
                sSQL = sSQL & vbCrLf & " (        " & rtxtBug(8).Text & ""
                sSQL = sSQL & vbCrLf & "        ,'" & rtxtBug(0).Text & "'          "
                sSQL = sSQL & vbCrLf & "        ,'" & rtxtBug(1).Text & "'          "
                sSQL = sSQL & vbCrLf & "        ,'" & rtxtBug(3).Text & "'          "
                sSQL = sSQL & vbCrLf & "        ,'" & rtxtBug(2).Text & "'          "
                sSQL = sSQL & vbCrLf & "        ,'" & "등록" & "'          "
                sSQL = sSQL & vbCrLf & "        ,'" & Replace(rtxtBug(6).Text, "'", "''") & "'          "
                sSQL = sSQL & vbCrLf & "        ,'" & Replace(rtxtBug(5).Text, "'", "''") & "'          "
                sSQL = sSQL & vbCrLf & "        ,sysdate          "
                sSQL = sSQL & vbCrLf & "        ,'" & chkUI.Value & "'          "
                sSQL = sSQL & vbCrLf & ")"

                conn.Execute sSQL
    
                sSQL = ""
                sSQL = sSQL & vbCrLf & "INSERT INTO HGDEV_T004 (         "
                sSQL = sSQL & vbCrLf & "          NO               "
                sSQL = sSQL & vbCrLf & "         ,등록일시         "
                sSQL = sSQL & vbCrLf & "         ,등록자           "
                sSQL = sSQL & vbCrLf & "         ,상태             "
                sSQL = sSQL & vbCrLf & "         ,조치의견         "
                sSQL = sSQL & vbCrLf & ") VALUES "
                sSQL = sSQL & vbCrLf & " (        " & rtxtBug(8).Text & ""
                sSQL = sSQL & vbCrLf & "        ,sysdate               "
                sSQL = sSQL & vbCrLf & "        ,'" & gUserNm & "'          "
                sSQL = sSQL & vbCrLf & "        ,'결함등록(수정)'           "
                sSQL = sSQL & vbCrLf & "        ,'" & Replace(rtxtBug(6).Text, "'", "''") & "'          "
                sSQL = sSQL & vbCrLf & ")"
    
                
                conn.Execute sSQL
                    
    gfMsg ("등록되었습니다.")
    
    '화면 클리어
    Command2_Click
    'commit
    conn.CommitTrans
    
   
    '재조회
    cmdSearch2_Click
    
    
    Screen.MousePointer = vbDefault

    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    Exit Sub
ErrorHandler:
    
    conn.RollbackTrans
    
    
    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description
End Sub

Private Sub RegistHistory(pType As String)

    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    
    

    Set conn = New ADODB.Connection
    conn.ConnectionString = GetConnStringHgDev

    conn.Open
    conn.BeginTrans
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
    Dim i As Integer
    
    If gfMsgQst("저장하시겠습니까?") = vbNo Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
        If pType = "조치완료" Then
                    sSQL = ""
                    sSQL = sSQL & vbCrLf & " UPDATE HGDEV_T003 "
                    sSQL = sSQL & vbCrLf & "    SET 상태 = '조치완료'"
                    sSQL = sSQL & vbCrLf & " WHERE NO =  " & rtxtBug(8).Text & ""
                    conn.Execute sSQL
                    
                        
                    sSQL = ""
                    sSQL = sSQL & vbCrLf & "INSERT INTO HGDEV_T004 (         "
                    sSQL = sSQL & vbCrLf & "          NO               "
                    sSQL = sSQL & vbCrLf & "         ,등록일시         "
                    sSQL = sSQL & vbCrLf & "         ,등록자           "
                    sSQL = sSQL & vbCrLf & "         ,상태             "
                    sSQL = sSQL & vbCrLf & "         ,조치의견         "
                    sSQL = sSQL & vbCrLf & ") VALUES "
                    sSQL = sSQL & vbCrLf & " (        " & rtxtBug(8).Text & ""
                    sSQL = sSQL & vbCrLf & "        ,sysdate               "
                    sSQL = sSQL & vbCrLf & "        ,'" & gUserNm & "'          "
                    sSQL = sSQL & vbCrLf & "        ,'조치완료 등록'           "
                    sSQL = sSQL & vbCrLf & "        ,'" & Replace(rtxtBug(7).Text, "'", "''") & "'          "
                    sSQL = sSQL & vbCrLf & ")"
        
                    
                    conn.Execute sSQL
        ElseIf pType = "반려" Then
        
                    sSQL = ""
                    sSQL = sSQL & vbCrLf & " UPDATE HGDEV_T003 "
                    sSQL = sSQL & vbCrLf & "    SET 상태 = '등록'"
                    sSQL = sSQL & vbCrLf & "       ,확인여부 = 'N' "
                    sSQL = sSQL & vbCrLf & " WHERE NO =  " & rtxtBug(8).Text & ""
                    conn.Execute sSQL
                    
                        
                    sSQL = ""
                    sSQL = sSQL & vbCrLf & "INSERT INTO HGDEV_T004 (         "
                    sSQL = sSQL & vbCrLf & "          NO               "
                    sSQL = sSQL & vbCrLf & "         ,등록일시         "
                    sSQL = sSQL & vbCrLf & "         ,등록자           "
                    sSQL = sSQL & vbCrLf & "         ,상태             "
                    sSQL = sSQL & vbCrLf & "         ,조치의견         "
                    sSQL = sSQL & vbCrLf & ") VALUES "
                    sSQL = sSQL & vbCrLf & " (        " & rtxtBug(8).Text & ""
                    sSQL = sSQL & vbCrLf & "        ,sysdate               "
                    sSQL = sSQL & vbCrLf & "        ,'" & gUserNm & "'          "
                    sSQL = sSQL & vbCrLf & "        ,'반려 등록'           "
                    sSQL = sSQL & vbCrLf & "        ,'" & Replace(rtxtBug(7).Text, "'", "''") & "'          "
                    sSQL = sSQL & vbCrLf & ")"
        
                    
                    conn.Execute sSQL
        
        ElseIf pType = "확인완료" Then
        
        
                    sSQL = ""
                    sSQL = sSQL & vbCrLf & " UPDATE HGDEV_T003 "
                    sSQL = sSQL & vbCrLf & "    SET 상태 = '확인완료'"
                    sSQL = sSQL & vbCrLf & " WHERE NO =  " & rtxtBug(8).Text & ""
                    conn.Execute sSQL
                    
                        
                    sSQL = ""
                    sSQL = sSQL & vbCrLf & "INSERT INTO HGDEV_T004 (         "
                    sSQL = sSQL & vbCrLf & "          NO               "
                    sSQL = sSQL & vbCrLf & "         ,등록일시         "
                    sSQL = sSQL & vbCrLf & "         ,등록자           "
                    sSQL = sSQL & vbCrLf & "         ,상태             "
                    sSQL = sSQL & vbCrLf & "         ,조치의견         "
                    sSQL = sSQL & vbCrLf & ") VALUES "
                    sSQL = sSQL & vbCrLf & " (        " & rtxtBug(8).Text & ""
                    sSQL = sSQL & vbCrLf & "        ,sysdate               "
                    sSQL = sSQL & vbCrLf & "        ,'" & gUserNm & "'          "
                    sSQL = sSQL & vbCrLf & "        ,'확인완료 등록'           "
                    sSQL = sSQL & vbCrLf & "        ,'" & Replace(rtxtBug(7).Text, "'", "''") & "'          "
                    sSQL = sSQL & vbCrLf & ")"
        
                    
                    conn.Execute sSQL
        
        ElseIf pType = "임시저장" Then
        
                    sSQL = sSQL & vbCrLf & "INSERT INTO HGDEV_T004 (         "
                    sSQL = sSQL & vbCrLf & "          NO               "
                    sSQL = sSQL & vbCrLf & "         ,등록일시         "
                    sSQL = sSQL & vbCrLf & "         ,등록자           "
                    sSQL = sSQL & vbCrLf & "         ,상태             "
                    sSQL = sSQL & vbCrLf & "         ,조치의견         "
                    sSQL = sSQL & vbCrLf & ") VALUES "
                    sSQL = sSQL & vbCrLf & " (        " & rtxtBug(8).Text & ""
                    sSQL = sSQL & vbCrLf & "        ,sysdate               "
                    sSQL = sSQL & vbCrLf & "        ,'" & gUserNm & "'          "
                    sSQL = sSQL & vbCrLf & "        ,'임시저장'           "
                    sSQL = sSQL & vbCrLf & "        ,'" & Replace(rtxtBug(7).Text, "'", "''") & "'          "
                    sSQL = sSQL & vbCrLf & ")"
                    
                    conn.Execute sSQL
        
        End If
        
    gfMsg ("등록되었습니다.")
        
    'commit
    conn.CommitTrans
    

    
   
    '재조회
    cmdSearch2_Click
    Call fpBugList_Click(1, 1)
    
    
    
    

    
    
    Screen.MousePointer = vbDefault

    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    Exit Sub
ErrorHandler:
    
    conn.RollbackTrans
    
    
    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description


End Sub
'---------------------------------------------
'SQL문 변환 처리
'---------------------------------------------
Private Sub cmdChange_Click()
On Error GoTo Err_Handler
    
    'SQL문을 공백 단위로 파싱하여 테이블
    Dim sSQLArr() As String
    
    Dim sOrgSQL As String
    
    If Trim(txtNewSQL.Text) = "" Then
        gfMsgInf ("문자열 정리 후 실행하세요")
        Exit Sub
    End If
    
    sOrgSQL = txtNewSQL.Text
    
    
    sOrgSQL = Replace(sOrgSQL, ".", " . ")
    
    sSQLArr = Split(sOrgSQL, " ")
    
    sOrgSQL = Replace(sOrgSQL, " . ", ".")
    
    Dim sTableID() As String
    Dim iTableCnt As Integer
    Dim i As Integer
    Dim j As Integer
    Dim bFind As Boolean
    iTableCnt = -1
    'Table만 찾아낸다
    For i = LBound(sSQLArr) To UBound(sSQLArr)
    
        If Left(Trim(sSQLArr(i)), 3) = "TB_" Or _
           Trim(sSQLArr(i)) = "VI_CU1_CUSTMST" Then
            
            bFind = False
            '중복이 아닌경우만 배열에 저장
            If iTableCnt > -1 Then
                For j = LBound(sTableID) To UBound(sTableID)
                    
                    If sTableID(j) = Trim(sSQLArr(i)) Then
                        bFind = True
                        Exit For
                    End If
                
                
                Next j
            End If
            
            If Not bFind Then
                iTableCnt = iTableCnt + 1
                
                ReDim Preserve sTableID(0 To iTableCnt)
                sTableID(iTableCnt) = Replace(Trim(sSQLArr(i)), vbCrLf, "")
            End If
        
        
        End If
    
    Next i
    
    
    If iTableCnt = -1 Then
        gfMsgWrn ("SQL문의 테이블에 대한 매핑 정보가 업습니다.")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    
    
    '매핑 테이블에서 as-is 테이블의 매핑 정보를 select 한다.
    
    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim rsTemp2 As ADODB.Recordset
    Dim sSQL As String
    
    
    Screen.MousePointer = vbHourglass
    



    Set conn = New ADODB.Connection
    conn.ConnectionString = GetConnStringHgDev

    conn.Open

    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
    Set rsTemp2 = New ADODB.Recordset
    rsTemp2.CursorLocation = adUseClient
    
    
    sSQL = ""
    sSQL = sSQL & vbCrLf & "SELECT NVL(trim(TGT_TAB_ID), ' ') AS TGT_TAB_ID "
    sSQL = sSQL & vbCrLf & "      ,NVL(trim(TGT_TAB_NM), ' ') AS TGT_TAB_NM "
    sSQL = sSQL & vbCrLf & "      ,NVL(trim(TGT_COL_ID), ' ') AS TGT_COL_ID "
    sSQL = sSQL & vbCrLf & "      ,NVL(trim(TGT_COL_NM), ' ') AS TGT_COL_NM "
    sSQL = sSQL & vbCrLf & "      ,NVL(trim(SRC_TAB_ID), ' ') AS SRC_TAB_ID "
    sSQL = sSQL & vbCrLf & "      ,NVL(trim(SRC_TAB_NM), ' ') AS SRC_TAB_NM "
    sSQL = sSQL & vbCrLf & "      ,NVL(trim(SRC_COL_ID), ' ') AS SRC_COL_ID "
    sSQL = sSQL & vbCrLf & "      ,NVL(trim(SRC_COL_NM), ' ') AS SRC_COL_NM "
    sSQL = sSQL & vbCrLf & "  FROM HGDEV_T007"
    sSQL = sSQL & vbCrLf & " WHERE 1=1"
    sSQL = sSQL & vbCrLf & "   AND TGT_COL_ID NOT IN ('FST_REG_DY','FST_REG_EPNO','FST_REG_SRVC_ID','LAST_CHG_DY','LAST_CHG_EPNO','LAST_CHG_SRVC_ID','FST_REG_TS','FST_REG_BROF_CD','LAST_CHG_TS','LAST_CHG_BROF_CD', 'FST_REG_EMP_NO', 'LAST_CHG_EMP_NO' )"
    
    sSQL = sSQL & vbCrLf & "   AND SRC_TAB_ID IN ("
    
    For i = LBound(sTableID) To UBound(sTableID)
        sSQL = sSQL & "'" & Trim(sTableID(i)) & "',"
    Next i
    
    
    '정상ORDER BY
    rsTemp2.Open sSQL & " '' )", conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    sSQL = sSQL & "'OWNER_NAME', 'STORED_FUNCTION', 'COMMON_CODE' ) "
    sSQL = sSQL & " ORDER BY LENGTH(TRIM(SRC_COL_ID)) DESC "

   
   
    
   
    fpMap.MaxRows = 0
    rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    
    
    
    '중복 컬럼 조회 SQL
    
    sSQL = ""
    sSQL = sSQL & vbCrLf & "SELECT NVL(trim(TGT_TAB_ID), ' ') AS TGT_TAB_ID "
    sSQL = sSQL & vbCrLf & "      ,NVL(trim(TGT_TAB_NM), ' ') AS TGT_TAB_NM "
    sSQL = sSQL & vbCrLf & "      ,NVL(trim(TGT_COL_ID), ' ') AS TGT_COL_ID "
    sSQL = sSQL & vbCrLf & "      ,NVL(trim(TGT_COL_NM), ' ') AS TGT_COL_NM "
    sSQL = sSQL & vbCrLf & "      ,NVL(trim(SRC_TAB_ID), ' ') AS SRC_TAB_ID "
    sSQL = sSQL & vbCrLf & "      ,NVL(trim(SRC_TAB_NM), ' ') AS SRC_TAB_NM "
    sSQL = sSQL & vbCrLf & "      ,NVL(trim(SRC_COL_ID), ' ') AS SRC_COL_ID "
    sSQL = sSQL & vbCrLf & "      ,NVL(trim(SRC_COL_NM), ' ') AS SRC_COL_NM "
    sSQL = sSQL & vbCrLf & "      ,NVL(trim(TGT_COL_ID), ' ') || NVL(trim(SRC_COL_ID), ' ') AS COMP_ID "
    
    sSQL = sSQL & vbCrLf & "  FROM HGDEV_T007"
    sSQL = sSQL & vbCrLf & " WHERE 1=1"
    sSQL = sSQL & vbCrLf & "   AND TGT_COL_ID NOT IN ('FST_REG_DY','FST_REG_EPNO','FST_REG_SRVC_ID','LAST_CHG_DY','LAST_CHG_EPNO','LAST_CHG_SRVC_ID','FST_REG_TS','FST_REG_BROF_CD','LAST_CHG_TS','LAST_CHG_BROF_CD',  'FST_REG_EMP_NO', 'LAST_CHG_EMP_NO' )"
    
    sSQL = sSQL & vbCrLf & "   AND SRC_TAB_ID IN ("
    
    For i = LBound(sTableID) To UBound(sTableID)
        sSQL = sSQL & "'" & Trim(sTableID(i)) & "',"
    Next i
    
    sSQL = sSQL & "'OWNER_NAME', 'STORED_FUNCTION', 'COMMON_CODE' ) "
    
    sSQL = sSQL & vbCrLf & "   AND SRC_COL_ID IN ("
    
    sSQL = sSQL & vbCrLf & "SELECT NVL(trim(SRC_COL_ID), ' ') AS SRC_COL_ID "
    sSQL = sSQL & vbCrLf & "  FROM HGDEV_T007"
    sSQL = sSQL & vbCrLf & " WHERE 1=1"
    sSQL = sSQL & vbCrLf & "   AND TGT_COL_ID NOT IN ('FST_REG_DY','FST_REG_EPNO','FST_REG_SRVC_ID','LAST_CHG_DY','LAST_CHG_EPNO','LAST_CHG_SRVC_ID','FST_REG_TS','FST_REG_BROF_CD','LAST_CHG_TS','LAST_CHG_BROF_CD' ,  'FST_REG_EMP_NO', 'LAST_CHG_EMP_NO')"
    
    sSQL = sSQL & vbCrLf & "   AND SRC_TAB_ID IN ("
    
    For i = LBound(sTableID) To UBound(sTableID)
        sSQL = sSQL & "'" & Trim(sTableID(i)) & "',"
    Next i
    
    sSQL = sSQL & vbCrLf & " '' ) "
    sSQL = sSQL & vbCrLf & " AND NVL(trim(SRC_COL_ID), ' ') <> ' '"
    'sSQL = sSQL & vbCrLf & " AND SRC_COL_ID <> '' "
    sSQL = sSQL & vbCrLf & " GROUP BY NVL(trim(SRC_COL_ID), ' ')  "
    sSQL = sSQL & vbCrLf & " HAVING COUNT(*) > 1 )"
    sSQL = sSQL & vbCrLf & " order by SRC_COL_ID , TGT_COL_ID "
        
    
    
    
    

    
    Dim sSrcColID() As String
    Dim sTgtColID() As String
    Dim iMatchCnt As Integer
    iMatchCnt = 0
    
    Dim sTableMap() As String
    Dim sTableMap2() As String
    Dim x As Integer
    Dim bTableMapFind As Boolean
    
    sTableMap = Split(txtTableMap.Text, vbCrLf)
    
    If rsTemp.RecordCount > 0 Then
    
        Call gfRsToFp(fpMap, rsTemp2)
    
        '매핑 조회 건수 만큼 반복하며 as-is 컬럼을 to-be로 변환한다.
        rsTemp.MoveFirst
        For i = 1 To rsTemp.RecordCount
            
            If Trim(rsTemp.Fields("SRC_COL_ID").Value) <> "" Then
            
                If InStr(1, sOrgSQL, " " & rsTemp.Fields("SRC_COL_ID").Value) > 0 Or _
                   InStr(1, sOrgSQL, "," & rsTemp.Fields("SRC_COL_ID").Value) > 0 Or _
                   InStr(1, sOrgSQL, "." & rsTemp.Fields("SRC_COL_ID").Value) > 0 Or _
                   InStr(1, sOrgSQL, "(" & rsTemp.Fields("SRC_COL_ID").Value) > 0 Or _
                   InStr(1, sOrgSQL, "-" & rsTemp.Fields("SRC_COL_ID").Value) > 0 Or _
                   InStr(1, sOrgSQL, "=" & rsTemp.Fields("SRC_COL_ID").Value) > 0 Or _
                   InStr(1, sOrgSQL, "||" & rsTemp.Fields("SRC_COL_ID").Value) > 0 Then
                    
                   iMatchCnt = iMatchCnt + 1
                   ReDim Preserve sSrcColID(1 To iMatchCnt)
                   ReDim Preserve sTgtColID(1 To iMatchCnt)
                   
                   sSrcColID(iMatchCnt) = rsTemp.Fields("SRC_COL_ID").Value
                   sTgtColID(iMatchCnt) = rsTemp.Fields("TGT_COL_ID").Value
                   
                   
                End If
                
            
                '정확도를 높이기 위하여
                sOrgSQL = Replace(sOrgSQL, " " & rsTemp.Fields("SRC_COL_ID").Value, " " & rsTemp.Fields("TGT_COL_ID").Value)
                sOrgSQL = Replace(sOrgSQL, "," & rsTemp.Fields("SRC_COL_ID").Value, "," & rsTemp.Fields("TGT_COL_ID").Value)
                sOrgSQL = Replace(sOrgSQL, "." & rsTemp.Fields("SRC_COL_ID").Value, "." & rsTemp.Fields("TGT_COL_ID").Value)
                sOrgSQL = Replace(sOrgSQL, "(" & rsTemp.Fields("SRC_COL_ID").Value, "(" & rsTemp.Fields("TGT_COL_ID").Value)
                sOrgSQL = Replace(sOrgSQL, "-" & rsTemp.Fields("SRC_COL_ID").Value, "-" & rsTemp.Fields("TGT_COL_ID").Value)
                sOrgSQL = Replace(sOrgSQL, "=" & rsTemp.Fields("SRC_COL_ID").Value, "=" & rsTemp.Fields("TGT_COL_ID").Value)
                sOrgSQL = Replace(sOrgSQL, "||" & rsTemp.Fields("SRC_COL_ID").Value, "||" & rsTemp.Fields("TGT_COL_ID").Value)
                


                
            End If
            
            
            
            bTableMapFind = False
            If Trim(rsTemp.Fields("SRC_TAB_ID").Value) <> "" Then
                For x = LBound(sTableMap) To UBound(sTableMap)
                    
                    If InStr(1, sTableMap(x), "^") Then
                        sTableMap2() = Split(sTableMap(x), "^")
                        
                        If UCase(Trim(sTableMap2(0))) = UCase(Trim(rsTemp.Fields("SRC_TAB_ID").Value)) Then
                        
                            sOrgSQL = Replace(sOrgSQL, rsTemp.Fields("SRC_TAB_ID").Value, sTableMap2(1))
                            bTableMapFind = True
                        End If
                    End If
                        
                Next x
                
                If Not bFind Then
                    sOrgSQL = Replace(sOrgSQL, rsTemp.Fields("SRC_TAB_ID").Value, rsTemp.Fields("TGT_TAB_ID").Value)
                End If
                
            End If
            
            rsTemp.MoveNext
        Next i
        
        txtNewSQL.Text = sOrgSQL
        
        Dim iPos As Integer
        
        rsTemp.MoveFirst
        '변경된 글자 색칠하기
        
        If iMatchCnt > 0 Then
            For i = 1 To UBound(sSrcColID)
    '
    '            If Trim(rsTemp.Fields("TGT_COL_ID").Value) <> "" Then
    '                iPos = 1
    '                Do
    '                    iPos = InStr(iPos, txtNewSQL.Text, rsTemp.Fields("TGT_COL_ID").Value)
    '
    '                    If iPos = 0 Then Exit Do
    '
    '                    txtNewSQL.SelStart = iPos - 1
    '                    txtNewSQL.SelLength = Len(rsTemp.Fields("TGT_COL_ID"))
    '                    txtNewSQL.SelColor = S_RED
    '                    txtNewSQL.SelBold = True
    '                    iPos = iPos + Len(rsTemp.Fields("TGT_COL_ID"))
    '
    '                Loop
    '            End If
    '
                
                
                
                If Trim(sTgtColID(i)) <> "" Then
                    iPos = 1
                    Do
                        iPos = InStr(iPos, txtNewSQL.Text, sTgtColID(i))
                        
                        If iPos = 0 Then Exit Do
                        
                        txtNewSQL.SelStart = iPos - 1
                        txtNewSQL.SelLength = Len(sTgtColID(i))
                        txtNewSQL.SelColor = S_RED
                        txtNewSQL.SelBold = True
                        iPos = iPos + Len(sTgtColID(i))
                    
                    Loop
                End If
                            
                            
                If Trim(sSrcColID(i)) <> "" Then
                    iPos = 1
                    Do
                        iPos = InStr(iPos, txtOldSQL.Text, sSrcColID(i))
    
                        If iPos = 0 Then Exit Do
    
                        txtOldSQL.SelStart = iPos - 1
                        txtOldSQL.SelLength = Len(sSrcColID(i))
                        txtOldSQL.SelColor = S_BLUE
                        txtOldSQL.SelBold = True
                        iPos = iPos + Len(sSrcColID(i))
    
    
                    Loop
                End If
                            
                
                
    '            If Trim(rsTemp.Fields("TGT_TAB_ID").Value) <> "" Then
    '                iPos = 1
    '                Do
    '                    iPos = InStr(iPos, txtNewSQL.Text, rsTemp.Fields("TGT_TAB_ID").Value)
    '
    '                    If iPos = 0 Then Exit Do
    '
    '                    txtNewSQL.SelStart = iPos - 1
    '                    txtNewSQL.SelLength = Len(rsTemp.Fields("TGT_TAB_ID"))
    '                    txtNewSQL.SelColor = S_BLUE
    '                    txtNewSQL.SelBold = True
    '                    iPos = iPos + Len(rsTemp.Fields("TGT_TAB_ID"))
    '
    '                Loop
    '            End If
                
                
    '            rsTemp.MoveNext
            
            Next i
            
            
            
            '변경된 글자 색칠하기 OLD S_BLUE
            
            
            
            rsTemp.MoveFirst
            '변경된 글자 색칠하기
            For i = 1 To rsTemp.RecordCount
                
                
    '            If Trim(rsTemp.Fields("SRC_COL_ID").Value) <> "" Then
    '                iPos = 1
    '                Do
    '                    iPos = InStr(iPos, txtOldSQL.Text, rsTemp.Fields("SRC_COL_ID").Value)
    '
    '                    If iPos = 0 Then Exit Do
    '
    '                    txtOldSQL.SelStart = iPos - 1
    '                    txtOldSQL.SelLength = Len(rsTemp.Fields("SRC_COL_ID"))
    '                    txtOldSQL.SelColor = S_BLUE
    '                    txtOldSQL.SelBold = True
    '                    iPos = iPos + Len(rsTemp.Fields("SRC_COL_ID"))
    '
    '
    '                Loop
    '            End If
                
                If Trim(rsTemp.Fields("TGT_TAB_ID").Value) <> "" Then
                    iPos = 1
                    Do
                        iPos = InStr(iPos, txtNewSQL.Text, rsTemp.Fields("TGT_TAB_ID").Value)
    
                        If iPos = 0 Then Exit Do
    
                        txtNewSQL.SelStart = iPos - 1
                        txtNewSQL.SelLength = Len(rsTemp.Fields("TGT_TAB_ID"))
                        txtNewSQL.SelColor = S_BLUE
                        txtNewSQL.SelBold = True
                        iPos = iPos + Len(rsTemp.Fields("TGT_TAB_ID"))
    
                    Loop
                End If
                            
                
                If Trim(rsTemp.Fields("SRC_TAB_ID").Value) <> "" Then
                    iPos = 1
                    Do
                        iPos = InStr(iPos, txtOldSQL.Text, rsTemp.Fields("SRC_TAB_ID").Value)
                        
                        If iPos = 0 Then Exit Do
                        
                        txtOldSQL.SelStart = iPos - 1
                        txtOldSQL.SelLength = Len(rsTemp.Fields("SRC_TAB_ID"))
                        txtOldSQL.SelColor = S_RED
                        txtOldSQL.SelBold = True
                        iPos = iPos + Len(rsTemp.Fields("SRC_TAB_ID"))
                    
                    Loop
                End If
                
                rsTemp.MoveNext
            
            Next i
            
        End If
        
    End If
    
    
    If iMatchCnt = 0 Then
        gfMsgInf ("SQL문의 컬럼 매핑 정보를 찾을 수 없습니다.")
    End If
    
    txtOldSQL.SelStart = 1
    txtOldSQL.SelLength = Len(txtOldSQL.Text)
    txtOldSQL.SelFontName = "굴림체"
    txtOldSQL.SelFontSize = 9
    txtOldSQL.SelStart = 1
    
    txtNewSQL.SelStart = 1
    txtNewSQL.SelLength = Len(txtOldSQL.Text)
    txtNewSQL.SelFontName = "굴림체"
    txtNewSQL.SelFontSize = 9
    txtNewSQL.SelStart = 1
    
'    txtOldSQL.Font.Name = "굴림체"
'    txtOldSQL.Font.Size = 9
'    txtNewSQL.Font.Name = "굴림체"
'    txtNewSQL.Font.Size = 9
'
            
    Dim rsTemp3 As ADODB.Recordset
    Set rsTemp3 = New ADODB.Recordset
    rsTemp3.CursorLocation = adUseClient
        
    
        
    rsTemp3.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        
    fpDup.MaxRows = 0
    'Call gfRsToFp(fpDup, rsTemp3)
    
    If rsTemp3.RecordCount > 0 Then
        rsTemp3.MoveFirst
        
        Dim sComp As String
        Dim iGridCnt As Integer
        sComp = ""
        
        
        For i = 1 To rsTemp3.RecordCount
            If i = 1 Then
                rsTemp3.MoveFirst
            End If
                
            If sComp <> rsTemp3.Fields("COMP_ID").Value Then
                fpDup.MaxRows = fpDup.MaxRows + 1
                fpDup.Row = fpDup.MaxRows
                
                For j = 1 To 8
                    fpDup.Col = j
                    fpDup.Value = rsTemp3.Fields(j - 1).Value
                Next j
                
                sComp = rsTemp3.Fields("COMP_ID").Value
            End If
            
            rsTemp3.MoveNext
        Next i
    End If
        
        
    Screen.MousePointer = vbDefault

    Set conn = Nothing
    Set rsTemp = Nothing
    Set rsTemp2 = Nothing
    Set rsTemp3 = Nothing

    
    
    
    Exit Sub
Err_Handler:
    
    Screen.MousePointer = vbDefault
    Set conn = Nothing
    Set rsTemp = Nothing
    Set rsTemp2 = Nothing
    Set rsTemp3 = Nothing
    gfMsgWrn (Err.Description)
    
    
    
End Sub

Private Sub cmdClear_Click()
    
    txtOldSQL.Text = Replace(txtOldSQL.Text, "+ " + """", "")
    txtOldSQL.Text = Replace(txtOldSQL.Text, "+ ""\n", "")
    
    txtOldSQL.Text = Replace(txtOldSQL.Text, "\n", "")
    txtOldSQL.Text = Replace(txtOldSQL.Text, """", "")
    txtOldSQL.Text = Replace(txtOldSQL.Text, "//", "--")
    
    txtOldSQL.Text = UCase(txtOldSQL.Text)
    txtNewSQL.Text = UCase(txtOldSQL.Text)
End Sub

Private Sub cmdCommentChange_Click()
'On Error GoTo Err_Handler
    
    
    Dim sSQLArr() As String
    Dim sOrgSQL As String
    Dim i, j As Integer
    Dim sWordArr() As String
    Dim sComment As String

    
    
    If Trim(txtOldComment.Text) = "" Then
        gfMsgInf ("SQL 문장을 입력하세요.")
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    sOrgSQL = txtOldComment.Text
    sSQLArr = Split(sOrgSQL, vbCrLf)
    
    txtNewComment.Text = ""
    
    For i = LBound(sSQLArr) To UBound(sSQLArr)
        DoEvents
        
        sSQLArr(i) = RTrim(sSQLArr(i))
        sWordArr = Split(sSQLArr(i), " ")
        
        If Trim(sSQLArr(i)) <> "" Then
            sComment = getMetaComment(Replace(sWordArr(UBound(sWordArr)), "'", ""))
        End If
        
        If Trim(sComment) <> "" Then
            txtNewComment.Text = txtNewComment.Text & sSQLArr(i) & "    --" & sComment & vbCrLf
        Else
            txtNewComment.Text = txtNewComment.Text & sSQLArr(i) & vbCrLf
        End If
        
        sComment = ""
    Next i
    
    
    
    '주석으로 정렬 하기
    
    Dim sOrgArr() As String
    Dim sSplit() As String
    Dim iMaxSize As Integer
    
    
    sOrgArr = Split(txtNewComment.Text, vbCrLf)
    
    txtNewComment.Text = ""
    iMaxSize = 0
    For i = LBound(sOrgArr) To UBound(sOrgArr)
                    
                    
        If InStr(1, UCase(sOrgArr(i)), "--") Then
            sSplit = Split(sOrgArr(i), "--")
            If iMaxSize < Len(sSplit(0)) Then
                iMaxSize = Len(sSplit(0))
            End If
        
        Else
            
        End If
    Next i
    
    For i = LBound(sOrgArr) To UBound(sOrgArr)
                    
                    
        If InStr(1, UCase(sOrgArr(i)), "--") Then
            sSplit = Split(sOrgArr(i), "--")
            
            txtNewComment.Text = txtNewComment.Text & sSplit(0) & gfMakeString(" ", iMaxSize - Len(sSplit(0)) + 5) & "--"
            txtNewComment.Text = txtNewComment.Text & sSplit(1) & vbCrLf
        Else
            txtNewComment.Text = txtNewComment.Text & sOrgArr(i) & vbCrLf
        End If
    Next i
'
'    '주석 정렬
'    sOrgArr = Split(txtTgtSQL.Text, vbCrLf)
'
'    txtTgtSQL.Text = ""
'    iMaxSize = 0
'    For i = LBound(sOrgArr) To UBound(sOrgArr)
'
'
'        If InStr(1, UCase(sOrgArr(i)), "--") Then
'            sSplit = Split(sOrgArr(i), "--")
'            If iMaxSize < Len(sSplit(0)) Then
'                iMaxSize = Len(sSplit(0))
'            End If
'        ElseIf InStr(1, UCase(sOrgArr(i)), "/*") Then
'            sSplit = Split(sOrgArr(i), "/*")
'            If iMaxSize < Len(sSplit(0)) Then
'                iMaxSize = Len(sSplit(0))
'            End If
'
'        ElseIf InStr(1, UCase(sOrgArr(i)), "//") Then
'            sSplit = Split(sOrgArr(i), "//")
'            If iMaxSize < Len(sSplit(0)) Then
'                iMaxSize = Len(sSplit(0))
'            End If
'
'
'        End If
'    Next i
'
'    For i = LBound(sOrgArr) To UBound(sOrgArr)
'
'
'        If InStr(1, UCase(sOrgArr(i)), "--") Then
'            sSplit = Split(sOrgArr(i), "--")
'
'            txtTgtSQL.Text = txtTgtSQL.Text & sSplit(0) & gfMakeString(" ", iMaxSize - Len(sSplit(0)) + 5) & "--"
'            txtTgtSQL.Text = txtTgtSQL.Text & sSplit(1) & vbCrLf
'        ElseIf InStr(1, UCase(sOrgArr(i)), "/*") Then
'            sSplit = Split(sOrgArr(i), "/*")
'
'            txtTgtSQL.Text = txtTgtSQL.Text & sSplit(0) & gfMakeString(" ", iMaxSize - Len(sSplit(0)) + 5) & "/*"
'            txtTgtSQL.Text = txtTgtSQL.Text & sSplit(1) & vbCrLf
'        Else
'            txtTgtSQL.Text = txtTgtSQL.Text & sOrgArr(i) & vbCrLf
'        End If
'    Next i
'
'
'
'
'
'
'
'
    
    
    
    
    
    
    
    
    
    Screen.MousePointer = vbDefault
    
    gfMsgInf ("주석 적용 완료")
    
    Exit Sub
Err_Handler:
    
    Screen.MousePointer = vbDefault

    gfMsgWrn (Err.Description)
    
        
End Sub

Private Sub cmdConfirm_Click()
    RegistHistory ("확인완료")
End Sub

Private Sub cmdDel_Click()

    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    

    Set conn = New ADODB.Connection
    conn.ConnectionString = GetConnStringHgDev

    conn.Open
    conn.BeginTrans
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
    Dim i As Integer
    
    If gfMsgQst("삭제 하시겠습니까?") = vbNo Then Exit Sub
   
       
    
    
        
        With fpMethod
            .Row = .ActiveRow
            .Col = 1
            
                sSQL = ""
                sSQL = sSQL & vbCrLf & " DELETE FROM HGDEV.HGDEV_T002 "
                sSQL = sSQL & vbCrLf & "  WHERE NO  = " & .Value & ""
    
                conn.Execute sSQL

    
        End With
   
    
    
    
    
    gfMsg ("삭제 되었습니다.")

    
    'commit
    conn.CommitTrans
    
   
    '재조회
    cmdSearch_Click
    
    
    Screen.MousePointer = vbDefault

    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    Exit Sub
ErrorHandler:
    
    conn.RollbackTrans
    
    
    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description
End Sub

Private Sub cmdDMBSearch_Click()
    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    

    Set conn = New ADODB.Connection
    conn.ConnectionString = GetConnStringHgDMB

    conn.Open

    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
    '엑셀에서 복사한 경우 enter key 제거
    cmdSrchTable.Text = Replace(cmdSrchTable.Text, vbCrLf, "")
    
    sSQL = ""
    sSQL = sSQL & vbCrLf & "   SELECT A.TBL_MPNG_ID || '(' || A.TBL_MPNG_ID_SRNO || ')'           AS TBL_MPNG_ID"
    sSQL = sSQL & vbCrLf & "         ,A.CLMN_MPNG_SQNC_NO       "
    sSQL = sSQL & vbCrLf & "         ,A.TARG_PHYC_TBL_NM        "
    sSQL = sSQL & vbCrLf & "         ,B.TARG_LGC_TBL_NM         "
    sSQL = sSQL & vbCrLf & "         ,A.TARG_PHYC_CLMN_NM       "
    sSQL = sSQL & vbCrLf & "         ,A.TARG_LGC_CLMN_NM        "
    sSQL = sSQL & vbCrLf & "         ,A.TARG_CLMN_DATA_TY_NM    "
    sSQL = sSQL & vbCrLf & "         ,A.TARG_CLMN_DATA_LNGT     "
    sSQL = sSQL & vbCrLf & "         ,A.TARG_CLMN_PK_SQNC_NO    "
    sSQL = sSQL & vbCrLf & "         ,A.TARG_CLMN_NN_YN         "
    sSQL = sSQL & vbCrLf & "         ,A.SOUR_PHYC_TBL_NM        "
    sSQL = sSQL & vbCrLf & "         ,C.SOUR_LGC_TBL_NM         "
    sSQL = sSQL & vbCrLf & "         ,A.SOUR_PHYC_CLMN_NM       "
    sSQL = sSQL & vbCrLf & "         ,A.SOUR_LGC_CLMN_NM        "
    sSQL = sSQL & vbCrLf & "         ,A.SOUR_CLMN_DATA_TY_NM    "
    sSQL = sSQL & vbCrLf & "         ,A.SOUR_CLMN_DATA_LNGT     "
    sSQL = sSQL & vbCrLf & "         ,A.SOUR_CLMN_PK_SQNC_NO    "
    sSQL = sSQL & vbCrLf & "         ,A.SOUR_CLMN_NN_YN         "
    sSQL = sSQL & vbCrLf & "         ,A.CLMN_MPNG_PTRN_CD       "
    sSQL = sSQL & vbCrLf & "         ,A.CLMN_MPNG_CNVR_RULE_CNTN "
    sSQL = sSQL & vbCrLf & "         ,A.CLMN_MPNG_ADI_TXT_CNTN   "
    sSQL = sSQL & vbCrLf & "         ,A.CLMN_MPNG_RTN_MTR_CNTN   "
    sSQL = sSQL & vbCrLf & "         ,A.CLMN_MPNG_STS_CD         "
    sSQL = sSQL & vbCrLf & "         ,A.CLMN_MPNG_CHG_DTTM       "
    sSQL = sSQL & vbCrLf & "     FROM MIGADM.TCLMP01 A   --컬럼 매핑정보"
    sSQL = sSQL & vbCrLf & "         ,MIGADM.TTBMP01 B   --타겟 테이블"
    sSQL = sSQL & vbCrLf & "         ,MIGADM.TTBMP02 C   --소스 테이블"
    sSQL = sSQL & vbCrLf & "    WHERE 1=1"
    sSQL = sSQL & vbCrLf & "      AND A.TBL_MPNG_ID = B.TBL_MPNG_ID(+)"
    sSQL = sSQL & vbCrLf & "      AND A.TBL_MPNG_ID = C.TBL_MPNG_ID(+)"
    sSQL = sSQL & vbCrLf & "      AND A.SOUR_PHYC_TBL_NM = C.SOUR_PHYC_TBL_NM(+)"
    
    sSQL = sSQL & vbCrLf & "      AND A.TBL_MPNG_ID_SRNO = B.TBL_MPNG_ID_SRNO(+)"
    sSQL = sSQL & vbCrLf & "      AND A.TBL_MPNG_ID_SRNO = C.TBL_MPNG_ID_SRNO(+)"
    
    If Trim(cmdSrchTable.Text) <> "" Then
        sSQL = sSQL & vbCrLf & "      AND ( A.TARG_PHYC_TBL_NM LIKE '%" & UCase(Trim(cmdSrchTable.Text)) & "%'"
        sSQL = sSQL & vbCrLf & "           OR TARG_LGC_TBL_NM LIKE '%" & UCase(Trim(cmdSrchTable.Text)) & "%'"
        sSQL = sSQL & vbCrLf & "           OR A.SOUR_PHYC_TBL_NM LIKE '%" & UCase(Trim(cmdSrchTable.Text)) & "%'"
        sSQL = sSQL & vbCrLf & "           OR  SOUR_LGC_TBL_NM LIKE '%" & UCase(Trim(cmdSrchTable.Text)) & "%')"
    End If
    
    If Trim(cmdSrchColumn.Text) <> "" Then
        sSQL = sSQL & vbCrLf & "      AND ( A.TARG_PHYC_CLMN_NM LIKE '" & UCase(Trim(cmdSrchColumn.Text)) & "'"
        sSQL = sSQL & vbCrLf & "           OR TARG_LGC_CLMN_NM LIKE '" & UCase(Trim(cmdSrchColumn.Text)) & "'"
        sSQL = sSQL & vbCrLf & "           OR A.SOUR_PHYC_CLMN_NM LIKE '" & UCase(Trim(cmdSrchColumn.Text)) & "'"
        sSQL = sSQL & vbCrLf & "           OR  SOUR_LGC_CLMN_NM LIKE '" & UCase(Trim(cmdSrchColumn.Text)) & "')"
    End If
    




    
    sSQL = sSQL & vbCrLf & "   ORDER BY A.TARG_PHYC_TBL_NM, A.TBL_MPNG_ID_SRNO, A.CLMN_MPNG_SQNC_NO   "

    rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

    fpDMB.MaxRows = 0
    If rsTemp.RecordCount < 1 Then
        fpDMB.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpDMB, rsTemp)
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

Private Sub cmdExcel_Click()

    Call gfFpToExcel(fpMap)
End Sub

Private Sub cmdExcelUpload_Click()
    
'    Dim listcount As Integer
'    'Call fpSrc.GetExcelSheetList("D:\SRC.xlsx", "", LIST_COUNT, "", True)
'    Call fpSrc.GetExcelSheetList("D:\SRC.xlsx", Null, listcount, "", True)
'    Call fpSrc.ImportExcelSheet(listcount, "")
'
    
       ' Declare variables
    Dim y As Boolean, z As Boolean
    Dim Var As Variant
    Dim x As Integer, listcount As Integer, handle As Integer
    Dim List(10) As String

    ' Check if file is an Excel file and set result to x
    x = fpSrc.IsExcelFile("c:\src\src.xls")

    ' If file is Excel file, tell user, import sheet
    ' list, and set result to y
    If x = 1 Then
        'MsgBox "File is an Excel file.", , "File Type"
        y = fpSrc.GetExcelSheetList(txtDirPath & "\VOLIST.xls", List, listcount, "C:\ILOGFILE.TXT", handle, True)
        ' If received sheet list, tell user, import file,
        ' and set result to z
        If y = True Then
            'MsgBox "Got sheet list.", , "Status"
            z = fpSrc.ImportExcelSheet(handle, 0)
            ' Tell user result based on T/F value of z
            If z = True Then
                'MsgBox "Import complete.", , "Result"
            Else
                'MsgBox "Import did not succeed.", , "Result"
            End If
        Else
            ' Tell user cannot obtain sheet list
            MsgBox "Cannot return information for Excel file.", , "Result"
        End If
    Else
        ' Tell user file is not Excel file or is locked
        MsgBox "File is not an Excel file or is locked and cannot be imported.", , "Invalid File Type or Locked"
    End If

End Sub

Private Sub cmdInsert_Click()
    txtData(0).Text = ""
End Sub

Private Sub cmdMakeSrc_Click()


On Error GoTo Err_Handler


    Dim sSrc As String
    
    
    Dim sPname, sLname, sPath As String
    
    
    Dim iCnt As Integer
    
    iCnt = 1
    Do
    With fpSrc
        iCnt = iCnt + 1
        
        .Row = iCnt
        .Col = 1
        
        If Trim(.Value) = "" Then Exit Do
        
        
        
        .Col = 1: sPname = Trim(.Value)
        .Col = 2: sLname = Trim(.Value)
        .Col = 3: sPath = Trim(.Value)
        

        sSrc = ""
        sSrc = sSrc & "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf
        sSrc = sSrc & "<ns5:dto xmlns:ns2=""http://www.tmax.co.kr/proobject/resource"" xmlns:ns3=""http://www.tmax.co.kr/proobject/testcase"" xmlns:ns4=""http://www.tmax.co.kr/proobject/message"" xmlns:ns5=""http://www.tmax.co.kr/proobject/dto"" xmlns:ns6=""http://www.tmaxsoft.co.kr/proobject/testsuite"" "
        sSrc = sSrc & "xmlns:ns7=""http://www.tmax.co.kr/proobject/dto/validator"" xmlns:ns8=""http://www.tmax.co.kr/proobject/serviceobject-automatic"" xmlns:ns9=""http://www.tmax.co.kr/proobject/sourcecode"" xmlns:ns10=""http://www.tmax.co.kr/proobject/flow"" xmlns:ns11=""http://www.tmax.co.kr/proobject/dataobjectfactory"" xmlns:ns12=""http://www.tmax.co.kr/proobject/queryobject"" xmlns:ns13=""http://www.tmax.co.kr/proobject/bizobject"" xmlns:ns14=""http://www.tmax.co.kr/proobject/serviceobject"" xmlns:ns15=""http://www.tmax.co.kr/proobject/dataobject"" xmlns:ns16=""http://www.tmax.co.kr/proobject/taskobject"" xmlns:ns17=""http://www.tmax.co.kr/proobject/jobobject"" xmlns:ns18=""http://www.tmax.co.kr/proobject/servicegroup"" xmlns:ns19="
        sSrc = sSrc & """http://www.tmax.co.kr/proobject/aspectj"" xmlns:ns20=""http://www.tmax.co.kr/proobject/mapping"" xmlns:ns21=""http://www.tmax.co.kr/proobject/application/runtime"" xmlns:ns22=""http://www.tmaxsoft.co.kr/proobject/testOperation"" xmlns:ns23=""http://www.tmax.co.kr/proobject/restriction_codes"" xmlns:ns24=""http://www.example.org/externalObjectConfig"" xmlns:ns25=""http://www.tmax.co.kr/proobject/application"" xmlns:ns26=""http://www.tmax.co.kr/proobject/contents"" xmlns:ns27=""http://www.example.org/testData"" xmlns:ns28=""http://www.tmaxsoft.co.kr/proobject/testresult"" xmlns:ns29=""http://www.tmax.co.kr/proobject/siteConfig"" xmlns:ns30=""http://www.tmax.co.kr/proobject/serverConfig"" xmlns:ns31=""http://www.tmax.co.kr/proobject/probuilder_config"" xmlns:ns32=""http://www.tmax.co.kr/proobject/property"""
        sSrc = sSrc & " xmlns:ns33=""http://www.tmax.co.kr/proobject/serviceGroup"" "
        sSrc = sSrc & "stringDefaultValue=""empty"" "
        sSrc = sSrc & "stringModifiedNullable=""true"" "
        sSrc = sSrc & "resourceId=""" & sPname & """ "
        sSrc = sSrc & "logicalName=""" & sLname & """ "
        sSrc = sSrc & "physicalName=""" & sPname & """ "
        sSrc = sSrc & "resourceType=""DTO"" resourcePackage="""" "
        sSrc = sSrc & "creator=""" & gUserId & """ "
        sSrc = sSrc & "owner=""" & gUserId & """ "
        sSrc = sSrc & "revision=""1"" "
        sSrc = sSrc & "description="""" "
        sSrc = sSrc & "resourcePath=""" & sPath & """ "
        sSrc = sSrc & "modifier=""" & gUserId & """ "
        sSrc = sSrc & "created=""2020-03-20 16:37:53"" "
        sSrc = sSrc & "lastUpdated=""2020-03-25 09:56:27"">" & vbCrLf
        'sSrc = sSrc & "    <ns5:dtoField fieldType=""include"" logicalName=""TEMP"" physicalName=""grntContVo"" columnName="""" tableName="""" schemaName="""" key=""NONE"" includeDtoPath=""hf/cor/hg/ig/vo/bvo/comn"" includeDtoName=""HgigTrstTlgmSendData1601Bvo"" length="""" array="""" decimal=""-1"" keyword="""" defaultValue="""" comments="""" metaId="""" referenceId=""000054892d9b00c757d9ab00000109bf"" nullable=""true"" encrypt=""false""/>" & vbCrLf
        sSrc = sSrc & "</ns5:dto>"
        
        
        
            Dim objStreamUTF8 As Stream
            Set objStreamUTF8 = New Stream
            
            
            Dim objStreamUTF8NoBOM As Stream
            Set objStreamUTF8NoBOM = New Stream
            
            
            With objStreamUTF8
                .Charset = "UTF-8"
                .Open
                .WriteText sSrc
                .Position = 0
                .Type = adTypeText
                .Position = 3
            End With
            
            With objStreamUTF8NoBOM
                .Type = adTypeBinary
                .Open
                objStreamUTF8.CopyTo objStreamUTF8NoBOM
                .SaveToFile txtDirPath.Text & "\" & sPname & ".dto", adSaveCreateOverWrite
                
            End With
            
            objStreamUTF8.Close
            objStreamUTF8NoBOM.Close
            
            Set objStreamUTF8 = Nothing
            Set objStreamUTF8NoBOM = Nothing
                            
    End With
    Loop

    gfMsgInf ("생성되었습니다.")
    Exit Sub
Err_Handler:
    gfMsgWrn (Err.Description)
End Sub

Private Sub cmdMakeVo_Click()
    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    

    Set conn = New ADODB.Connection
    conn.ConnectionString = GetConnStringCorApp

    conn.Open

    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
    
    
    sSQL = ""
    sSQL = "SELECT * FROM ( " & txtVOSQL.Text & " )" & " WHERE 1 = 2"
    
 
    rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

'    fpDMB.MaxRows = 0
'    If rsTemp.RecordCount < 1 Then
'        fpDMB.MaxRows = 1
'    Else
'        '조회된 데이타 디스플레이
'        Call gfRsToFp(fpDMB, rsTemp)
'    End If
    
    fpSQLVO.MaxRows = 0
    fpSQLVO.MaxRows = rsTemp.Fields.Count
    
    
    fpDataCollection.MaxRows = 0
    fpDataCollection.MaxRows = rsTemp.Fields.Count
    
    Dim i As Integer
    For i = 0 To rsTemp.Fields.Count - 1
        With fpSQLVO
            .Row = i + 1
            .Col = 14
            .Value = rsTemp.Fields(i).Name
            
            
            
            
            '매타정보 가져오기
            Call getMetaDB(i + 1, rsTemp.Fields(i).Name)
        End With
        
    
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

Private Sub getMetaDB(iRow As Integer, sWord As String)

    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    'Screen.MousePointer = vbHourglass
    

    Set conn = New ADODB.Connection
    conn.ConnectionString = GetConnStringHgMETA

    conn.Open

    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
        sSQL = ""
        'sSQL = sSQL & vbCrLf & "SELECT DIC_LOG_NM, DIC_PHY_NM"
        sSQL = sSQL & vbCrLf & "SELECT LOGICAL_NAME,FIELD_TYPE , to_char(LENGTH) as LENGTH   "
        sSQL = sSQL & vbCrLf & "       , to_char(CASE WHEN FIELD_TYPE = 'String' then -1 else DECIMAL_SIZE END )  DECIMAL_SIZE"
        sSQL = sSQL & vbCrLf & "       , ENCRYPT "
        sSQL = sSQL & vbCrLf & "  FROM VW_MET_TERM           "
        sSQL = sSQL & vbCrLf & " WHERE 1=1"
        sSQL = sSQL & vbCrLf & "   AND PHYSICAL_NAME  = '" & LCase(sWord) & "'"
    rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Dim sResult As String
    
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        With fpSQLVO
            .Row = iRow
            .Col = 1    '논리명
            .Value = rsTemp.Fields("LOGICAL_NAME").Value
    
            .Col = 11   '논리명
            .Value = rsTemp.Fields("LOGICAL_NAME").Value
            
            .Col = 4    '필드타입
            .Value = rsTemp.Fields("FIELD_TYPE").Value
            
            
            .Col = 6    'length
            .Value = rsTemp.Fields("LENGTH").Value
            
            
            .Col = 2    '변수타입
            .Value = makeEngWord(sWord)
            
            
            fpDataCollection.Row = iRow
            fpDataCollection.Col = 1
            fpDataCollection.Value = makeEngWord(sWord)
            
            fpDataCollection.Col = 2
            fpDataCollection.Value = rsTemp.Fields("LOGICAL_NAME").Value
            
            If UCase(rsTemp.Fields("FIELD_TYPE").Value) = "STRING" Then
                fpDataCollection.Col = 3
                fpDataCollection.Value = "text"
            ElseIf UCase(rsTemp.Fields("FIELD_TYPE").Value) = "BIGDECIMAL" Then
                fpDataCollection.Col = 3
                fpDataCollection.Value = "bigDecimal"
            Else
                fpDataCollection.Col = 3
                fpDataCollection.Value = "number"
            End If
            
            
            
            .Col = 8    'DECIMAL TYPE
            .Value = rsTemp.Fields("DECIMAL_SIZE").Value
            
            .Col = 18
            .Value = rsTemp.Fields("ENCRYPT").Value
            
            .Col = 10
            .Value = "FALSE"
            
            .Col = 15
            .Value = "n"
            
        End With
    End If
    
        
    'Screen.MousePointer = vbDefault

    Set conn = Nothing
    Set rsTemp = Nothing
    

    Exit Sub
ErrorHandler:
    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description
End Sub




Private Function getMetaDBOne(sWord As String)

    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    'Screen.MousePointer = vbHourglass
    

    Set conn = New ADODB.Connection
    conn.ConnectionString = GetConnStringHgMETA

    conn.Open

    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
        sSQL = ""
        'sSQL = sSQL & vbCrLf & "SELECT DIC_LOG_NM, DIC_PHY_NM"
        sSQL = sSQL & vbCrLf & "SELECT DIC_LOG_NM"
        sSQL = sSQL & vbCrLf & "  FROM VW_MET_WORD           "
        sSQL = sSQL & vbCrLf & " WHERE 1=1"
        sSQL = sSQL & vbCrLf & "   AND DIC_PHY_NM  = '" & LCase(sWord) & "'"
    rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    Dim sResult As String
    
    
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        getMetaDBOne = rsTemp.Fields("DIC_LOG_NM").Value
    Else
        getMetaDBOne = ""
    End If
    
        
    'Screen.MousePointer = vbDefault

    Set conn = Nothing
    Set rsTemp = Nothing
    

    Exit Function
ErrorHandler:
    Set conn = Nothing
    Set rsTemp = Nothing
    
    getMetaDBOne = ""
    
    Screen.MousePointer = vbDefault
    
    'gfMsg Err.Description
End Function



Private Function getMetaComment(sStr As String)

    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    'Screen.MousePointer = vbHourglass
    

    Set conn = New ADODB.Connection
    conn.ConnectionString = GetConnStringHgMETA

    conn.Open

    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
        
    Dim sNum As String
    
    sNum = ""
    
    If InStr(1, sStr, ".") > 0 Then
        sStr = Mid(sStr, InStr(1, sStr, ".") + 1)
    End If
    
    If Right(sStr, 1) = "1" Or _
       Right(sStr, 1) = "2" Or _
       Right(sStr, 1) = "3" Or _
       Right(sStr, 1) = "4" Or _
       Right(sStr, 1) = "5" Or _
       Right(sStr, 1) = "6" Or _
       Right(sStr, 1) = "7" Or _
       Right(sStr, 1) = "8" Or _
       Right(sStr, 1) = "9" Then
        
        sNum = Right(sStr, 1)
        sStr = Left(sStr, Len(sStr) - 1)
        
    End If
        
    sSQL = ""
    sSQL = sSQL & vbCrLf & "SELECT LOGICAL_NAME,FIELD_TYPE , to_char(LENGTH) as LENGTH   "
    sSQL = sSQL & vbCrLf & "       , to_char(CASE WHEN FIELD_TYPE = 'String' then -1 else DECIMAL_SIZE END )  DECIMAL_SIZE"
    sSQL = sSQL & vbCrLf & "       , ENCRYPT "
    sSQL = sSQL & vbCrLf & "  FROM VW_MET_TERM           "
    sSQL = sSQL & vbCrLf & " WHERE 1=1"
    sSQL = sSQL & vbCrLf & "   AND PHYSICAL_NAME  = '" & LCase(sStr) & "'"
    rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    

    
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        getMetaComment = rsTemp.Fields("LOGICAL_NAME").Value
        
        If sNum <> "" Then
            getMetaComment = getMetaComment & sNum
        End If
        
    Else
        getMetaComment = ""
    End If
    
            

    Set conn = Nothing
    Set rsTemp = Nothing
    

    Exit Function
ErrorHandler:
    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description
End Function



Private Function makeEngWord(sInStr As String)

    Dim sWord() As String
    Dim i As Integer
    
    sWord = Split(sInStr, "_")
    Dim sRet As String
    
    For i = LBound(sWord) To UBound(sWord)
        If i = 0 Then
            sRet = LCase(sWord(i))
        Else
            sRet = sRet & UCase(Left(sWord(i), 1)) & LCase(Mid(sWord(i), 2))
        End If
            
    
    Next i
    
    makeEngWord = sRet

End Function





Private Sub cmdMig_Click()
    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    Dim sColId As String
    Dim sColVal As String
    
    If optMig1.Value = True Then
    
        If Trim(txtGrntNo.Text) = "" Then
            gfMsgWrn "보증번호를 입력하세요!!"
            txtGrntNo.SetFocus
            Exit Sub
        End If
        
        sColId = "GRNT_NO"
        sColVal = txtGrntNo.Text
    End If
    
    If optMig2.Value = True Then
    
        If Trim(txtColId.Text) = "" Then
            gfMsgWrn "컬럼명을 입력하세요!!"
            txtColId.SetFocus
            Exit Sub
        End If
        
        If Trim(txtColVal.Text) = "" Then
            gfMsgWrn "이행값을 입력하세요!!"
            txtColVal.SetFocus
            Exit Sub
        End If
        
        sColId = UCase(txtColId.Text)
        sColVal = txtColVal.Text
        
    End If
    
    
    
    
    
    Screen.MousePointer = vbHourglass
    

    Set conn = New ADODB.Connection
    conn.ConnectionString = GetConnStringCorApp

    conn.Open

    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
    
    
    sSQL = ""
    sSQL = ""
    sSQL = sSQL & "SELECT  A.TABLE_NAME       "
    sSQL = sSQL & "       ,(SELECT X.COMMENTS                     "
    sSQL = sSQL & "          FROM ALL_TAB_COMMENTS X"
    sSQL = sSQL & "         WHERE A.OWNER = X.OWNER"
    sSQL = sSQL & "          AND  A.TABLE_NAME = X.TABLE_NAME )   TABLE_COMM       "
    sSQL = sSQL & "       ,''      "
'    sSQL = sSQL & "       ,B.COLUMN_NAME       "
'    sSQL = sSQL & "       ,B.COLUMN_POSITION  "
'    sSQL = sSQL & "       ,B.DESCEND                           "
'    sSQL = sSQL & "       ,A.INDEX_TYPE                        "
'    sSQL = sSQL & "       ,A.TABLESPACE_NAME                   "
    sSQL = sSQL & "  FROM  ALL_INDEXES  A"
    sSQL = sSQL & "       ,ALL_IND_COLUMNS B"
    sSQL = sSQL & "  WHERE 1=1"
    sSQL = sSQL & "   AND A.TABLE_OWNER = 'USHG'"
    sSQL = sSQL & "   AND A.TABLE_OWNER = B.TABLE_OWNER          "
    sSQL = sSQL & "   AND A.INDEX_NAME = B.INDEX_NAME          "
    sSQL = sSQL & "   AND A.TABLE_NAME = B.TABLE_NAME          "
    sSQL = sSQL & "   AND B.COLUMN_NAME = '" & sColId & "'"
    sSQL = sSQL & "   AND B.COLUMN_POSITION   = 1"
    sSQL = sSQL & "ORDER BY A.INDEX_NAME, B.COLUMN_POSITION    "
    
 
    rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText


    fpMig.MaxRows = 0
    If rsTemp.RecordCount < 1 Then
        fpMig.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpMig, rsTemp)
    End If
    
    
    Dim i As Integer
    Dim iCnt As Integer

    For i = 1 To fpMig.MaxRows
DoEvents
        conn.BeginTrans
            
        fpMig.Row = i
        fpMig.Col = 1
        sSQL = ""
        sSQL = sSQL & " DELETE FROM " & fpMig.Value
        sSQL = sSQL & " WHERE " & sColId & " = '" & sColVal & "'"
    
        On Error GoTo DbErr1
        
        conn.Execute sSQL
            
        GoTo Step2
            
DbErr1:
            fpMig.Col = 3
            fpMig.Value = Err.Description
        
Step2:
        sSQL = ""
        sSQL = sSQL & " INSERT INTO  " & fpMig.Value
        sSQL = sSQL & " SELECT * FROM " & fpMig.Value & "@DL_ENTDBT "
        sSQL = sSQL & " WHERE  " & sColId & "  = '" & sColVal & "'"
        
        On Error GoTo DbErr2
        
         conn.Execute sSQL, iCnt
        
        GoTo Step3
DbErr2:
        fpMig.Col = 3
        fpMig.Value = Err.Description
        

Step3:
       On Error GoTo DbErr3
        
        conn.CommitTrans
        
DbErr3:
        fpMig.Col = 3
        fpMig.Value = Err.Description
        
                
        fpMig.Col = 3
        If Trim(fpMig.Value) = "" Then
            fpMig.Value = iCnt & "건 이관"
        End If
    Next i
    
    
        
    Screen.MousePointer = vbDefault

    gfMsgInf ("이관이 완료되었습니다.")

    Set conn = Nothing
    Set rsTemp = Nothing
    

    Exit Sub
ErrorHandler:
    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description
End Sub

Private Sub cmdMinMax_Click(Index As Integer)
    If Index = 0 Then
        fpDup.Top = 1440
        fpDup.Height = 11175
    Else
        fpDup.Top = 11220
        fpDup.Height = 1395
    End If
End Sub

Private Sub cmdNameChange_Click()

    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    
    

    Set conn = New ADODB.Connection
    conn.ConnectionString = GetConnStringHgDev

    conn.Open
    conn.BeginTrans
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
    Dim i As Integer
    
    If gfMsgQst("조치자를 변경하시겠습니까?") = vbNo Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    

            sSQL = ""
            sSQL = sSQL & vbCrLf & " UPDATE HGDEV_T003 "
            sSQL = sSQL & vbCrLf & "    SET 개발자 = '" & txtName.Text & "'"
            sSQL = sSQL & vbCrLf & " WHERE NO =  " & rtxtBug(8).Text & ""
            conn.Execute sSQL
            
                
            sSQL = ""
            sSQL = sSQL & vbCrLf & "INSERT INTO HGDEV_T004 (         "
            sSQL = sSQL & vbCrLf & "          NO               "
            sSQL = sSQL & vbCrLf & "         ,등록일시         "
            sSQL = sSQL & vbCrLf & "         ,등록자           "
            sSQL = sSQL & vbCrLf & "         ,상태             "
            sSQL = sSQL & vbCrLf & "         ,조치의견         "
            sSQL = sSQL & vbCrLf & ") VALUES "
            sSQL = sSQL & vbCrLf & " (        " & rtxtBug(8).Text & ""
            sSQL = sSQL & vbCrLf & "        ,sysdate               "
            sSQL = sSQL & vbCrLf & "        ,'" & gUserNm & "'          "
            sSQL = sSQL & vbCrLf & "        ,'조치자 변경'           "
            sSQL = sSQL & vbCrLf & "        ,'조치자 변경 " & rtxtBug(2).Text & "   ->  " & txtName.Text & "'"
            sSQL = sSQL & vbCrLf & ")"

            
            conn.Execute sSQL
        
    gfMsg ("변경 되었습니다.")
        
    'commit
    conn.CommitTrans
    

    
   
    '재조회
    cmdSearch2_Click
    Call fpBugList_Click(1, 1)
    

    
    
    Screen.MousePointer = vbDefault

    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    Exit Sub
ErrorHandler:
    
    conn.RollbackTrans
    
    
    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description
End Sub

Private Sub cmdNew_Click()
    Dim i As Integer
    For i = txtData.LBound To txtData.UBound
        txtData(i).Text = ""
    Next i
End Sub

Private Sub cmdNewInput_Click()
    rtxtBug(7).Text = ""
End Sub

Private Sub cmdProcComp_Click()
    RegistHistory ("조치완료")
End Sub

Private Sub cmdReject_Click()
    RegistHistory ("반려")
End Sub

Private Sub cmdSave_Click()

    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    

    Set conn = New ADODB.Connection
    conn.ConnectionString = GetConnStringHgDev

    conn.Open
    conn.BeginTrans
    
    'Set rsTemp = New ADODB.Recordset
    'rsTemp.CursorLocation = adUseClient
    
    Dim i As Integer
    
    If gfMsgQst("저장하시겠습니까?") = vbNo Then Exit Sub
   
       
    
    For i = 1 To fpMethod.MaxRows
        
        With fpMethod
            .Row = i
            .Col = 1
            
            If .BackColor = C_ADD_COLOR Then
                
                '삭제하고 insert
                If Trim(.Value) <> "" Then
                    
                    sSQL = ""
                    sSQL = sSQL & vbCrLf & " DELETE FROM HGDEV.HGDEV_T002 "
                    sSQL = sSQL & vbCrLf & "  WHERE NO  = " & .Value & ""
        
                    conn.Execute sSQL
                                
                '채번하기
                Else
                    sSQL = ""
                    sSQL = sSQL & vbCrLf & " SELECT MAX(NO) + 1 FROM  HGDEV.HGDEV_T002  "
                    
                    Set rsTemp = New ADODB.Recordset
                    rsTemp.CursorLocation = adUseClient
                    
                    rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
                    
                    
                    
                    .Row = i
                    .Col = 1
                    Dim sNo As String
                    sNo = rsTemp.Fields(0).Value
                    fpMethod.Value = sNo
                    
                    Set rsTemp = Nothing
                End If
                
                    '등록하기
                              sSQL = ""
                              sSQL = sSQL & vbCrLf & " INSERT INTO HGDEV_T002 (         "
                              sSQL = sSQL & vbCrLf & "           NO                        "
                              sSQL = sSQL & vbCrLf & "          ,ASIS패키지                   "
                              sSQL = sSQL & vbCrLf & "          ,ASIS하위그룹                  "
                              sSQL = sSQL & vbCrLf & "          ,ASIS_BEAN                 "
                              sSQL = sSQL & vbCrLf & "          ,ASIS클래스설명                 "
                              sSQL = sSQL & vbCrLf & "          ,ASIS_METHOD               "
                              sSQL = sSQL & vbCrLf & "          ,ASIS_METHOD설명             "
                              sSQL = sSQL & vbCrLf & "          ,프로그램구분                    "
                              sSQL = sSQL & vbCrLf & "          ,TOBE_BIZ                  "
                              sSQL = sSQL & vbCrLf & "          ,TOBE_BIZ명                 "
                              sSQL = sSQL & vbCrLf & "          ,TOBE_업무구분                 "
                              sSQL = sSQL & vbCrLf & "          ,TOBE_METHOD               "
                              sSQL = sSQL & vbCrLf & "          ,TOBE_METHOD명              "
                              sSQL = sSQL & vbCrLf & "          ,비고                        "
                              sSQL = sSQL & vbCrLf & "          ,이관구분                        "
                              sSQL = sSQL & vbCrLf & " ) VALUES "
                    .Col = 1: sSQL = sSQL & vbCrLf & " (       " & .Value & ""
                    .Col = 2: sSQL = sSQL & vbCrLf & "        ,'" & .Value & "'          "
                    .Col = 3: sSQL = sSQL & vbCrLf & "        ,'" & .Value & "'          "
                    .Col = 4: sSQL = sSQL & vbCrLf & "        ,'" & .Value & "'          "
                    .Col = 5: sSQL = sSQL & vbCrLf & "        ,'" & .Value & "'          "
                    .Col = 6: sSQL = sSQL & vbCrLf & "        ,'" & .Value & "'          "
                    .Col = 7: sSQL = sSQL & vbCrLf & "        ,'" & .Value & "'          "
                    .Col = 8: sSQL = sSQL & vbCrLf & "        ,'" & .Value & "'          "
                    .Col = 9: sSQL = sSQL & vbCrLf & "        ,'" & .Value & "'          "
                    .Col = 10: sSQL = sSQL & vbCrLf & "       ,'" & .Value & "'          "
                    .Col = 11: sSQL = sSQL & vbCrLf & "       ,'" & .Value & "'          "
                    .Col = 12: sSQL = sSQL & vbCrLf & "       ,'" & .Value & "'          "
                    .Col = 13: sSQL = sSQL & vbCrLf & "       ,'" & .Value & "'          "
                    .Col = 14: sSQL = sSQL & vbCrLf & "       ,'" & .Value & "'          "
                    .Col = 15: sSQL = sSQL & vbCrLf & "       ,'" & .Value & "'          "
                    .Col = 16: sSQL = sSQL & vbCrLf & " )"
                                    
            
                    conn.Execute sSQL
        
            End If
    
        End With
   
    Next i
    
    
    
    
    gfMsg ("등록되었습니다.")

    'commit
    conn.CommitTrans
    
   
    '재조회
    Call cmdSearch_Click
    
    
    
    Screen.MousePointer = vbDefault

    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    Exit Sub
ErrorHandler:
    
    conn.RollbackTrans
    
    
    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description
End Sub

'--------------------------
'조회
'--------------------------
Private Sub cmdSearch_Click()
    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    

    Set conn = New ADODB.Connection
    conn.ConnectionString = GetConnStringHgDev

    conn.Open

    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
    sSQL = ""
    sSQL = sSQL & vbCrLf & "SELECT  NO                         "
    sSQL = sSQL & vbCrLf & "       ,ASIS패키지                    "
    sSQL = sSQL & vbCrLf & "       ,ASIS하위그룹                   "
    sSQL = sSQL & vbCrLf & "       ,ASIS_BEAN                  "
    sSQL = sSQL & vbCrLf & "       ,ASIS클래스설명                  "
    sSQL = sSQL & vbCrLf & "       ,ASIS_METHOD                "
    sSQL = sSQL & vbCrLf & "       ,ASIS_METHOD설명              "
    sSQL = sSQL & vbCrLf & "       ,프로그램구분                     "
    sSQL = sSQL & vbCrLf & "       ,TOBE_BIZ                   "
    sSQL = sSQL & vbCrLf & "       ,TOBE_BIZ명                  "
    sSQL = sSQL & vbCrLf & "       ,TOBE_업무구분                  "
    sSQL = sSQL & vbCrLf & "       ,TOBE_METHOD                "
    sSQL = sSQL & vbCrLf & "       ,TOBE_METHOD명               "
    sSQL = sSQL & vbCrLf & "       ,비고                         "
    sSQL = sSQL & vbCrLf & "       ,이관구분                       "
    sSQL = sSQL & vbCrLf & "  FROM HGDEV_T002        /* */"
    sSQL = sSQL & vbCrLf & " WHERE 1=1"
    
    '엑셀에서 복사한 경우 enter key 제거
    txtCondition.Text = Replace(txtCondition.Text, vbCrLf, "")

    If Trim(txtCondition.Text) <> "" Then
    
    sSQL = sSQL & vbCrLf & "   AND ( UPPER(ASIS패키지)                LIKE '%" & UCase(Trim(txtCondition.Text)) & "%'"
    sSQL = sSQL & vbCrLf & "   OR UPPER(ASIS하위그룹)              LIKE '%" & UCase(Trim(txtCondition.Text)) & "%'"
    sSQL = sSQL & vbCrLf & "   OR UPPER(ASIS_BEAN)                 LIKE '%" & UCase(Trim(txtCondition.Text)) & "%'"
    sSQL = sSQL & vbCrLf & "   OR UPPER(ASIS클래스설명)            LIKE '%" & UCase(Trim(txtCondition.Text)) & "%'"
    sSQL = sSQL & vbCrLf & "   OR UPPER(ASIS_METHOD)               LIKE '%" & UCase(Trim(txtCondition.Text)) & "%'"
    sSQL = sSQL & vbCrLf & "   OR UPPER(ASIS_METHOD설명)           LIKE '%" & UCase(Trim(txtCondition.Text)) & "%'"
    sSQL = sSQL & vbCrLf & "   OR UPPER(TOBE_BIZ)                  LIKE '%" & UCase(Trim(txtCondition.Text)) & "%'"
    sSQL = sSQL & vbCrLf & "   OR UPPER(TOBE_BIZ명)                LIKE '%" & UCase(Trim(txtCondition.Text)) & "%'"
    sSQL = sSQL & vbCrLf & "   OR UPPER(TOBE_업무구분)             LIKE '%" & UCase(Trim(txtCondition.Text)) & "%'"
    sSQL = sSQL & vbCrLf & "   OR UPPER(TOBE_METHOD)               LIKE '%" & UCase(Trim(txtCondition.Text)) & "%'"
    sSQL = sSQL & vbCrLf & "   OR UPPER(TOBE_METHOD명)             LIKE '%" & UCase(Trim(txtCondition.Text)) & "%'"
    sSQL = sSQL & vbCrLf & "   )"
    End If

    sSQL = sSQL & vbCrLf & " ORDER BY ASIS패키지, ASIS하위그룹, ASIS_BEAN, ASIS_METHOD"
    
   

    rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

    fpMethod.MaxRows = 0
    If rsTemp.RecordCount < 1 Then
        fpMethod.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpMethod, rsTemp)
    End If
    
    
    
    
    makeCellBorderline
    
        
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


Private Sub cmdSearch2_Click()
    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    

    Set conn = New ADODB.Connection
    conn.ConnectionString = GetConnStringHgDev

    conn.Open

    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
        sSQL = ""
        sSQL = sSQL & vbCrLf & "SELECT  NO                  "
        sSQL = sSQL & vbCrLf & "       ,등록자              "
        sSQL = sSQL & vbCrLf & "       ,개발자              "
        sSQL = sSQL & vbCrLf & "       ,상태                "
        sSQL = sSQL & vbCrLf & "       ,DECODE(NVL(UI여부,'0') , '1', 'Y','')               "
        sSQL = sSQL & vbCrLf & "       ,프로그램ID                "
        sSQL = sSQL & vbCrLf & "       ,프로그램명          "
        sSQL = sSQL & vbCrLf & "       ,등록일시          "
        sSQL = sSQL & vbCrLf & "       ,결함내용            "
        sSQL = sSQL & vbCrLf & "  FROM HGDEV_T003           "
        sSQL = sSQL & vbCrLf & " WHERE 1=1"

        If Trim(txtUserNm.Text) <> "" Then
            sSQL = sSQL & vbCrLf & "   AND (     등록자                 = '" & Trim(txtUserNm.Text) & "'          "
            sSQL = sSQL & vbCrLf & "         OR  개발자                 = '" & Trim(txtUserNm.Text) & "'          "
            sSQL = sSQL & vbCrLf & "       ) "
        End If
        
        If Trim(txtPGMid.Text) <> "" Then
            sSQL = sSQL & vbCrLf & "   AND 프로그램ID                LIKE  '%" & txtPGMid.Text & "%'          -- "
        
        End If
        
        If Trim(txtPGMNM.Text) <> "" Then
            sSQL = sSQL & vbCrLf & "   AND 프로그램명                LIKE  '%" & txtPGMNM.Text & "%'          -- "
        
        End If
        
        
        If optStatus(1).Value = True Then
            sSQL = sSQL & vbCrLf & "   AND 상태                  <> '확인완료'          -- "
        End If
        
        If optStatus(2).Value = True Then
            sSQL = sSQL & vbCrLf & "   AND 상태                  = '확인완료'          -- "
        End If
        If optStatus(3).Value = True Then
            sSQL = sSQL & vbCrLf & "   AND 상태                  = '조치완료'          -- "
        End If
        
        
        If chkSelUI.Value = "1" Then
            sSQL = sSQL & vbCrLf & "   AND UI여부                  = '1'          -- "
        End If
        
        
        sSQL = sSQL & vbCrLf & " ORDER BY 1 DESC "
'
'        sSQL = sSQL & vbCrLf & "   AND NO                  = XXX            -- "
'        sSQL = sSQL & vbCrLf & "   AND 프로그램ID              = 'XXX'          -- "
'        sSQL = sSQL & vbCrLf & "   AND 프로그램명               = 'XXX'          -- "
'        sSQL = sSQL & vbCrLf & "   AND 상태                  = 'XXX'          -- "
'        sSQL = sSQL & vbCrLf & "   AND 결함내용                = 'XXX'          -- "
'        sSQL = sSQL & vbCrLf & "   AND 파일명                 = 'XXX'          -- "
'        sSQL = sSQL & vbCrLf & "   AND 등록일시                = 'XXX'          -- "
'
'    sSQL = sSQL & vbCrLf & "   AND ( UPPER(ASIS패키지)                LIKE '%" & UCase(Trim(txtCondition.Text)) & "%'"
'    sSQL = sSQL & vbCrLf & "   OR UPPER(ASIS하위그룹)              LIKE '%" & UCase(Trim(txtCondition.Text)) & "%'"
'    sSQL = sSQL & vbCrLf & "   OR UPPER(ASIS_BEAN)                 LIKE '%" & UCase(Trim(txtCondition.Text)) & "%'"
'    sSQL = sSQL & vbCrLf & "   OR UPPER(ASIS클래스설명)            LIKE '%" & UCase(Trim(txtCondition.Text)) & "%'"
'    sSQL = sSQL & vbCrLf & "   OR UPPER(ASIS_METHOD)               LIKE '%" & UCase(Trim(txtCondition.Text)) & "%'"
'    sSQL = sSQL & vbCrLf & "   OR UPPER(ASIS_METHOD설명)           LIKE '%" & UCase(Trim(txtCondition.Text)) & "%'"
'    sSQL = sSQL & vbCrLf & "   OR UPPER(TOBE_BIZ)                  LIKE '%" & UCase(Trim(txtCondition.Text)) & "%'"
'    sSQL = sSQL & vbCrLf & "   OR UPPER(TOBE_BIZ명)                LIKE '%" & UCase(Trim(txtCondition.Text)) & "%'"
'    sSQL = sSQL & vbCrLf & "   OR UPPER(TOBE_업무구분)             LIKE '%" & UCase(Trim(txtCondition.Text)) & "%'"
'    sSQL = sSQL & vbCrLf & "   OR UPPER(TOBE_METHOD)               LIKE '%" & UCase(Trim(txtCondition.Text)) & "%'"
'    sSQL = sSQL & vbCrLf & "   OR UPPER(TOBE_METHOD명)             LIKE '%" & UCase(Trim(txtCondition.Text)) & "%'"
'    sSQL = sSQL & vbCrLf & "   )"
'    End If
'
'    sSQL = sSQL & vbCrLf & " ORDER BY ASIS패키지, ASIS하위그룹, ASIS_BEAN, ASIS_METHOD"
'
'

    rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

    fpBugList.MaxRows = 0
    If rsTemp.RecordCount < 1 Then
        fpBugList.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpBugList, rsTemp)
    End If
    
    
    Dim i As Integer
    For i = 1 To fpBugList.MaxRows
        fpBugList.RowHeight(i) = 15
    Next i
    
    
    fpBugList.Row = -1
    fpBugList.Col = -1
    fpBugList.Lock = True
    
    
    'makeCellBorderline
    
        
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

Private Sub cmdSingleSave_Click()

    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    

    Set conn = New ADODB.Connection
    conn.ConnectionString = GetConnStringHgDev

    conn.Open
    conn.BeginTrans
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
    Dim i As Integer
    
    If gfMsgQst("저장하시겠습니까?") = vbNo Then Exit Sub
   
       
    
            

                
                '삭제하고 insert
                If Trim(txtData(0).Text) <> "" Then
                    
                    sSQL = ""
                    sSQL = sSQL & vbCrLf & " DELETE FROM HGDEV.HGDEV_T002 "
                    sSQL = sSQL & vbCrLf & "  WHERE NO  = " & txtData(0).Text & ""
        
                    conn.Execute sSQL
                                
                '채번하기
                Else
                    sSQL = ""
                    sSQL = sSQL & vbCrLf & " SELECT MAX(NO) + 1 FROM  HGDEV.HGDEV_T002  "
                    
                    rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
                    
                    
                    Dim sNo As String
                    sNo = rsTemp.Fields(0).Value
                    txtData(0).Text = sNo
                    
                    
                End If
                
                    '등록하기
                              sSQL = ""
                              sSQL = sSQL & vbCrLf & " INSERT INTO HGDEV_T002 (         "
                              sSQL = sSQL & vbCrLf & "           NO                        "
                              sSQL = sSQL & vbCrLf & "          ,ASIS패키지                   "
                              sSQL = sSQL & vbCrLf & "          ,ASIS하위그룹                  "
                              sSQL = sSQL & vbCrLf & "          ,ASIS_BEAN                 "
                              sSQL = sSQL & vbCrLf & "          ,ASIS클래스설명                 "
                              sSQL = sSQL & vbCrLf & "          ,ASIS_METHOD               "
                              sSQL = sSQL & vbCrLf & "          ,ASIS_METHOD설명             "
                              sSQL = sSQL & vbCrLf & "          ,프로그램구분                    "
                              sSQL = sSQL & vbCrLf & "          ,TOBE_BIZ                  "
                              sSQL = sSQL & vbCrLf & "          ,TOBE_BIZ명                 "
                              sSQL = sSQL & vbCrLf & "          ,TOBE_업무구분                 "
                              sSQL = sSQL & vbCrLf & "          ,TOBE_METHOD               "
                              sSQL = sSQL & vbCrLf & "          ,TOBE_METHOD명              "
                              sSQL = sSQL & vbCrLf & "          ,비고                        "
                              sSQL = sSQL & vbCrLf & "          ,이관구분                        "
                              sSQL = sSQL & vbCrLf & " ) VALUES "
                              sSQL = sSQL & vbCrLf & " (        " & txtData(0).Text & ""
                              sSQL = sSQL & vbCrLf & "        ,'" & txtData(1).Text & "'          "
                              sSQL = sSQL & vbCrLf & "        ,'" & txtData(2).Text & "'          "
                              sSQL = sSQL & vbCrLf & "        ,'" & txtData(3).Text & "'          "
                              sSQL = sSQL & vbCrLf & "        ,'" & txtData(4).Text & "'          "
                              sSQL = sSQL & vbCrLf & "        ,'" & txtData(5).Text & "'          "
                              sSQL = sSQL & vbCrLf & "        ,'" & txtData(6).Text & "'          "
                              sSQL = sSQL & vbCrLf & "        ,'" & txtData(7).Text & "'          "
                              sSQL = sSQL & vbCrLf & "        ,'" & txtData(8).Text & "'          "
                              sSQL = sSQL & vbCrLf & "       ,'" & txtData(9).Text & "'          "
                              sSQL = sSQL & vbCrLf & "       ,'" & txtData(10).Text & "'          "
                              sSQL = sSQL & vbCrLf & "       ,'" & txtData(11).Text & "'          "
                              sSQL = sSQL & vbCrLf & "       ,'" & txtData(12).Text & "'          "
                              sSQL = sSQL & vbCrLf & "       ,'" & txtData(13).Text & "'          "
                              sSQL = sSQL & vbCrLf & "       ,'" & txtData(14).Text & "'          "
                              sSQL = sSQL & vbCrLf & " )"
                                    
            
                    conn.Execute sSQL
        
            
    

    
    
    
    
    gfMsg ("등록되었습니다.")

    'commit
    conn.CommitTrans
    
   
    '재조회
    cmdSearch_Click
    
    
    Screen.MousePointer = vbDefault

    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    Exit Sub
ErrorHandler:
    
    conn.RollbackTrans
    
    
    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description
End Sub

Private Sub cmdSQlClear_Click()
    txtVOSQL.Text = ""
    fpSQLVO.MaxRows = 0
    fpDataCollection.MaxRows = 0
End Sub

Private Sub cmdSrchColumn_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdDMBSearch_Click
    End If
End Sub

Private Sub cmdSrchTable_GotFocus()
    cmdSrchTable.SelStart = 0
    cmdSrchTable.SelLength = Len(cmdSrchTable.Text)
End Sub

Private Sub cmdSrchTable_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdDMBSearch_Click
    End If
End Sub

Private Sub cmdTempSave_Click()
    RegistHistory ("임시저장")
End Sub

Private Sub cmdTxtClear_Click()
    txtNewSQL.Text = ""
    txtOldSQL.Text = ""
End Sub

Private Sub Command1_Click()

    Call gfFpToExcel(fpDMB)
End Sub

Private Sub Command2_Click()
    Dim i As Integer
    For i = rtxtBug.LBound To rtxtBug.UBound
        Select Case i
            Case 0, 1, 2, 3, 4, 5, 6, 8, 7
                rtxtBug(i).Text = ""
        End Select
    Next i
    
    rtxtBug(3).Text = gUserNm
    cmdBugReg.Enabled = True
    fpProcHist.MaxRows = 0
    fpProcHist.MaxRows = 1
    chkUI.Value = "0"
    
    
End Sub

Private Sub Command3_Click()
    Dim i As Integer
    For i = 1 To fpDMB.MaxRows
        fpDMB.Col = -1
        fpDMB.Row = i
        fpDMB.BackColor = S_WHITE
    Next i
End Sub

Private Sub Command4_Click()
    Call gfFpToExcel(fpDup)
End Sub

Private Sub Command5_Click()
    fpSQLVO.SetSelection 1, 1, fpSQLVO.MaxCols, fpSQLVO.MaxRows
    fpSQLVO.ClipboardCopy
    gfMsgInf "Copy clipboard"
End Sub

Private Sub Command6_Click()
    fpDataCollection.SetSelection 1, 1, fpDataCollection.MaxCols, fpDataCollection.MaxRows
    fpDataCollection.ClipboardCopy
    gfMsgInf "Copy clipboard"
End Sub

Private Sub Command7_Click()
    fpJson.MaxRows = 0
    fpJson.MaxRows = 500
End Sub

Private Sub Command8_Click()
    Dim i As Integer
    Dim iRow As Integer
    Dim sTemp As String
    Dim sWord As String
    Dim sDaneo As String
    Dim iCnt As Integer
    
    Dim sUpper As String
    sUpper = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"
    
    Dim sUpperArr() As String
    
    iCnt = 1
    
    ReDim sUpperArr(1 To Len(sUpper))
    
    For i = 1 To UBound(sUpperArr)
        sUpperArr(i) = Mid(sUpper, i, 1)
    Next i
    
    
    For iRow = 1 To fpJson.MaxRows
        fpJson.Row = iRow

        fpJson.Col = 1
        sTemp = fpJson.Value
        
        fpJson.Value = Trim(Replace(fpJson.Value, """", ""))
        
        fpJson.Col = 2
        fpJson.Value = sTemp
        
        
        For i = 1 To UBound(sUpperArr)
            fpJson.Value = Replace(fpJson.Value, sUpperArr(i), "_" & sUpperArr(i))
        Next i
        
        fpJson.Value = Trim(Replace(UCase(fpJson.Value), """", ""))
        
    
    Next iRow
    

    Dim sMap() As String
    
    '단어를 전역 배열에 셋팅
    subSetMetaWord
    
    Dim k As Integer
    Dim sMatch As String
    With fpJson
        For i = 1 To fpJson.MaxRows
            .Row = i
            .Col = 2
            
            sMap = Split(.Value, "_")
            
            For iCnt = LBound(sMap) To UBound(sMap)
            
                .Row = i
                .Col = 3
                
                '.Value = .Value & IIf(getMetaDBOne(sMap(iCnt)) = "", sMap(iCnt), getMetaDBOne(sMap(iCnt)))
            
                sMatch = ""
                For k = LBound(gsEnglisWord) To UBound(gsEnglisWord)
                    If Trim(gsEnglisWord(k)) = Trim(sMap(iCnt)) Then
                        sMatch = gsHangulWord(k)
                    End If
                    
                Next k
                
                If Trim(sMatch) <> "" Then
                    .Value = .Value & sMatch
                Else
                    .Value = .Value & Trim(sMap(iCnt))
                End If
            
            
            Next iCnt

        Next i
    End With
End Sub

Private Sub Command9_Click()
    Dim i As Integer
    For i = 1 To fpJson.MaxRows
        fpJson.Col = -1
        fpJson.Row = i
        fpJson.BackColor = S_WHITE
    Next i
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.Height = 13965
    Me.Width = 25875
    
    
    SSTab1.Tab = 0
    SSTab2.Tab = 0
    SSTab3.Tab = 0
    
    fpMethod.MaxRows = 0
    
    
    Dim i As Integer
    
    '포커스갔을때 한글 설정
    For i = txtData.LBound To txtData.UBound
        Select Case i
            Case 4, 6, 9, 12, 13
                txtData(i).IMEMode = 10
            Case Else
                txtData(i).IMEMode = 8
        End Select
            
        
    Next i
    
    '초기화
    For i = rtxtBug.LBound To rtxtBug.UBound
        rtxtBug(i).Text = ""
    Next i
    
    '신규입력 상태
    Command2_Click
    
    txtUserNm.Text = gUserNm
    
    txtTableMap.Text = GetSetting("ORAIB_TAB_MAN_BY_HAK", "SQL_CONVERT", "TABLE_MAP2", "AS-IS^TO-BE" & vbCrLf & "TB_HGG115_GUARMST^TB_HGA001M_GRNT")
    
    
    '메타설정하기
    'Call subSetMetaDB(1)
End Sub

Private Sub fpBugList_Click(ByVal Col As Long, ByVal Row As Long)
    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

'On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    

    Set conn = New ADODB.Connection
    conn.ConnectionString = GetConnStringHgDev

    conn.Open

    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
        
    fpBugList.Row = Row
    fpBugList.Col = 1
        
        sSQL = ""
        sSQL = sSQL & vbCrLf & "SELECT  NO                 -- "
        sSQL = sSQL & vbCrLf & "       ,NVL(프로그램ID, ' ')  프로그램ID           -- "
        sSQL = sSQL & vbCrLf & "       ,NVL(프로그램명, ' ')   프로그램명           -- "
        sSQL = sSQL & vbCrLf & "       ,NVL(등록자, ' ')   등록자             -- "
        sSQL = sSQL & vbCrLf & "       ,NVL(개발자, ' ')   개발자             -- "
        sSQL = sSQL & vbCrLf & "       ,NVL(상태, ' ')       상태          -- "
        sSQL = sSQL & vbCrLf & "       ,NVL(결함내용, ' ')    결함내용           -- "
        sSQL = sSQL & vbCrLf & "       ,NVL(파일명, ' ')    파일명            -- "
        sSQL = sSQL & vbCrLf & "       ,등록일시               -- "
        sSQL = sSQL & vbCrLf & "       ,NVL(UI여부,'0')  AS UI여부              -- "
        sSQL = sSQL & vbCrLf & "  FROM HGDEV_T003        /* */"
        sSQL = sSQL & vbCrLf & " WHERE 1=1"
        sSQL = sSQL & vbCrLf & "   AND NO                  = " & fpBugList.Value & " "



    rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

    rtxtBug(0).Text = rsTemp.Fields("프로그램ID").Value
    rtxtBug(1).Text = rsTemp.Fields("프로그램명").Value
    rtxtBug(2).Text = rsTemp.Fields("개발자").Value
    rtxtBug(3).Text = rsTemp.Fields("등록자").Value
    rtxtBug(4).Text = rsTemp.Fields("상태").Value
    rtxtBug(5).Text = rsTemp.Fields("파일명").Value
    rtxtBug(6).Text = rsTemp.Fields("결함내용").Value
    rtxtBug(8).Text = rsTemp.Fields("NO").Value
    chkUI.Value = rsTemp.Fields("UI여부").Value
    

    Set rsTemp = Nothing
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
    sSQL = ""
    sSQL = sSQL & vbCrLf & "SELECT  NO             "
    sSQL = sSQL & vbCrLf & "       ,등록일시       "
    sSQL = sSQL & vbCrLf & "       ,등록자         "
    sSQL = sSQL & vbCrLf & "       ,상태           "
    sSQL = sSQL & vbCrLf & "       ,조치의견       "
    sSQL = sSQL & vbCrLf & "  FROM HGDEV_T004       "
    sSQL = sSQL & vbCrLf & " WHERE 1=1"
    sSQL = sSQL & vbCrLf & "   AND NO                  = " & fpBugList.Value & " "
    sSQL = sSQL & vbCrLf & "  ORDER BY 등록일시 DESC "
    
    rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText


    
    fpProcHist.MaxRows = 0
    If rsTemp.RecordCount < 1 Then
        fpProcHist.MaxRows = 1
    Else
        '조회된 데이타 디스플레이
        Call gfRsToFp(fpProcHist, rsTemp)
    End If
    
    
    Dim i As Integer
    For i = 1 To fpProcHist.MaxRows
        fpProcHist.RowHeight(i) = 15
    Next i

    cmdProcComp.Enabled = False
    cmdReject.Enabled = False
    cmdConfirm.Enabled = False

    
    
    cmdProcComp.Enabled = False
    cmdReject.Enabled = False
    cmdConfirm.Enabled = False
            
            
    '로그인 한 사람이 개발자이면
    If Trim(rtxtBug(2).Text) = gUserNm Then
        '결함등록 버튼 disable
        cmdBugReg.Enabled = False
        
        '상태가 등록 상태인 경우만
        If rtxtBug(4).Text = "등록" Then
            '조치완료 버튼 활성
            cmdProcComp.Enabled = True
            
            '반려, 확인완료 비활성
            cmdReject.Enabled = False
            cmdConfirm.Enabled = False
        End If
        
        
    
    End If
            
    '로그인 한 사람이 등록자이면이면
    If Trim(rtxtBug(3).Text) = gUserNm Then
        
        cmdBugReg.Enabled = True
        '상태가 등록 상태인 경우만
        If rtxtBug(4).Text = "조치완료" Then
            '조치완료 버튼 활성
            cmdProcComp.Enabled = False
            
            '반려, 확인완료 비활성
            cmdReject.Enabled = True
            cmdConfirm.Enabled = True
        End If
    
    
    
    End If
    
    

    
    rtxtBug(7).Text = ""
    
    
    
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

Private Sub fpDevPgm_DblClick(ByVal Col As Long, ByVal Row As Long)
    With fpDevPgm
        .Row = Row
        
        '화면id를 프로그램id로
        If Col = 2 Or Col = 3 Then
            .Col = 2
            rtxtBug(0).Text = .Value
            .Col = 3
            rtxtBug(1).Text = .Value
            .Col = 4
            rtxtBug(2).Text = .Value     '개발자
        End If
        
        
        '보고서id를 프로그램id로
        If Col = 5 Or Col = 6 Then
            .Col = 5
            rtxtBug(0).Text = .Value
            .Col = 6
            rtxtBug(1).Text = .Value
            .Col = 4
            rtxtBug(2).Text = .Value     '개발자
        End If
        
        
        '프로그램id를 프로그램id로
        If Col = 7 Or Col = 8 Then
            .Col = 7
            rtxtBug(0).Text = .Value
            .Col = 8
            rtxtBug(1).Text = .Value
            .Col = 4
            rtxtBug(2).Text = .Value     '개발자
        End If
    
    End With
End Sub

Private Sub fpMethod_Change(ByVal Col As Long, ByVal Row As Long)
    

    fpMethod.Row = Row
    fpMethod.Col = -1
    fpMethod.BackColor = C_ADD_COLOR
End Sub

Private Sub fpMethod_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i As Integer
    With fpMethod
        .Row = Row
        
        For i = 1 To .MaxCols
            .Col = i
            txtData(i - 1).Text = .Value
        
        Next i
    
    End With
End Sub

Private Sub fpMethod_DblClick(ByVal Col As Long, ByVal Row As Long)
    Call gfGrid_Sort(fpMethod, Col, Row)
End Sub

Private Sub fpProcHist_Click(ByVal Col As Long, ByVal Row As Long)
    fpProcHist.Row = Row
    fpProcHist.Col = 5
    rtxtBug(7).Text = fpProcHist.Value
End Sub



Private Sub optMig1_Click()
    txtGrntNo.Enabled = True
    txtGrntNo.Text = ""
    
    txtColId.Enabled = False
    txtColId.Text = ""
    txtColVal.Enabled = False
    txtColVal.Text = ""
    
    txtGrntNo.SetFocus
End Sub

Private Sub optMig2_Click()
    txtGrntNo.Enabled = False
    txtGrntNo.Text = ""
    
    txtColId.Enabled = True
    txtColId.Text = ""
    txtColVal.Enabled = True
    txtColVal.Text = ""
    
    txtColId.SetFocus
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
    
'엑셀에서 복사한 경우 enter key 제거
txtCodeSrch.Text = Replace(txtCodeSrch.Text, vbCrLf, "")

    sSearchKey = Trim(UCase(txtCodeSrch.Text))
    
    
    iRow = fpMap.SearchCol(3, 1, fpMap.MaxRows, sSearchKey, SearchFlagsValue)
    If iRow > 0 Then
        Call fpMap.ShowCell(1, IIf(iRow - 3 <= 1, 1, iRow - 3), PositionUpperCenter)
        fpMap.Col = -1
        fpMap.Row = iRow
        fpMap.BackColor = S_CYAN
        'Exit Sub
    End If
    
    iRow = fpMap.SearchCol(7, 1, fpMap.MaxRows, sSearchKey, SearchFlagsValue)
    If iRow > 0 Then
        
        Call fpMap.ShowCell(1, IIf(iRow - 3 <= 1, 1, iRow - 3), PositionUpperCenter)
        fpMap.Col = -1
        fpMap.Row = iRow
        fpMap.BackColor = S_CYAN
        'Exit Sub
    End If

    
    
   
    Dim iMatchCount As Integer
    Dim iFirstRow As Integer
    

    iMatchCount = 0
    iFirstRow = 0
    
    For i = 1 To fpMap.MaxRows
        fpMap.Row = i
        fpMap.Col = 3
        
        If InStr(1, Trim(fpMap.Value), sSearchKey) > 0 Then
        
            iMatchCount = iMatchCount + 1
            
            iRow = i
            
            If iFirstRow = 0 Then
                iFirstRow = iRow
            End If
            

            fpMap.Col = -1
            fpMap.Row = iRow
            fpMap.BackColor = S_LIGHT_GRAY
        
        End If
        
        
        '_ 없이 찾기
        If InStr(1, Replace(Trim(fpMap.Value), "_", ""), sSearchKey) > 0 Then
        
            iMatchCount = iMatchCount + 1
            
            iRow = i
            
            If iFirstRow = 0 Then
                iFirstRow = iRow
            End If
            

            fpMap.Col = -1
            fpMap.Row = iRow
            fpMap.BackColor = S_LIGHT_GRAY
        
        End If
        
        
        
        
        fpMap.Col = 7

        If InStr(1, Trim(fpMap.Value), sSearchKey) > 0 Then
            
            iMatchCount = iMatchCount + 1
            
            iRow = i
            
            
            If iFirstRow = 0 Then
                iFirstRow = iRow
            End If
            
            
            
            fpMap.Col = -1
            fpMap.Row = iRow
            fpMap.BackColor = S_LIGHT_GRAY
        
        End If
    
    Next i
    
    

    Call fpMap.ShowCell(1, IIf(iFirstRow - 3 <= 1, 1, iFirstRow - 3), PositionUpperCenter)

End Sub

Private Sub txtColumnSrch_GotFocus()
    txtColumnSrch.SelStart = 0
    txtColumnSrch.SelLength = Len(txtColumnSrch.Text)
End Sub

Private Sub txtColumnSrch_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim iRow As Integer
    If KeyAscii <> 13 Then
        
        Exit Sub
        
    End If
    Dim sSearchKey As String
    
'엑셀에서 복사한 경우 enter key 제거
txtColumnSrch.Text = Replace(txtColumnSrch.Text, vbCrLf, "")

    sSearchKey = Trim(UCase(txtColumnSrch.Text))
    
    
    iRow = fpDMB.SearchCol(5, 1, fpDMB.MaxRows, sSearchKey, SearchFlagsValue)
    If iRow > 0 Then
        Call fpDMB.ShowCell(5, IIf(iRow - 3 <= 1, 1, iRow - 3), PositionUpperCenter)
        fpDMB.Col = -1
        fpDMB.Row = iRow
        fpDMB.BackColor = S_CYAN
        'Exit Sub
    End If
    
    iRow = fpDMB.SearchCol(6, 1, fpDMB.MaxRows, sSearchKey, SearchFlagsValue)
    If iRow > 0 Then
        
        Call fpDMB.ShowCell(6, IIf(iRow - 3 <= 1, 1, iRow - 3), PositionUpperCenter)
        fpDMB.Col = -1
        fpDMB.Row = iRow
        fpDMB.BackColor = S_CYAN
        'Exit Sub
    End If

    
    iRow = fpDMB.SearchCol(13, 1, fpDMB.MaxRows, sSearchKey, SearchFlagsValue)
    If iRow > 0 Then
        Call fpDMB.ShowCell(13, IIf(iRow - 3 <= 1, 1, iRow - 3), PositionUpperCenter)
        fpDMB.Col = -1
        fpDMB.Row = iRow
        fpDMB.BackColor = S_CYAN
        'Exit Sub
    End If
    
    iRow = fpDMB.SearchCol(14, 1, fpDMB.MaxRows, sSearchKey, SearchFlagsValue)
    If iRow > 0 Then
        
        Call fpDMB.ShowCell(14, IIf(iRow - 3 <= 1, 1, iRow - 3), PositionUpperCenter)
        fpDMB.Col = -1
        fpDMB.Row = iRow
        fpDMB.BackColor = S_CYAN
        'Exit Sub
    End If
    
    
    
   
    Dim iMatchCount As Integer
    Dim iFirstRow As Integer
    

    iMatchCount = 0
    iFirstRow = 0
    
    For i = 1 To fpDMB.MaxRows
        fpDMB.Row = i
        fpDMB.Col = 5
        
        If InStr(1, Trim(fpDMB.Value), sSearchKey) > 0 Then
        
            iMatchCount = iMatchCount + 1
            
            iRow = i
            
            If iFirstRow = 0 Then
                iFirstRow = iRow
            End If
            

            fpDMB.Col = -1
            fpDMB.Row = iRow
            fpDMB.BackColor = S_LIGHT_GRAY
        
        End If
        
        
        '_ 없이 찾기
        If InStr(1, Replace(Trim(fpDMB.Value), "_", ""), sSearchKey) > 0 Then
        
            iMatchCount = iMatchCount + 1
            
            iRow = i
            
            If iFirstRow = 0 Then
                iFirstRow = iRow
            End If
            

            fpDMB.Col = -1
            fpDMB.Row = iRow
            fpDMB.BackColor = S_LIGHT_GRAY
        
        End If
        
        
        
        
        fpDMB.Col = 6

        If InStr(1, Trim(fpDMB.Value), sSearchKey) > 0 Then
            
            iMatchCount = iMatchCount + 1
            
            iRow = i
            
            
            If iFirstRow = 0 Then
                iFirstRow = iRow
            End If
            
            
            
            fpDMB.Col = -1
            fpDMB.Row = iRow
            fpDMB.BackColor = S_LIGHT_GRAY
        
        End If
    
    
    
    
        fpDMB.Col = 13
        
        If InStr(1, Trim(fpDMB.Value), sSearchKey) > 0 Then
        
            iMatchCount = iMatchCount + 1
            
            iRow = i
            
            If iFirstRow = 0 Then
                iFirstRow = iRow
            End If
            

            fpDMB.Col = -1
            fpDMB.Row = iRow
            fpDMB.BackColor = S_LIGHT_GRAY
        
        End If
        
        
        '_ 없이 찾기
        If InStr(1, Replace(Trim(fpDMB.Value), "_", ""), sSearchKey) > 0 Then
        
            iMatchCount = iMatchCount + 1
            
            iRow = i
            
            If iFirstRow = 0 Then
                iFirstRow = iRow
            End If
            

            fpDMB.Col = -1
            fpDMB.Row = iRow
            fpDMB.BackColor = S_LIGHT_GRAY
        
        End If
        
        
        
        
        fpDMB.Col = 14

        If InStr(1, Trim(fpDMB.Value), sSearchKey) > 0 Then
            
            iMatchCount = iMatchCount + 1
            
            iRow = i
            
            
            If iFirstRow = 0 Then
                iFirstRow = iRow
            End If
            
            
            
            fpDMB.Col = -1
            fpDMB.Row = iRow
            fpDMB.BackColor = S_LIGHT_GRAY
        
        End If
        
    
    
    
    Next i
    
    

    Call fpDMB.ShowCell(1, IIf(iFirstRow - 3 <= 1, 1, iFirstRow - 3), PositionUpperCenter)
End Sub

Private Sub txtCondition_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call cmdSearch_Click
    End If
End Sub




Private Sub makeCellBorderline()


'    'Determines the section of the cell border displayed around the entire spreadsheet
'    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_LEFT, &HC0C0C0, SS_BORDER_STYLE_SOLID
'    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_TOP, &HC0C0C0, SS_BORDER_STYLE_SOLID
'    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_RIGHT, &HC0C0C0, SS_BORDER_STYLE_SOLID
'    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_BOTTOM, &HC0C0C0, SS_BORDER_STYLE_SOLID
'
'    'Determines the section of the cell border displayed around the entire spreadsheet
'    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_LEFT, &HC0C0C0, SS_BORDER_STYLE_SOLID
'    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_TOP, &HC0C0C0, SS_BORDER_STYLE_SOLID
'    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_RIGHT, &HC0C0C0, SS_BORDER_STYLE_SOLID
'    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_BOTTOM, &HC0C0C0, SS_BORDER_STYLE_SOLID
'
'    'Determines the section of the cell border displayed around the entire spreadsheet
'    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_LEFT, &HC0C0C0, SS_BORDER_STYLE_SOLID
'    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_TOP, &HC0C0C0, SS_BORDER_STYLE_SOLID
'    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_RIGHT, &HC0C0C0, SS_BORDER_STYLE_SOLID
'    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_BOTTOM, &HC0C0C0, SS_BORDER_STYLE_SOLID
'
'    'Determines the section of the cell border displayed around the entire spreadsheet
'    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_LEFT, &HC0C0C0, SS_BORDER_STYLE_SOLID
'    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_TOP, &HC0C0C0, SS_BORDER_STYLE_SOLID
'    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_RIGHT, &HC0C0C0, SS_BORDER_STYLE_SOLID
'    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_BOTTOM, &HC0C0C0, SS_BORDER_STYLE_SOLID
'
'    'Determines the section of the cell border displayed around the entire spreadsheet
'    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_LEFT, &HC0C0C0, SS_BORDER_STYLE_SOLID
'    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_TOP, &HC0C0C0, SS_BORDER_STYLE_SOLID
'    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_RIGHT, &HC0C0C0, SS_BORDER_STYLE_SOLID
'    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_BOTTOM, &HC0C0C0, SS_BORDER_STYLE_SOLID

    'Determines the section of the cell border displayed around the entire spreadsheet
    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_LEFT, &H0&, SS_BORDER_STYLE_SOLID
    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_TOP, &H0&, SS_BORDER_STYLE_SOLID
    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_RIGHT, &H0&, SS_BORDER_STYLE_SOLID
    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_BOTTOM, &H0&, SS_BORDER_STYLE_SOLID

    'Determines the section of the cell border displayed around the entire spreadsheet
    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_LEFT, &H0&, SS_BORDER_STYLE_SOLID
    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_TOP, &H0&, SS_BORDER_STYLE_SOLID
    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_RIGHT, &H0&, SS_BORDER_STYLE_SOLID
    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_BOTTOM, &H0&, SS_BORDER_STYLE_SOLID

    'Determines the section of the cell border displayed around the entire spreadsheet
    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_LEFT, &H0&, SS_BORDER_STYLE_SOLID
    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_TOP, &H0&, SS_BORDER_STYLE_SOLID
    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_RIGHT, &H0&, SS_BORDER_STYLE_SOLID
    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_BOTTOM, &H0&, SS_BORDER_STYLE_SOLID

    'Determines the section of the cell border displayed around the entire spreadsheet
    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_LEFT, &H0&, SS_BORDER_STYLE_SOLID
    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_TOP, &H0&, SS_BORDER_STYLE_SOLID
    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_RIGHT, &H0&, SS_BORDER_STYLE_SOLID
    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_BOTTOM, &H0&, SS_BORDER_STYLE_SOLID

    'Determines the section of the cell border displayed around the entire spreadsheet
    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_LEFT, &H0&, SS_BORDER_STYLE_SOLID
    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_TOP, &H0&, SS_BORDER_STYLE_SOLID
    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_RIGHT, &H0&, SS_BORDER_STYLE_SOLID
    fpMethod.SetCellBorder 1, 1, 15, 1, SS_BORDER_TYPE_BOTTOM, &H0&, SS_BORDER_STYLE_SOLID



End Sub



Private Sub txtDevName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cdmDevSearch_Click
    End If
End Sub

Private Sub txtDevPgm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cdmDevSearch_Click
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim iRow As Integer
    If KeyAscii <> 13 Then
        
        Exit Sub
        
    End If
    Dim sSearchKey As String
    
'엑셀에서 복사한 경우 enter key 제거
txtFind.Text = Replace(txtFind.Text, vbCrLf, "")

    sSearchKey = Trim(UCase(txtFind.Text))
    
    
    iRow = fpJson.SearchCol(3, 1, fpJson.MaxRows, sSearchKey, SearchFlagsValue)
    If iRow > 0 Then
        Call fpJson.ShowCell(3, IIf(iRow - 3 <= 1, 1, iRow - 3), PositionUpperCenter)
        fpJson.Col = -1
        fpJson.Row = iRow
        fpJson.BackColor = S_CYAN
        'Exit Sub
    End If
    
    iRow = fpJson.SearchCol(1, 1, fpJson.MaxRows, sSearchKey, SearchFlagsValue)
    If iRow > 0 Then
        
        Call fpJson.ShowCell(1, IIf(iRow - 3 <= 1, 1, iRow - 3), PositionUpperCenter)
        fpJson.Col = -1
        fpJson.Row = iRow
        fpJson.BackColor = S_CYAN
        'Exit Sub
    End If
'
'
'    iRow = fpJson.SearchCol(13, 1, fpJson.MaxRows, sSearchKey, SearchFlagsValue)
'    If iRow > 0 Then
'        Call fpJson.ShowCell(13, IIf(iRow - 3 <= 1, 1, iRow - 3), PositionUpperCenter)
'        fpJson.Col = -1
'        fpJson.Row = iRow
'        fpJson.BackColor = S_CYAN
'        'Exit Sub
'    End If
'
'    iRow = fpJson.SearchCol(14, 1, fpJson.MaxRows, sSearchKey, SearchFlagsValue)
'    If iRow > 0 Then
'
'        Call fpJson.ShowCell(14, IIf(iRow - 3 <= 1, 1, iRow - 3), PositionUpperCenter)
'        fpJson.Col = -1
'        fpJson.Row = iRow
'        fpJson.BackColor = S_CYAN
'        'Exit Sub
'    End If
'
    
    
   
    Dim iMatchCount As Integer
    Dim iFirstRow As Integer
    

    iMatchCount = 0
    iFirstRow = 0
    
    For i = 1 To fpJson.MaxRows
        fpJson.Row = i
        fpJson.Col = 3
        
        If InStr(1, Trim(fpJson.Value), sSearchKey) > 0 Then
        
            iMatchCount = iMatchCount + 1
            
            iRow = i
            
            If iFirstRow = 0 Then
                iFirstRow = iRow
            End If
            

            fpJson.Col = -1
            fpJson.Row = iRow
            fpJson.BackColor = S_LIGHT_GRAY
        
        End If
        
        
        '_ 없이 찾기
        If InStr(1, Replace(Trim(fpJson.Value), "_", ""), sSearchKey) > 0 Then
        
            iMatchCount = iMatchCount + 1
            
            iRow = i
            
            If iFirstRow = 0 Then
                iFirstRow = iRow
            End If
            

            fpJson.Col = -1
            fpJson.Row = iRow
            fpJson.BackColor = S_LIGHT_GRAY
        
        End If
        
        
        
        
        fpJson.Col = 1

        If InStr(1, Trim(fpJson.Value), sSearchKey) > 0 Then
            
            iMatchCount = iMatchCount + 1
            
            iRow = i
            
            
            If iFirstRow = 0 Then
                iFirstRow = iRow
            End If
            
            
            
            fpJson.Col = -1
            fpJson.Row = iRow
            fpJson.BackColor = S_LIGHT_GRAY
        
        End If
    
    
'
'
'        fpJson.Col = 13
'
'        If InStr(1, Trim(fpJson.Value), sSearchKey) > 0 Then
'
'            iMatchCount = iMatchCount + 1
'
'            iRow = i
'
'            If iFirstRow = 0 Then
'                iFirstRow = iRow
'            End If
'
'
'            fpJson.Col = -1
'            fpJson.Row = iRow
'            fpJson.BackColor = S_LIGHT_GRAY
'
'        End If
'
'
'        '_ 없이 찾기
'        If InStr(1, Replace(Trim(fpJson.Value), "_", ""), sSearchKey) > 0 Then
'
'            iMatchCount = iMatchCount + 1
'
'            iRow = i
'
'            If iFirstRow = 0 Then
'                iFirstRow = iRow
'            End If
'
'
'            fpJson.Col = -1
'            fpJson.Row = iRow
'            fpJson.BackColor = S_LIGHT_GRAY
'
'        End If
'
'
'
'
'        fpJson.Col = 14
'
'        If InStr(1, Trim(fpJson.Value), sSearchKey) > 0 Then
'
'            iMatchCount = iMatchCount + 1
'
'            iRow = i
'
'
'            If iFirstRow = 0 Then
'                iFirstRow = iRow
'            End If
'
'
'
'            fpJson.Col = -1
'            fpJson.Row = iRow
'            fpJson.BackColor = S_LIGHT_GRAY
'
'        End If
'
'
    
    
    Next i
    
    

    Call fpJson.ShowCell(1, IIf(iFirstRow - 3 <= 1, 1, iFirstRow - 3), PositionUpperCenter)
End Sub

Private Sub txtPGMid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSearch2_Click
    End If
End Sub

Private Sub txtPGMNM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSearch2_Click
    End If
End Sub

Private Sub txtTableMap_LostFocus()
    Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", "SQL_CONVERT", "TABLE_MAP2", txtTableMap.Text)
End Sub

Private Sub txtUserNm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSearch2_Click
    End If
End Sub
