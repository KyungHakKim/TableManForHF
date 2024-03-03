VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Begin VB.Form frmDataView 
   BackColor       =   &H80000005&
   Caption         =   "DataView"
   ClientHeight    =   12480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17670
   LinkTopic       =   "Form1"
   ScaleHeight     =   12480
   ScaleWidth      =   17670
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Excel"
      Height          =   495
      Left            =   16020
      Style           =   1  '그래픽
      TabIndex        =   10
      Top             =   600
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   12255
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   17295
      _ExtentX        =   30506
      _ExtentY        =   21616
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "DataView"
      TabPicture(0)   =   "frmDataView.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblHint"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fpDataView"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkAll"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "edtTablenm"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "edtTable"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdBookMakr"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtSrchId"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdUpdate"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "UpdateSQL"
      TabPicture(1)   =   "frmDataView.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txUpdateSQL"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmDataView.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
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
         Height          =   11355
         IMEMode         =   8  '영문
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   9
         Top             =   600
         Width           =   16635
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H80000016&
         Caption         =   "UPDATE SQL"
         Height          =   375
         Left            =   8820
         Style           =   1  '그래픽
         TabIndex        =   8
         Top             =   660
         Width           =   1875
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
         Left            =   1140
         TabIndex        =   5
         Top             =   780
         Width           =   3195
      End
      Begin VB.CommandButton cmdBookMakr 
         BackColor       =   &H80000016&
         Caption         =   "BookMark Reset"
         Height          =   375
         Left            =   4440
         Style           =   1  '그래픽
         TabIndex        =   4
         Top             =   720
         Width           =   1875
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
         Left            =   10740
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   2415
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
         Left            =   13140
         TabIndex        =   2
         Top             =   720
         Width           =   2475
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "Check1"
         Height          =   255
         Left            =   720
         TabIndex        =   1
         Top             =   1200
         Width           =   315
      End
      Begin FPSpreadADO.fpSpread fpDataView 
         Height          =   10935
         Left            =   180
         TabIndex        =   6
         Top             =   1140
         Width           =   16935
         _Version        =   393216
         _ExtentX        =   29871
         _ExtentY        =   19288
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
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
         GrayAreaBackColor=   14737632
         MaxCols         =   6
         MaxRows         =   50
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ShadowColor     =   12648447
         SpreadDesigner  =   "frmDataView.frx":0054
         ScrollBarTrack  =   1
      End
      Begin VB.Label lblHint 
         BackStyle       =   0  '투명
         Caption         =   "....."
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6900
         TabIndex        =   7
         Top             =   720
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmDataView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAll_Click()
    Dim i As Integer
    For i = 1 To fpDataView.MaxRows
        fpDataView.Row = i
        fpDataView.Col = 1
        If chkAll.Value = 1 Then
            fpDataView.Value = 1
        Else
            fpDataView.Value = 0
        End If
    Next i
End Sub

Private Sub cmdBookMakr_Click()
    Dim i As Integer
    For i = 1 To fpDataView.MaxRows
        fpDataView.Row = i
        fpDataView.Col = 5

        If fpDataView.Value <> "" Then
            fpDataView.Col = -1

            fpDataView.BackColor = &HFFC0FF
        Else
            fpDataView.Col = -1
            fpDataView.BackColor = S_WHITE
        End If



    Next i
End Sub

Private Sub cmdExcel_Click()
     Call gfFpToExcel(fpDataView)
End Sub

Private Sub cmdUpdate_Click()

    Dim sCol() As String
    Dim sWhere() As String
    Dim i, J As Integer
    Dim iCnt As Integer
    txUpdateSQL.Text = ""
    iCnt = 0
    Screen.MousePointer = vbHourglass
    With fpDataView
        For J = C_DADA_VIEW_BASIC_COL_SIZE + 1 To fpDataView.MaxCols
            ReDim sCol(fpDataView.MaxRows)
            ReDim sWhere(fpDataView.MaxRows)
            For i = 1 To fpDataView.MaxRows
                .Row = i
                .Col = 1

                '갱신필드 체크
                If .Value = 1 Then
                    iCnt = iCnt + 1
                    .Col = 4
                    If .Value = "VARCHAR2" Or .Value = "DATE" Then

                        .Col = 2
                        sCol(i) = .Value & " = "
                        .Col = J
                        sCol(i) = sCol(i) & "'" & Trim(.Value) & "'"
                        .Col = 3
                        sCol(i) = sCol(i) & "                -- " & .Value
                    Else
                        .Col = 2
                        sCol(i) = .Value & " = "
                        .Col = J
                        sCol(i) = sCol(i) & Trim(.Value)
                        .Col = 3
                        sCol(i) = sCol(i) & "                -- " & .Value

                    End If
                End If

                'Where 절 체크
                .Col = 5
                If Trim(.Value) <> "" Then
                    .Col = 4
                    If .Value = "VARCHAR2" Or .Value = "DATE" Then

                        .Col = 2
                        sWhere(i) = .Value & " = "
                        .Col = J
                        sWhere(i) = sWhere(i) & "'" & Trim(.Value) & "'"
                        .Col = 3
                        sWhere(i) = sWhere(i) & "                -- " & .Value

                    Else
                        .Col = 2
                        sWhere(i) = .Value & " = "
                        .Col = J
                        sWhere(i) = sWhere(i) & .Value
                        .Col = 3
                        sWhere(i) = sWhere(i) & "                -- " & .Value

                    End If
                End If

            Next i

            If iCnt = 0 Then
                gfMsg ("변경 컬럼을 선택하세요")
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            Dim bFirst As Boolean
            txUpdateSQL.Text = txUpdateSQL.Text & "UPDATE " & edtTable.Text & vbCrLf
            bFirst = True
            For i = 1 To .MaxRows
                If Trim(sCol(i)) <> "" Then
                    If bFirst Then
                        txUpdateSQL.Text = txUpdateSQL.Text & "   SET " & sCol(i) & vbCrLf
                    Else
                        txUpdateSQL.Text = txUpdateSQL.Text & "     , " & sCol(i) & vbCrLf
                    End If

                    bFirst = False
                End If


            Next i

             txUpdateSQL.Text = txUpdateSQL.Text & "     , SYS_CHG_DTTM = SYSDATE              --시스템변경일시" & vbCrLf
            txUpdateSQL.Text = txUpdateSQL.Text & " WHERE 1=1 " & vbCrLf
            For i = 1 To .MaxRows
                If Trim(sWhere(i)) <> "" Then
                    txUpdateSQL.Text = txUpdateSQL.Text & "   AND " & sWhere(i) & vbCrLf
                End If
            Next i

            txUpdateSQL.Text = txUpdateSQL.Text & ";" & vbCrLf


        Next J
    End With
    SSTab1.Tab = 1
    Screen.MousePointer = vbDefault
End Sub


Private Sub fpDataView_Click(ByVal Col As Long, ByVal Row As Long)

    fpDataView.Col = Col
    fpDataView.Row = Row


    'rscode.Filter = adFilterNone
    'rscode.Filter = "SMPL_SCCD='" & fpDataView.Value & "'"

On Error Resume Next
    'lblHint.Caption = rscode.Fields(5)
End Sub

Private Sub txtSrchId_GotFocus()
    txtSrchId.SelStart = 0
    txtSrchId.SelLength = Len(txtSrchId.Text)
End Sub

Private Sub txtSrchId_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim iRow As Integer
    If KeyAscii <> 13 Then

        Exit Sub

    End If
    Dim sSearchKey As String
    sSearchKey = Trim(UCase(txtSrchId.Text))


    iRow = fpDataView.SearchCol(2, 1, fpDataView.MaxRows, sSearchKey, SearchFlagsValue)
    If iRow > 0 Then
        Call fpDataView.ShowCell(1, IIf(iRow - 3 <= 1, 1, iRow - 3), PositionUpperCenter)
        fpDataView.Col = -1
        fpDataView.Row = iRow
        fpDataView.BackColor = S_CYAN
        Exit Sub
    End If

    iRow = fpDataView.SearchCol(3, 1, fpDataView.MaxRows, sSearchKey, SearchFlagsValue)
    If iRow > 0 Then

        Call fpDataView.ShowCell(1, IIf(iRow - 3 <= 1, 1, iRow - 3), PositionUpperCenter)
        fpDataView.Col = -1
        fpDataView.Row = iRow
        fpDataView.BackColor = S_CYAN

        Exit Sub
    End If




    Dim iMatchCount As Integer
    Dim iFirstRow As Integer


    iMatchCount = 0
    iFirstRow = 0

    For i = 1 To fpDataView.MaxRows
        fpDataView.Row = i
        fpDataView.Col = 2

        If InStr(1, Trim(fpDataView.Value), sSearchKey) > 0 Then

            iMatchCount = iMatchCount + 1

            iRow = i

            If iFirstRow = 0 Then
                iFirstRow = iRow
            End If


            fpDataView.Col = -1
            fpDataView.Row = iRow
            fpDataView.BackColor = S_LIGHT_GRAY


        End If

        fpDataView.Col = 3

        If InStr(1, Trim(fpDataView.Value), sSearchKey) > 0 Then

            iMatchCount = iMatchCount + 1

            iRow = i


            If iFirstRow = 0 Then
                iFirstRow = iRow
            End If



            fpDataView.Col = -1
            fpDataView.Row = iRow
            fpDataView.BackColor = S_LIGHT_GRAY

        End If

    Next i





    Call fpDataView.ShowCell(1, IIf(iFirstRow - 3 <= 1, 1, iFirstRow - 3), PositionUpperCenter)


'    If iMatchCount > 1 Then
'        lblMatch.Caption = "Matched " & iMatchCount & " columns...."
'        picMatch.Visible = True
'    End If

End Sub
