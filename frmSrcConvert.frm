VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSrcConvert 
   BackColor       =   &H00D6E9FE&
   Caption         =   "Source Converter"
   ClientHeight    =   13590
   ClientLeft      =   1065
   ClientTop       =   795
   ClientWidth     =   25845
   BeginProperty Font 
      Name            =   "����ü"
      Size            =   9
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSrcConvert.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13590
   ScaleWidth      =   25845
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2655
      Left            =   23580
      TabIndex        =   15
      Top             =   4020
      Visible         =   0   'False
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   4683
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmSrcConvert.frx":048A
   End
   Begin VB.CommandButton cmdVOsetReplace 
      Caption         =   "VO set ġȯ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   23460
      TabIndex        =   14
      Top             =   8640
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "clipboard"
      Height          =   495
      Left            =   17280
      TabIndex        =   13
      Top             =   0
      Width           =   2235
   End
   Begin VB.TextBox txtTgtSrc 
      Height          =   6615
      Left            =   13020
      MultiLine       =   -1  'True
      ScrollBars      =   3  '�����
      TabIndex        =   12
      Top             =   480
      Width           =   12675
   End
   Begin VB.TextBox txtOrgSRc 
      Height          =   6615
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  '�����
      TabIndex        =   11
      Text            =   "frmSrcConvert.frx":0521
      Top             =   480
      Width           =   12735
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6195
      Left            =   1500
      TabIndex        =   6
      Top             =   7260
      Width           =   21615
      _ExtentX        =   38126
      _ExtentY        =   10927
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Variable"
      TabPicture(0)   =   "frmSrcConvert.frx":2638
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fpVariable"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtType"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Value Object"
      TabPicture(1)   =   "frmSrcConvert.frx":2654
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fpVOList"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmSrcConvert.frx":2670
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fpVO"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.TextBox txtType 
         Appearance      =   0  '���
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   10  '�ѱ� 
         Left            =   180
         TabIndex        =   8
         Top             =   4380
         Width           =   10935
      End
      Begin FPSpreadADO.fpSpread fpVariable 
         Height          =   3975
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   10995
         _Version        =   393216
         _ExtentX        =   19394
         _ExtentY        =   7011
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "���� ���"
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
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ShadowColor     =   12648447
         ShadowDark      =   12632256
         SpreadDesigner  =   "frmSrcConvert.frx":268C
         ScrollBarTrack  =   1
      End
      Begin FPSpreadADO.fpSpread fpVOList 
         Height          =   5595
         Left            =   -74820
         TabIndex        =   9
         Top             =   420
         Width           =   21135
         _Version        =   393216
         _ExtentX        =   37280
         _ExtentY        =   9869
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "���� ���"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   12640511
         MaxCols         =   8
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ShadowColor     =   12648447
         ShadowDark      =   12632256
         SpreadDesigner  =   "frmSrcConvert.frx":2A8F
         ScrollBarTrack  =   1
      End
      Begin FPSpreadADO.fpSpread fpVO 
         Height          =   4095
         Left            =   -74880
         TabIndex        =   10
         Top             =   540
         Width           =   5475
         _Version        =   393216
         _ExtentX        =   9657
         _ExtentY        =   7223
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "���� ���"
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
         SpreadDesigner  =   "frmSrcConvert.frx":2FBE
         ScrollBarTrack  =   1
      End
   End
   Begin VB.CommandButton cmdVarChange 
      Caption         =   "����ġȯ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   23460
      TabIndex        =   5
      Top             =   7980
      Width           =   1935
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "���� ��Ī"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   9360
      Width           =   1215
   End
   Begin VB.CommandButton cmdVarLIst 
      Caption         =   "������ϰ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   10380
      Width           =   1215
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "copy"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   23460
      TabIndex        =   2
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   19560
      TabIndex        =   1
      Top             =   0
      Width           =   3135
   End
   Begin VB.CommandButton cmdGetReturn 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   120
      TabIndex        =   0
      Top             =   7260
      Width           =   1215
   End
End
Attribute VB_Name = "frmSrcConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private gsVarType() As String
Private gsVOlist(100) As String

Private gsASIS_ID() As String
Private gsASIS_NM() As String
Private gsTOBE_NM() As String
Private gsCNT() As String


'------------------------------------------
' ���� ��Ͽ��� ��Ī�ϱ�
'------------------------------------------
Private Sub cmdChange_Click()
On Error GoTo ErrHandler
    Dim ffile As Integer
    Dim sTemp  As String
    Dim sTemp2() As String
    Dim iCnt As Integer
    Dim i, j As Integer
    Dim sArr1(), sArr2(), sArr3(), sArr4() As String
    

    ffile = FreeFile

    Open App.Path & "\" & "����ġȯ.txt" For Input As #ffile   ' ������ ���ϴ�.
    
    iCnt = 0
    Do
    
        Input #ffile, sTemp
        
        sTemp2 = Split(sTemp, ";")
        
        iCnt = iCnt + 1
        
        ReDim Preserve sArr1(1 To iCnt)
        ReDim Preserve sArr2(1 To iCnt)
        ReDim Preserve sArr3(1 To iCnt)
        ReDim Preserve sArr4(1 To iCnt)
        
        
        sArr1(iCnt) = sTemp2(0)
        sArr2(iCnt) = sTemp2(1)
        sArr3(iCnt) = sTemp2(2)
        sArr4(iCnt) = sTemp2(3)
        
        
    Loop Until EOF(1)
    
    Close #ffile
    
    
    Dim sCol1, sCol2 As String
    For i = 1 To fpVariable.MaxRows
    
        fpVariable.Row = i
        fpVariable.Col = 1
        sCol1 = fpVariable.Value
        
        fpVariable.Col = 2
        sCol2 = fpVariable.Value
        
        For j = 1 To UBound(sArr1)
            If UCase(Trim(sCol1)) = UCase(Trim(sArr1(j))) And _
               UCase(Trim(sCol2)) = UCase(Trim(sArr2(j))) Then
                           
                fpVariable.Col = 3
                fpVariable.Value = sArr3(j)
                
                fpVariable.Col = 4
                fpVariable.Value = sArr4(j)
               
            End If
        Next j
    
    Next i
    
    
    
    
    
    Exit Sub
ErrHandler:
    gfMsgInf (Err.Description)

End Sub

Private Sub cmdClear_Click()
    txtOrgSRc.Text = ""
    txtTgtSrc.Text = ""
    fpVariable.MaxRows = 0
    fpVOList.MaxRows = 0
    fpVO.MaxRows = 0
End Sub

Private Sub cmdCopy_Click()
    txtTgtSrc.Text = txtOrgSRc.Text
End Sub


'-----------------------------------------------
'������ ���� ���
'-----------------------------------------------
Private Sub cmdGetReturn_Click()
    Dim i, j As Integer
    Dim sSrcArr() As String
    Dim sWordArr() As String
    Dim sTemp As String
    
    If fpVariable.MaxRows > 0 Or fpVO.MaxRows > 0 Then
        If gfMsgQst("������ ��Ī �� ��� ��Ī ������ �ʱ�ȭ �˴ϴ�. " & vbCrLf & "������ ���� �Ͻðڽ��ϱ�?") = vbNo Then
            Exit Sub
        End If
    End If
    
    
    '����Ÿ������
    gsVarType = Split(txtType.Text, ",")
    
    Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", "SORURCE_CONVERT", "VAR_TYPE", Trim(txtType.Text))
    

    
    '�ҽ��� ���δ����� �и��Ѵ�.
    sSrcArr = Split(txtOrgSRc.Text, vbCrLf)
    
    fpVariable.MaxRows = 0
    
    Dim iVoCnt As Integer
    iVoCnt = 0
    
    For i = 1 To 100
        gsVOlist(i) = ""
    Next i
    
    '���� �� ��ŭ �ݺ�
    For i = LBound(sSrcArr) To UBound(sSrcArr)
            
        '������ �����ϱ� ���Ͽ� ���� ������ �����Ͽ� �迭�� ����
        '��ȣ�� ������ ����
        sTemp = sSrcArr(i)
        sTemp = Replace(sTemp, "(", " ( ")
        sTemp = Replace(sTemp, ")", " ( ")
        sTemp = Replace(sTemp, "  ", " ")
        
        sWordArr = Split(sTemp, " ")
        
        '�ܾ�� ��ŭ �ݺ�
        For j = LBound(sWordArr) To UBound(sWordArr)
                            
            '���� Ÿ�Կ� �ش��ϸ� ���� �ܾ�� ������
            If fFindVariableType(sWordArr(j)) = True Then
                
                'new Object() ���� ����
                If Trim(sWordArr(j + 1)) <> "(" Then
                    fpVariable.MaxRows = fpVariable.MaxRows + 1
                    
                    fpVariable.Row = fpVariable.MaxRows
                    fpVariable.Col = 1
                    
                    fpVariable.Value = sWordArr(j)
                        
                    j = j + 1
                    fpVariable.Col = 2
                    fpVariable.Value = sWordArr(j)
                End If
                
                'Value Object�� ������ ����
                fpVariable.Col = 1
                If UCase(fpVariable.Value) = UCase("ValueObject") Then
                   'Or UCase(fpVariable.Value) = UCase("ValueObjectAssembler")
                   
                        
                    fpVariable.Col = 2
                    iVoCnt = iVoCnt + 1
                    gsVOlist(iVoCnt) = fpVariable.Value
                                                           
                End If
            End If
    
    
        Next j
    
    Next i
    
    'vo�Ӽ��� find�Ͽ� �׸��忡 �����Ѵ�.
    Call subSetVoList
    
    
End Sub

'-------------------------------------------
'VO ����� ã�´�.
'-------------------------------------------
Private Sub subSetVoList()
    Dim i, j, K As Integer
    Dim sSrcArr() As String
    Dim sTemp As String
    Dim sVO1() As String
    Dim sVO2() As String
    Dim iVoCnt As Integer
    
    
    For K = LBound(gsVOlist) To UBound(gsVOlist)
        
        '�迭�� ������ �ƴϸ� �ҽ� ã��
        If Trim(gsVOlist(K)) <> "" Then
        
                    
                    
            '�ҽ��� ���δ����� �и��Ѵ�.
            sSrcArr = Split(txtOrgSRc.Text, vbCrLf)
            
           '���� �� ��ŭ �ݺ�
           For i = LBound(sSrcArr) To UBound(sSrcArr)
                   
                sTemp = sSrcArr(i)
               '������ ������ ����
               sTemp = Replace(sTemp, " ", "")
               
               
               '�ּ� ������ ����
               If Left(sTemp, 2) <> "//" And Left(sTemp, 2) <> "/*" And Right(sTemp, 2) <> "*/" Then
                
                
                    '���忡 VO��� Set�� ������
                    If InStr(1, sTemp, Trim(gsVOlist(K)) & ".set") Then
                    
                         
                         
                        'vo��.set(" ����
                         sTemp = Replace(sTemp, Trim(gsVOlist(K)) & ".set(""", "")
                         
                         iVoCnt = iVoCnt + 1
                         ReDim Preserve sVO1(1 To iVoCnt)
                         ReDim Preserve sVO2(1 To iVoCnt)
                         
                         sVO1(iVoCnt) = gsVOlist(K)
                         sVO2(iVoCnt) = Mid(sTemp, 1, InStr(1, sTemp, """") - 1)
                    End If
                    
               End If
                   
           Next i
        
        End If 'end If Trim(gsVOlist) <> "" Then
    
    Next K
        
        
    
    fpVOList.MaxRows = 0
    Dim bFind As Boolean
    
    Dim sT1, sT2 As String
    
    Dim sASISNM As String
    Dim sTOBENM As String
    
    Dim sASISNM_SOLO As String
    Dim sTOBENM_SOLO As String
    
    
    
    'Vo �Ӽ��� �ִ� ��츸 ����
    If iVoCnt > 0 Then
        
        '������ ���Ͽ� vo��� ���� ����
        For K = LBound(gsVOlist) To UBound(gsVOlist)
        

            For i = 1 To UBound(sVO1)
                
                If gsVOlist(K) = sVO1(i) Then
                
                    bFind = False
                    '�̹� �׸��忡 �ִ� �� �Ǵ�.
                    For j = 1 To fpVOList.MaxRows
                            
                        fpVOList.Row = j
                        fpVOList.Col = 1
                        sT1 = fpVOList.Value
                        
                        fpVOList.Col = 2
                        sT2 = fpVOList.Value
                    
                        If Trim(sT1) = Trim(sVO1(i)) And Trim(sT2) = Trim(sVO2(i)) Then
                            bFind = True
                            Exit For
                        End If
                    Next j
                    
                    If Not bFind Then
                        fpVOList.MaxRows = fpVOList.MaxRows + 1
                        fpVOList.Col = 1
                        fpVOList.Value = Trim(sVO1(i))
                        
                        fpVOList.Col = 2
                        fpVOList.Value = Trim(sVO2(i))
                    End If
                End If
            
            Next i
        Next K
        
        
        'AS-IS ��Ī�� �´� �ѱ۸�Ī�� ã�´�.
        For i = 1 To fpVOList.MaxRows
            With fpVOList
                .Row = i
                .Col = 2
                
                Call subFindASISid(.Value, sASISNM, sTOBENM, sASISNM_SOLO, sTOBENM_SOLO)
                
                .Col = 5
                .Value = sASISNM_SOLO
                .Col = 6
                .Value = sTOBENM_SOLO
                                
                                
                                
                .Col = 7
                .Value = sASISNM
                .Col = 8
                .Value = sTOBENM
            End With
        Next i
    End If
    
    
End Sub
'-----------------------------------------
'as-is id�� �´� �ѱ� ��Ī�� ã�´�.
'-----------------------------------------
Private Sub subFindASISid(pFindNm As String, pASISNM As String, pTOBENM As String, pASISNM_SOLO As String, pTOBENM_SOLO As String)
    Dim i As Integer
    
    pASISNM = ""
    pTOBENM = ""
    pASISNM_SOLO = ""
    pTOBENM_SOLO = ""
    
    Dim iMaxCnt As Integer
    
    iMaxCnt = 0
    
    For i = LBound(gsASIS_ID) To UBound(gsASIS_ID)
            
        If Trim(UCase(pFindNm)) = Trim(UCase(gsASIS_ID(i))) Then
            '��ǥ���� ã�´�.
            If gsCNT(i) > iMaxCnt Then
                pASISNM_SOLO = gsASIS_NM(i)
                pTOBENM_SOLO = gsTOBE_NM(i)
                iMaxCnt = gsCNT(i)
            End If
        
        
            If Trim(pASISNM) = "" Then
                pASISNM = gsASIS_NM(i)
                pTOBENM = gsTOBE_NM(i)
            Else
                pASISNM = pASISNM & "@" & gsASIS_NM(i)
                pTOBENM = pTOBENM & "@" & gsTOBE_NM(i)
            End If
        End If
    Next i
End Sub







'------------------------------------------
'������ ġȯ�Ѵ�.
'------------------------------------------
Private Sub cmdVarChange_Click()
    
    
On Error GoTo Err_Handler
    If Trim(txtTgtSrc.Text) = "" Then
        gfMsgInf ("��ȯ ��� �ҽ��� ī���ϼ���")
        Exit Sub
    End If
    
    
    Dim i, j As Integer
    Dim sSrcArr() As String
    Dim sWordArr() As String
    Dim sTemp As String

    
    '�ҽ��� ���δ����� �и��Ѵ�.
    sSrcArr = Split(txtTgtSrc.Text, vbCrLf)
    
    
    Dim iVoCnt As Integer
    iVoCnt = 0
    

    
    Dim sLastSentence As String
    sLastSentence = ""
    Dim iVarCnt As Integer
    
    Dim sTy1, sTy2 As String
    '���� �� ��ŭ �ݺ�
    For i = LBound(sSrcArr) To UBound(sSrcArr)
            
        '������ �����ϱ� ���Ͽ� ���� ������ �����Ͽ� �迭�� ����
        '��ȣ�� ������ ����
        sTemp = sSrcArr(i)
        'sTemp = Replace(sTemp, "(", " ( ")
        'sTemp = Replace(sTemp, ")", " ( ")
        'sTemp = Replace(sTemp, "  ", " ")
        
        sWordArr = Split(sTemp, " ")
        
        '�ܾ�� ��ŭ �ݺ�
        For j = LBound(sWordArr) To UBound(sWordArr)
                            
            '���� Ÿ�Կ� �ش��ϸ� ���� �ܾ�� ������
            If fFindVariableType(sWordArr(j)) = True Then
                
                'new Object() ���� ����
                If Trim(sWordArr(j + 1)) <> "(" Then
                       
                    
                       
                    'ġȯ ��� ������ �˻�
                    For iVarCnt = 1 To fpVariable.MaxRows
                        
                        fpVariable.Row = iVarCnt
                        fpVariable.Col = 1
                        sTy1 = fpVariable.Value
                        
                        fpVariable.Col = 2
                        sTy2 = fpVariable.Value
                         
                        If sWordArr(j) = sTy1 And sWordArr(j + 1) = sTy2 Then
                            
                            'to be�� ������ �� ��츸
                            fpVariable.Col = 3
                            If fpVariable.Value <> "" Then
                                sTemp = Replace(sTemp, sTy1, fpVariable.Value)
                            End If
                            
                            fpVariable.Col = 4
                            If fpVariable.Value <> "" Then
                                sTemp = Replace(sTemp, sTy2, fpVariable.Value)
                            End If
                        End If
                    
                    Next iVarCnt
                End If
            End If
    
    
        Next j
        

        sLastSentence = sLastSentence & sTemp + vbCrLf
        
        
    Next i
    
    
    txtTgtSrc.Text = sLastSentence
    
    Exit Sub
    
Err_Handler:
    gfMsgInf (Err.Description)
    
    
End Sub

'------------------------------------------
'���� ��� ���� �˾� ȣ��
'------------------------------------------
Private Sub cmdVarLIst_Click()
    frmVarList.Show vbModal
End Sub

'------------------------------------------
'VO Set Replace
'------------------------------------------
Private Sub cmdVOsetReplace_Click()
    
    
On Error GoTo Err_Handler
    If Trim(txtTgtSrc.Text) = "" Then
        gfMsgInf ("��ȯ ��� �ҽ��� ī���ϼ���")
        Exit Sub
    End If
    
    
    
    Dim i, j, K As Integer
    Dim sSrcArr() As String
    Dim sWordArr() As String
    Dim sTemp As String
    
    
    
    

    
    '�ҽ��� ���δ����� �и��Ѵ�.
    sSrcArr = Split(txtTgtSrc.Text, vbCrLf)
    
    
    Dim iVoCnt As Integer
    iVoCnt = 0
    

    
    Dim sLastSentence As String
    sLastSentence = ""
    Dim iVarCnt As Integer
    
    Dim sTy1, sTy2 As String
    
    Dim sOldVo As String
    Dim sOldItem As String
    Dim sNewVo As String
    Dim sNewItem As String
    
    Dim sOrgTemp As String
    
    Dim sOldVoName As String
    Dim sOldValueName As String
    Dim sOldValue As String
    Dim sOldComment As String
    
    Dim iTypeCnt As Integer
    Dim iSapceCnt As Integer
    
    '���� �� ��ŭ �ݺ�
    For i = LBound(sSrcArr) To UBound(sSrcArr)
            
        sTemp = sSrcArr(i)
        sOrgTemp = sSrcArr(i)
        'sTemp = Replace(sTemp, "(", " ( ")
        'sTemp = Replace(sTemp, ")", " ( ")
        'sTemp = Replace(sTemp, "  ", " ")
        
        
        
        For j = 1 To fpVOList.MaxRows
        
            fpVOList.Row = j
            fpVOList.Col = 1: sOldVo = fpVOList.Value
            fpVOList.Col = 2: sOldItem = fpVOList.Value
            fpVOList.Col = 3: sNewVo = fpVOList.Value
            fpVOList.Col = 4: sNewItem = fpVOList.Value
            
            sOldVoName = ""
            sOldValueName = ""
            sOldValue = ""
            sOldComment = ""
            iTypeCnt = 0
            iSapceCnt = 0
            
            '���忡 VO��.set �ְ� new vo��� item���� ������ ġȯ
            If InStr(1, sTemp, sOldVo & ".set") > 0 And _
                    Trim(sNewVo) <> "" And _
                    Trim(sNewItem) <> "" Then
                

                '���¼� ������ �и��Ͽ� �迭�� �����Ѵ�.
                ReDim sWordArr(0)
                Call gfSplitNuclear(sTemp, sWordArr)
                
                
                '���¼ҷ� �и��� �ܾ ���������� �˻��ϸ� �ĺ��� ����.
                For K = LBound(sWordArr) To UBound(sWordArr)
                                                                                                                
                    If iTypeCnt = 0 And Trim(sWordArr(K)) = "" Then
                        iSapceCnt = iSapceCnt + 1
                    
                    'ù��° ������ �ܾ�� vo name
                    ElseIf iTypeCnt = 0 And Trim(sWordArr(K)) <> "" Then
                        sOldVoName = Trim(sWordArr(K))
                        iTypeCnt = 1
                        
                    'VO�� ������ ������ ���� ��ȣ
                    ElseIf iTypeCnt = 1 And Trim(sWordArr(K)) = "(" Then
                        iTypeCnt = 2
                    
                    ElseIf iTypeCnt = 2 And Trim(sWordArr(K)) = """" Then
                        iTypeCnt = 3
                    
                    '�Ӽ��� ã�Ҵ�.
                    ElseIf iTypeCnt = 3 And Trim(sWordArr(K)) <> "" Then
                        sOldValueName = Trim(sWordArr(K))
                        iTypeCnt = 4
                    
                    ElseIf iTypeCnt = 4 And Trim(sWordArr(K)) <> "," Then
                        iTypeCnt = 5
                                                                
                    'value ���� ��� �����Ѵ�.
                    ElseIf iTypeCnt = 5 And Trim(sWordArr(K)) <> "," Then
                        
                        sOldValue = sOldValue & Trim(sWordArr(K))
                        
                        If Trim(sWordArr(K)) = ";" Then
                            iTypeCnt = 6
                                    
                            '��ɾ �������Ƿ� value ������ ; �Ѱ��� �����Ѵ�.
                            
                            sOldValue = Trim(sOldValue)
                            sOldValue = Mid(sOldValue, 1, Len(sOldValue) - 1)
                            iTypeCnt = 7
                        End If
                        
                    
                    ElseIf iTypeCnt = 6 And Trim(sWordArr(K)) = ";" Then
                    
                        
                    
                    ElseIf iTypeCnt = 7 And (Trim(sWordArr(K)) <> "//" Or Trim(sWordArr(K)) <> "/*") Then
                        iTypeCnt = 8
                                                 
                    ElseIf iTypeCnt = 8 And Trim(sWordArr(K)) <> "" Then
                        sOldComment = sOldComment & Trim(sWordArr(K))
                    End If
                
                Next K
                
                '���ο� �������� �����Ѵ�. nevonam.set�Ӽ���(value)
                'Space ����
                sTemp = gfMakeString(" ", iSapceCnt / 2)
                sTemp = sTemp & Trim(sNewVo) & ".set" & Trim(sNewItem) & "(" & sOldValue & ");" & "    //" & sOldComment
                    
                
            End If
            
        
        Next j
        
        

        sLastSentence = sLastSentence & sTemp + vbCrLf
        
        
    Next i
    
    
    txtTgtSrc.Text = sLastSentence
    
    Exit Sub
    
Err_Handler:
    gfMsgInf (Err.Description)
    
End Sub

Private Sub Command1_Click()
    Clipboard.SetText txtTgtSrc.Text
End Sub



Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
    
    Me.Top = 0
    Me.Left = 0
    
    Me.Width = 26085
    Me.Height = 14175
    
    fpVariable.MaxRows = 0
    fpVO.MaxRows = 0
    
    
    txtType.Text = GetSetting("ORAIB_TAB_MAN_BY_HAK", "SORURCE_CONVERT", "VAR_TYPE", "ValueObject,ValueObjectAssembler,String,double,float,int")
    
    If txtType.Text = "" Then
        txtType.Text = "ValueObject,ValueObjectAssembler,String,double,float,int"
    End If
    
    SSTab1.Tab = 0
    
    'asis �ܾ��� ���� ������ ����
    Call setArray
End Sub

'------------------------------------------
'ASIS �ܾ��� ���� ������ ����
'------------------------------------------
Private Sub setArray()

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
    sSQL = sSQL & vbCrLf & "SELECT NVL(ASIS_ID,' ') AS ASIS_ID "
    sSQL = sSQL & vbCrLf & "      ,NVL(ASIS_NM,' ') AS ASIS_NM"
    sSQL = sSQL & vbCrLf & "      ,NVL(TOBE_NM,' ') AS TOBE_NM"
    sSQL = sSQL & vbCrLf & "      ,NVL(CNT,0)       AS CNT"
    sSQL = sSQL & vbCrLf & "  FROM HGDEV_T005"
    sSQL = sSQL & vbCrLf & " ORDER BY ASIS_ID, CNT DESC"
        
    rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    
    Dim i As Integer
    Dim iCnt As Integer
    
    rsTemp.MoveFirst
    iCnt = 0
    For i = 1 To rsTemp.RecordCount
        
        iCnt = iCnt + 1
        
        ReDim Preserve gsASIS_ID(iCnt)
        ReDim Preserve gsASIS_NM(iCnt)
        ReDim Preserve gsTOBE_NM(iCnt)
        ReDim Preserve gsCNT(iCnt)
        
        
    
        gsASIS_ID(iCnt) = Trim(rsTemp.Fields("ASIS_ID").Value)
        gsASIS_NM(iCnt) = Trim(rsTemp.Fields("ASIS_NM").Value)
        gsTOBE_NM(iCnt) = Trim(rsTemp.Fields("TOBE_NM").Value)
        gsCNT(iCnt) = Trim(rsTemp.Fields("CNT").Value)
        
        
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

'���� Ÿ�Կ� �ش� �ϴ��� Find
Private Function fFindVariableType(pVar) As Boolean
    
    Dim i As Integer
    
    For i = LBound(gsVarType) To UBound(gsVarType)
        If UCase(pVar) = UCase(gsVarType(i)) Then
            fFindVariableType = True
            Exit Function
        Else
            fFindVariableType = False
        End If
    Next i
    
End Function







Private Sub txtType_LostFocus()
    Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", "SORURCE_CONVERT", "VAR_TYPE", Trim(txtType.Text))
End Sub
