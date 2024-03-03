VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmReport 
   BackColor       =   &H00EAF4D9&
   Caption         =   "OZ Report Convert"
   ClientHeight    =   11670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21765
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11670
   ScaleWidth      =   21765
   Begin VB.TextBox txtPathName 
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
      Left            =   16800
      TabIndex        =   9
      Top             =   180
      Width           =   5115
   End
   Begin VB.TextBox txtFileName 
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
      Left            =   11580
      TabIndex        =   8
      Top             =   180
      Width           =   5115
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "저장"
      Height          =   495
      Left            =   19680
      TabIndex        =   7
      Top             =   420
      Width           =   1995
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10635
      Left            =   6660
      TabIndex        =   4
      Top             =   900
      Width           =   14955
      _ExtentX        =   26379
      _ExtentY        =   18759
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "원본"
      TabPicture(0)   =   "frmReport.frx":048A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtOrg"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "변환"
      TabPicture(1)   =   "frmReport.frx":04A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtTgt"
      Tab(1).ControlCount=   1
      Begin RichTextLib.RichTextBox txtOrg 
         Height          =   10095
         Left            =   120
         TabIndex        =   5
         Top             =   420
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   17806
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmReport.frx":04C2
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
      Begin RichTextLib.RichTextBox txtTgt 
         Height          =   10095
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   17806
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmReport.frx":0567
      End
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
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   9315
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Scan"
      Height          =   495
      Left            =   9420
      TabIndex        =   0
      Top             =   0
      Width           =   1995
   End
   Begin FPSpreadADO.fpSpread fpFileList 
      Height          =   10635
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   6375
      _Version        =   393216
      _ExtentX        =   11245
      _ExtentY        =   18759
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
      SpreadDesigner  =   "frmReport.frx":0604
      ScrollBarTrack  =   1
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   540
      Width           =   16515
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
    Call SaveSetting("ORAIB_TAB_MAN_BY_HAK", "SORURCE_CONVERT", "REPORT_FOLDER_NAME", Trim(txtFolder.Text))
        
        
        
    Screen.MousePointer = vbHourglass
    
    
    
    '파일명을 검색한다.
    Call GetMetaInfo(txtFolder.Text, Me.lblStatus)
    
    lblStatus.Caption = ""
    
    fpFileList.MaxRows = 0
    
    Dim sFilename() As String
    Dim sFileDir() As String
    

    
    For i = LBound(gBas_FileDir) + 1 To UBound(gBas_FileDir)
    
        
        If LCase(Right(gBas_FileName(i), 4)) = ".ozr" Then
            fpFileList.MaxRows = fpFileList.MaxRows + 1
            fpFileList.Row = i + 1
            fpFileList.Col = 1
            fpFileList.Value = gBas_FileName(i)
            
            fpFileList.Col = 2
            fpFileList.Value = gBas_FileDir(i)
            
        End If
    
    Next i
    
    
    
    fpFileList.Row = -1
    fpFileList.Col = -1
    fpFileList.Lock = True
    
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
Err_Handler:
    Screen.MousePointer = vbDefault
    gfMsgWrn (Err.Description)
End Sub

'파일을 저장한다.
Private Sub cmdSave_Click()
    Dim stmFileControl As Stream
    Dim sWriteData As String
    
    
           Set stmFileControl = New Stream

            stmFileControl.Open
            stmFileControl.Position = 0
            stmFileControl.Charset = "UTF-8"
            


            stmFileControl.WriteText (txtOrg.Text)
            
            stmFileControl.SaveToFile Trim(txtPathName.Text) & "\" & "변경" & txtFileName.Text, adSaveCreateOverWrite
            stmFileControl.Close
            Set stmFileControl = Nothing


            '파일 write
            'stmFileControl.WriteText (sWriteData & vbCrLf)

    
    
    
    
'    Do
'
'           Set stmFileControl = New Stream
'
'            stmFileControl.Open
'            stmFileControl.Position = 0
'            stmFileControl.Charset = "UTF-8"
'            sOldFileName = sFilename
'
'            '파일 write
'            'stmFileControl.WriteText (sWriteData & vbCrLf)
'
'            '파일 write
'            stmFileControl.WriteText (sWriteData & vbCrLf)
'        End If
'    Loop
'
    
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.Height = 12255
    Me.Width = 22005
    
        
    txtFolder.Text = GetSetting("ORAIB_TAB_MAN_BY_HAK", "SORURCE_CONVERT", "REPORT_FOLDER_NAME", "D:\TEMP")
    
    SSTab1.Tab = 0
    
    
End Sub



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
    Dim sArr() As String
    Dim i As Integer
    
    fpFileList.Row = Row
    fpFileList.Col = 1
    sFilename = fpFileList.Value
    
    fpFileList.Row = Row
    fpFileList.Col = 2
    sPathName = fpFileList.Value
        
    
    txtFileName.Text = sFilename
    txtPathName.Text = sPathName
    
    sTemp = TextStreamRead(sPathName & "\" & sFilename)
    sArr = Split(sTemp, Chr(10))
    
    
    Dim sWriteData As String
    sWriteData = ""
    
    Dim sChangeStr As String
    Dim bTableStart As Boolean
    
    Dim sProperty() As String
    
    Dim sTableLEFT, sTableTOP, sTableWIDTH, sTableHEGIHT As String
    
    bTableStart = False
    '문장 단위로 for 문 돌며 xml을 변경한다.
    For i = LBound(sArr) To UBound(sArr)
        sChangeStr = sArr(i)
        
        ' 빈문자열은 그냥 write 한다.
        If Trim(sChangeStr) = "" Then GoTo NEXT_FOR
                    
                            
        '닫는 태그 뒤에 문자가 있을 수 있으므로 구분하기 위하여 스페이스를 한칸 띄운다.
        sChangeStr = fTagRenewal(sChangeStr)
                    
        'xml 구문을 속성별로 분류하여 배열에 저장한다.
        sProperty = Split(sChangeStr, " ")
                                        
        '마지막 배열에는 닫는 >, /> 태그가 있으므로 삭제한다.
        'sProperty(UBound(sProperty)) = Replace(sProperty(UBound(sProperty)), "/>", "")
        'sProperty(UBound(sProperty)) = Replace(sProperty(UBound(sProperty)), ">", "")
        
         
        'Table 끝 구문이면
        If InStr(1, sChangeStr, "</OZTABLESTATIC>") > 0 Then
            bTableStart = False
        End If
        
        
                        
        'Table 시작 구문이 시작 됐고 oztablestat 태그가 있으면
        If bTableStart And InStr(1, sChangeStr, "<OZTABLELABEL") > 1 Then
            
            '0.5, 0.125
            
            '위에 라인에 해당되는 지 찾는다. 위에 라인은 TOP과 같은T
            If getProperty(sProperty, "TOP") = sTableTOP Then
                setProperty sProperty, "DRAWTOP", "0.5"
            End If
                        
            '좌측 라인에 해당되는 지 찾는다. 좌측 라인은 LEFT이 0인 건
            If getProperty(sProperty, "LEFT") = sTableLEFT Then
                setProperty sProperty, "DRAWLEFT", "0.5"
            End If
                        
                                                
            '하단 라인에 해당되는 지 찾는다. 하단 라인은 TOP이 테이블 전체의 HEIGHT - 행의 HEIGHT = 행의 TOP
            If CStr(CDbl(Trim(sTableTOP)) + CDbl(Trim(sTableHEGIHT)) - CDbl(getProperty(sProperty, "HEIGHT"))) = CStr(CDbl(Trim(getProperty(sProperty, "TOP")))) Then
                setProperty sProperty, "DRAWBOTTOM", "0.5"
            End If
            
            '우측 라인에 해당되는 지 찾는다. 우측 라인은 LEFT가 테이블 전체의 WIDTH - 행의 WIDTH = 행의 TOP
            If CStr(CDbl(sTableLEFT) + CDbl(sTableWIDTH) - CDbl(getProperty(sProperty, "WIDTH"))) = CStr(CDbl(getProperty(sProperty, "LEFT"))) Then
                setProperty sProperty, "DRAWRIGHT", "0.5"
            End If
            
            
            'Porperty 배열을 String으로 변환 한다.
            sChangeStr = Join(sProperty, " ")
        End If
                    
                    
        'Table 시작 구문이면
        If InStr(1, sChangeStr, "<OZTABLESTATIC") > 0 Then
            bTableStart = True
            
            '테이블의 위치를 찾는다
            sTableLEFT = getProperty(sProperty, "LEFT")
            sTableTOP = getProperty(sProperty, "TOP")
            sTableWIDTH = getProperty(sProperty, "WIDTH")
            sTableHEGIHT = getProperty(sProperty, "HEIGHT")
            
        End If
            
            
NEXT_FOR:
            
        sWriteData = sWriteData & sChangeStr & Chr(13) & Chr(10)
        


    Next i
    
    
    Dim objStreamUTF8 As Stream
    Set objStreamUTF8 = New Stream
    
    
    Dim objStreamUTF8NoBOM As Stream
    Set objStreamUTF8NoBOM = New Stream
    
    
    With objStreamUTF8
        .Charset = "UTF-8"
        .Open
        .WriteText sWriteData
        .Position = 0
        '.SaveToFile Trim(txtPathName.Text) & "\" & "UTF8" & txtFileName.Text, adSaveCreateOverWrite
        .Type = adTypeText
        .Position = 3
    End With
    
    With objStreamUTF8NoBOM
        .Type = adTypeBinary
        .Open
        objStreamUTF8.CopyTo objStreamUTF8NoBOM
        .SaveToFile Trim(txtPathName.Text) & "\" & "변경-" & txtFileName.Text, adSaveCreateOverWrite
        
    End With
    
    objStreamUTF8.Close
    objStreamUTF8NoBOM.Close
    
    Set objStreamUTF8 = Nothing
    Set objStreamUTF8NoBOM = Nothing
    

    
    Exit Sub
Err_Handler:
    If Err.Number = 53 Then
        gfMsgInf " 파일을 찾을 수 없습니다."
    Else
        gfMsgInf (Err.Description & " " & Err.Number)
    End If
End Sub

'------------------------------------------------------------
'Property의 속성 값을 구한다.
'------------------------------------------------------------
Function getProperty(pProperty() As String, pName As String)
    Dim i As Integer
    Dim sPare() As String
    Dim sRet As String
    
    sRet = "NOT_FOUND"
    
    For i = LBound(pProperty) To UBound(pProperty)
    
        '태그 구분자가 아니고 속성이면
        If InStr(1, pProperty(i), "=") Then
            
            sPare = Split(pProperty(i), "=")
            
            sPare(0) = Trim(sPare(0))
            If UCase(Trim(sPare(0))) = UCase(pName) Then
                sRet = Trim(Replace(sPare(1), """", ""))
                sRet = Trim(Replace(sRet, "/>", ""))
                sRet = Trim(Replace(sRet, ">", ""))
                Exit For
            End If
        End If
    Next i
    
    getProperty = sRet
    
    
End Function

'------------------------------------------------------------
'Property의 속성 값을 변경한다.
'------------------------------------------------------------
Sub setProperty(pProperty() As String, pName As String, pValue As String)
    Dim i As Integer
    Dim sPare() As String
    Dim sRet As String
    
    Dim bCloseTagSalsh2 As Boolean
    Dim bCloseTagSalsh  As Boolean
    
    sRet = "NOT_FOUND"
    bCloseTagSalsh2 = False
    bCloseTagSalsh = False
    For i = LBound(pProperty) To UBound(pProperty)
        '태그 구분자가 아니고 속성이면
        If InStr(1, pProperty(i), "=") Then
            
            sPare = Split(pProperty(i), "=")
            
            sPare(0) = Trim(sPare(0))
            
            If InStr(1, pProperty(i), "/>") > 0 Then
                bCloseTagSalsh2 = True
                pProperty(i) = Replace(pProperty(i), "/>", "")
                
            ElseIf InStr(1, pProperty(i), ">") > 0 Then
                bCloseTagSalsh = True
                pProperty(i) = Replace(pProperty(i), ">", "")
            End If
            
            
            
            If UCase(Trim(sPare(0))) = UCase(pName) Then
            
                pProperty(i) = sPare(0) & "=" & """" & pValue & """"
                
            End If
            
            If bCloseTagSalsh2 Then
                pProperty(i) = pProperty(i) & "/>"
                bCloseTagSalsh2 = False
            ElseIf bCloseTagSalsh Then
                pProperty(i) = pProperty(i) & ">"
                bCloseTagSalsh = False
            End If
            
        End If
    Next i
    
    '해당 속성이 존재하지 않으면 추가
    If sRet = "NOT_FOUND" Then
        For i = LBound(pProperty) To UBound(pProperty)
        
            If InStr(1, pProperty(i), "/>") > 0 Then
 
                pProperty(i) = pName & "=" & """" & pValue & """" & " " & pProperty(i)
                Exit For
            ElseIf InStr(1, pProperty(i), ">") > 0 Then

                pProperty(i) = pName & "=" & """" & pValue & """" & " " & pProperty(i)
                Exit For
                
            End If
        
        Next i
    End If
    
    
End Sub


'닫는 태그 뒤에 문자가 있을 수 있으므로 구분하기 위하여 스페이스를 한칸 띄운다.
Private Function fTagRenewal(pStr) As String
    Dim iCnt As Integer
    Dim sRet As String
    Dim sToken As String
    Dim i As Integer
    For i = 1 To Len(pStr)
            
        sToken = Mid(pStr, i, 1)
        
        '시작되는 < 태그가 아니면
        
        If sToken = "<" And iCnt <> 1 Then
            sToken = " <"
        ElseIf sToken = ">" Then
            sToken = "> "
        End If
        sRet = sRet & sToken
    Next i
    
    
    fTagRenewal = sRet
End Function



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
