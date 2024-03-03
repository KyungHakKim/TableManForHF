VERSION 5.00
Begin VB.Form frmMetaConvert 
   BackColor       =   &H00F9F9D5&
   Caption         =   "용어생성기"
   ClientHeight    =   12480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20130
   Icon            =   "frmMetaConvert.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12480
   ScaleWidth      =   20130
   Begin VB.CommandButton Command1 
      Caption         =   "clipboard"
      Height          =   495
      Left            =   17700
      TabIndex        =   29
      Top             =   1800
      Width           =   2235
   End
   Begin VB.TextBox txtPostFix 
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
      Left            =   13620
      TabIndex        =   28
      Top             =   1860
      Width           =   2115
   End
   Begin VB.TextBox txtPrefix 
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
      Left            =   11340
      TabIndex        =   26
      Top             =   1860
      Width           =   2115
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   14820
      TabIndex        =   21
      Top             =   1080
      Width           =   4875
      Begin VB.OptionButton optUnderBar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "none"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton optUnderBar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Under bar"
         Height          =   255
         Index           =   0
         Left            =   1260
         TabIndex        =   22
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Under bar"
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
         TabIndex        =   24
         Top             =   270
         Width           =   945
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   915
      Left            =   14100
      TabIndex        =   16
      Top             =   120
      Width           =   6075
      Begin VB.OptionButton optUpperLower 
         BackColor       =   &H00FFFFFF&
         Caption         =   "단어 첫 대문자(첫 단어 소문자)"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2820
         TabIndex        =   25
         Top             =   600
         Value           =   -1  'True
         Width           =   3075
      End
      Begin VB.OptionButton optUpperLower 
         BackColor       =   &H00FFFFFF&
         Caption         =   "모두 대문자"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   900
         TabIndex        =   19
         Top             =   240
         Width           =   1395
      End
      Begin VB.OptionButton optUpperLower 
         BackColor       =   &H00FFFFFF&
         Caption         =   "모두 소문자"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2820
         TabIndex        =   18
         Top             =   240
         Width           =   1395
      End
      Begin VB.OptionButton optUpperLower 
         BackColor       =   &H00FFFFFF&
         Caption         =   "단어 첫 대문자"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   900
         TabIndex        =   17
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "대소문자"
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
         TabIndex        =   20
         Top             =   270
         Width           =   780
      End
   End
   Begin VB.TextBox txtConvertWord 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9915
      Left            =   10920
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   15
      Top             =   2340
      Width           =   9015
   End
   Begin VB.TextBox txtSplitWord 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9915
      Left            =   5400
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   14
      Top             =   2340
      Width           =   4575
   End
   Begin VB.TextBox txtOrgWord 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9855
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   13
      Top             =   2340
      Width           =   4395
   End
   Begin VB.CommandButton cmdDong 
      Caption         =   "동음이의사전"
      Enabled         =   0   'False
      Height          =   1035
      Left            =   10140
      TabIndex        =   12
      Top             =   6180
      Width           =   735
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   10080
      TabIndex        =   7
      Top             =   1080
      Width           =   4575
      Begin VB.OptionButton optNoneJs 
         BackColor       =   &H00FFFFFF&
         Caption         =   "none"
         Height          =   315
         Left            =   3540
         TabIndex        =   11
         Top             =   240
         Width           =   795
      End
      Begin VB.OptionButton optJS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Java Script"
         Height          =   315
         Left            =   2040
         TabIndex        =   10
         Top             =   240
         Width           =   1275
      End
      Begin VB.OptionButton optJava 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Java"
         Height          =   315
         Left            =   1020
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Type"
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
         TabIndex        =   8
         Top             =   270
         Width           =   465
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   915
      Left            =   10080
      TabIndex        =   2
      Top             =   120
      Width           =   3975
      Begin VB.OptionButton optComment 
         BackColor       =   &H00FFFFFF&
         Caption         =   "none"
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   6
         Top             =   240
         Width           =   1035
      End
      Begin VB.OptionButton optComment 
         BackColor       =   &H00FFFFFF&
         Caption         =   "/* */"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optComment 
         BackColor       =   &H00FFFFFF&
         Caption         =   "//"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "주석 "
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
         TabIndex        =   3
         Top             =   270
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdStep2 
      Caption         =   "☞"
      Height          =   2175
      Left            =   10080
      TabIndex        =   1
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdStep1 
      Caption         =   "☞"
      Height          =   2175
      Left            =   4560
      TabIndex        =   0
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '투명
      Caption         =   "접두어/접미어"
      Height          =   255
      Left            =   10020
      TabIndex        =   27
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "frmMetaConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdDong_Click()
    frmDong.Show vbModal
End Sub

'----------------------------------------
'1단계 분리
'----------------------------------------
Private Sub cmdStep1_Click()
    Dim sArr() As String
    Dim i As Integer
    Dim sResult As String
    
    
    txtSplitWord.Text = ""
    
    sArr = Split(txtOrgWord.Text, vbCrLf)
    
    For i = LBound(sArr) To UBound(sArr)
        '분리하는 함수 호출
        sResult = gfSplitWord(sArr(i))
            
        txtSplitWord.Text = txtSplitWord.Text & sResult & vbCrLf
            
    Next i
    
    If InStr(1, txtSplitWord.Text, "동음이의") > 0 Then
        cmdDong.Enabled = True
    End If
End Sub

'----------------------------------------
'2단계 분리
'----------------------------------------
Private Sub cmdStep2_Click()
    Dim sArr() As String
    Dim i, j As Integer
    Dim sResult As String
    Dim sComment As String
    Dim sWordList() As String
    
    txtConvertWord.Text = ""
    
    '변환 대상 문자를 Line 단위로 배열 설정
    sArr = Split(txtSplitWord.Text, vbCrLf)
    
    '주석 형태
    If optComment(0).Value = True Then
        sComment = "/"
    ElseIf optComment(1).Value = True Then
        sComment = "*"
    Else
        sComment = ""
    End If
            
            
    
            
    '변환 대상만큼 반복
    For i = LBound(sArr) To UBound(sArr)
        
        If Trim(sArr(i)) <> "" Then
        
            sResult = gfMappingEng(sArr(i), sComment, _
                                   IIf(optUpperLower(0).Value = True, "U", IIf(optUpperLower(1).Value = True, "L", IIf(optUpperLower(2).Value = True, "F", "l"))), _
                                   IIf(optUnderBar(0).Value = True, True, False))
                
                
            '첫단어만 소문자 나머지 대문자인 경우
            If optUpperLower(3).Value = True Then
                sResult = LCase(Left(sResult, 1)) & Mid(sResult, 2)
            End If
            
            
            sWordList = Split(sArr(i), "_")
                
            
            'Type 정의
            If Trim(sArr(i)) <> "" Then
            
                'Java Script
                If optJS.Value = True Then
                    sResult = "var " & sResult
                
                'None
                ElseIf optNoneJs.Value = True Then
                    sResult = sResult
                
                'Java
                Else
                    '도메인 찾기
                    For j = LBound(gsDomainName) To UBound(gsDomainName)
                                    
                        '도메인타입에 따른 변수타입 지정
                        If gsDomainName(j) = sWordList(UBound(sWordList)) Then
                        
                            Select Case UCase(gsDomianType(j))
                                
                                Case "NUMERIC"
                                    sResult = "BigDecimal " & sResult
                                    Exit For
                                    
                                Case "VARCHAR", "CLOB", "TIMESTAMP", "BLOB", "DATE"
                                    sResult = "String " & sResult
                                    Exit For
                                    
                                    
                            End Select
                            
                        End If
                            
                    Next j
                End If
            End If
        
            If Trim(txtPrefix.Text) <> "" Then
                sResult = Trim(txtPrefix.Text) & sResult
            End If
            
            If Trim(txtPostFix.Text) <> "" Then
                sResult = sResult & Trim(txtPostFix.Text)
            End If
            
            sResult = Trim(sResult)
            txtConvertWord.Text = txtConvertWord.Text & sResult & vbCrLf
        End If
                
    Next i
    
    Dim iMaxLen As Integer
    

    
    
    Dim sEngArr() As String
    Dim sMadeWord As String
    '----------------------------------
    ' 주석 설정 하기
    '----------------------------------
    sEngArr = Split(txtConvertWord.Text, vbCrLf)
    
    '최대 용어 사이즈 계산
    For i = LBound(sEngArr) To UBound(sEngArr)
        
        If Len(Trim(sEngArr(i))) > iMaxLen Then
            iMaxLen = Len(Trim(sEngArr(i)))
        End If
        
    Next i
    
    
    txtConvertWord.Text = ""
    For i = LBound(sEngArr) To UBound(sEngArr)
            
        sMadeWord = ""
        If Trim(sEngArr(i)) <> "" Then
            'java 형태가 아니면 ; 붙이지 않음
            If optNoneJs.Value = True Then
                sMadeWord = sEngArr(i) & "" & gfMakeString(" ", iMaxLen - Len(sEngArr(i)) + 10)  '10자리 공백띄기
            Else
                sMadeWord = sEngArr(i) & ";" & gfMakeString(" ", iMaxLen - Len(sEngArr(i)) + 10)  '10자리 공백띄기
            End If
            
        End If
        

        
        If sEngArr(i) <> "" Then
            
            sArr(i) = Replace(sArr(i), "(동음이의어)", "")
                    
            If sComment = "/" Then
                sMadeWord = sMadeWord & "//" & Replace(sArr(i), "_", "")
            ElseIf sComment = "*" Then
                sMadeWord = sMadeWord & "/* " & Replace(sArr(i), "_", "") & " */"
            Else
                sMadeWord = sMadeWord
            End If
        End If
        
        
        sMadeWord = Trim(sMadeWord)
        
        txtConvertWord.Text = txtConvertWord.Text & sMadeWord & vbCrLf
    

    
    Next i
    
    
End Sub


Private Sub Command1_Click()
        Clipboard.Clear
        Clipboard.SetText txtConvertWord.Text
End Sub

'----------------------------------
' Form Load 이벤트
'----------------------------------
Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    
    Me.Height = 13065
    Me.Width = 20370


    '단어를 전역 배열에 셋팅
    Call subSetMetaWord
    
    
    'test 용어 셋팅
'    txtOrgWord.Text = "개인주택금융공사대출금액" & vbCrLf & _
'                                "개인주택금액" & vbCrLf & _
'                                "개인주택금융공사" & vbCrLf & _
'                                "대출금액"
End Sub



Private Sub optComment_Click(Index As Integer)
    cmdStep2_Click
End Sub

Private Sub optJava_Click()
    cmdStep2_Click
End Sub

Private Sub optJS_Click()
    cmdStep2_Click
End Sub

Private Sub optNoneJs_Click()
    cmdStep2_Click
End Sub

Private Sub optUnderBar_Click(Index As Integer)
    cmdStep2_Click
End Sub

Private Sub optUpperLower_Click(Index As Integer)
    cmdStep2_Click
End Sub

Private Sub txtPostFix_Change()
    cmdStep2_Click
End Sub

Private Sub txtPrefix_Change()
    cmdStep2_Click
End Sub
