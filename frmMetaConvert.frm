VERSION 5.00
Begin VB.Form frmMetaConvert 
   BackColor       =   &H00F9F9D5&
   Caption         =   "��������"
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
      IMEMode         =   8  '����
      Left            =   13620
      TabIndex        =   28
      Top             =   1860
      Width           =   2115
   End
   Begin VB.TextBox txtPrefix 
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
      IMEMode         =   8  '����
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
         BackStyle       =   0  '����
         Caption         =   "Under bar"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "�ܾ� ù �빮��(ù �ܾ� �ҹ���)"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "��� �빮��"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "��� �ҹ���"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "�ܾ� ù �빮��"
         BeginProperty Font 
            Name            =   "����"
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
         BackStyle       =   0  '����
         Caption         =   "��ҹ���"
         BeginProperty Font 
            Name            =   "����"
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
      ScrollBars      =   3  '�����
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
      ScrollBars      =   3  '�����
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
      ScrollBars      =   3  '�����
      TabIndex        =   13
      Top             =   2340
      Width           =   4395
   End
   Begin VB.CommandButton cmdDong 
      Caption         =   "�������ǻ���"
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
         BackStyle       =   0  '����
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "����"
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
         BackStyle       =   0  '����
         Caption         =   "�ּ� "
         BeginProperty Font 
            Name            =   "����"
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
      Caption         =   "��"
      Height          =   2175
      Left            =   10080
      TabIndex        =   1
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdStep1 
      Caption         =   "��"
      Height          =   2175
      Left            =   4560
      TabIndex        =   0
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '����
      Caption         =   "���ξ�/���̾�"
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
'1�ܰ� �и�
'----------------------------------------
Private Sub cmdStep1_Click()
    Dim sArr() As String
    Dim i As Integer
    Dim sResult As String
    
    
    txtSplitWord.Text = ""
    
    sArr = Split(txtOrgWord.Text, vbCrLf)
    
    For i = LBound(sArr) To UBound(sArr)
        '�и��ϴ� �Լ� ȣ��
        sResult = gfSplitWord(sArr(i))
            
        txtSplitWord.Text = txtSplitWord.Text & sResult & vbCrLf
            
    Next i
    
    If InStr(1, txtSplitWord.Text, "��������") > 0 Then
        cmdDong.Enabled = True
    End If
End Sub

'----------------------------------------
'2�ܰ� �и�
'----------------------------------------
Private Sub cmdStep2_Click()
    Dim sArr() As String
    Dim i, j As Integer
    Dim sResult As String
    Dim sComment As String
    Dim sWordList() As String
    
    txtConvertWord.Text = ""
    
    '��ȯ ��� ���ڸ� Line ������ �迭 ����
    sArr = Split(txtSplitWord.Text, vbCrLf)
    
    '�ּ� ����
    If optComment(0).Value = True Then
        sComment = "/"
    ElseIf optComment(1).Value = True Then
        sComment = "*"
    Else
        sComment = ""
    End If
            
            
    
            
    '��ȯ ���ŭ �ݺ�
    For i = LBound(sArr) To UBound(sArr)
        
        If Trim(sArr(i)) <> "" Then
        
            sResult = gfMappingEng(sArr(i), sComment, _
                                   IIf(optUpperLower(0).Value = True, "U", IIf(optUpperLower(1).Value = True, "L", IIf(optUpperLower(2).Value = True, "F", "l"))), _
                                   IIf(optUnderBar(0).Value = True, True, False))
                
                
            'ù�ܾ �ҹ��� ������ �빮���� ���
            If optUpperLower(3).Value = True Then
                sResult = LCase(Left(sResult, 1)) & Mid(sResult, 2)
            End If
            
            
            sWordList = Split(sArr(i), "_")
                
            
            'Type ����
            If Trim(sArr(i)) <> "" Then
            
                'Java Script
                If optJS.Value = True Then
                    sResult = "var " & sResult
                
                'None
                ElseIf optNoneJs.Value = True Then
                    sResult = sResult
                
                'Java
                Else
                    '������ ã��
                    For j = LBound(gsDomainName) To UBound(gsDomainName)
                                    
                        '������Ÿ�Կ� ���� ����Ÿ�� ����
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
    ' �ּ� ���� �ϱ�
    '----------------------------------
    sEngArr = Split(txtConvertWord.Text, vbCrLf)
    
    '�ִ� ��� ������ ���
    For i = LBound(sEngArr) To UBound(sEngArr)
        
        If Len(Trim(sEngArr(i))) > iMaxLen Then
            iMaxLen = Len(Trim(sEngArr(i)))
        End If
        
    Next i
    
    
    txtConvertWord.Text = ""
    For i = LBound(sEngArr) To UBound(sEngArr)
            
        sMadeWord = ""
        If Trim(sEngArr(i)) <> "" Then
            'java ���°� �ƴϸ� ; ������ ����
            If optNoneJs.Value = True Then
                sMadeWord = sEngArr(i) & "" & gfMakeString(" ", iMaxLen - Len(sEngArr(i)) + 10)  '10�ڸ� ������
            Else
                sMadeWord = sEngArr(i) & ";" & gfMakeString(" ", iMaxLen - Len(sEngArr(i)) + 10)  '10�ڸ� ������
            End If
            
        End If
        

        
        If sEngArr(i) <> "" Then
            
            sArr(i) = Replace(sArr(i), "(�������Ǿ�)", "")
                    
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
' Form Load �̺�Ʈ
'----------------------------------
Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    
    Me.Height = 13065
    Me.Width = 20370


    '�ܾ ���� �迭�� ����
    Call subSetMetaWord
    
    
    'test ��� ����
'    txtOrgWord.Text = "�������ñ����������ݾ�" & vbCrLf & _
'                                "�������ñݾ�" & vbCrLf & _
'                                "�������ñ�������" & vbCrLf & _
'                                "����ݾ�"
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
