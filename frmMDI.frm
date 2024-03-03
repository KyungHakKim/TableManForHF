VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H00FFFFFF&
   Caption         =   "HGDevTool Ver 0.3"
   ClientHeight    =   13455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22650
   Icon            =   "frmMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows �⺻��
   WindowState     =   2  '�ִ�ȭ
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  '�� ����
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   22650
      _ExtentX        =   39952
      _ExtentY        =   1164
      ButtonWidth     =   2646
      ButtonHeight    =   1005
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   9
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "��������"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Source Converter"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Rule Converter"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Vlookup"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "VO manager"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "OZ Report"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "ASIS mapping"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "UI manager"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "���κ���"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      Begin VB.Timer Timer1 
         Interval        =   60000
         Left            =   21000
         Top             =   120
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private sForm(100) As Boolean



Private Sub MDIForm_Load()
On Error GoTo Err_Handler
    Me.Top = 100
    Me.Left = 100
    

    Dim i As Integer
    For i = 1 To 100
        sForm(i) = False
    Next i
    
    
    
    Dim sUser() As String
    
    '����� ���� üũ
    If Not gfReadFile(App.Path & "\" & "�����.txt", ";", sUser) Then
       Unload Me
    End If
        
    gUserId = Trim(sUser(0, 0))
    gUserNm = Trim(sUser(0, 1))
    
    If Trim(gUserId) = "" Or Trim(gUserNm) = "" Then
        gfMsgWrn ("����� ������ ����� ���Ͽ� ��ϵ��� �ʾҽ��ϴ�.")
        Unload Me
        End
    ElseIf Left(gUserId, 2) <> "ns" And gUserId <> "admin" Then
        gfMsgWrn ("����� ������ ����� ���Ͽ� �߸� ��� �Ǿ����ϴ�. nsXXX ���̵� �Է��ϼ���")
        Unload Me
        End
    ElseIf gUserId = "ns041" Then
        gfMsgWrn ("����� ns041 ������ ���� �������� ���� �ϼ���.")
        Unload Me
        End
    End If
    
    
    Dim sJuSinId() As String
    sJuSinId = Split("admin;ns042;ns045;ns047;ns043;ns044;ns046;ns048;ns049;ns050;ns051;ns084;ns111;ns112;ns136;ns097;ns100;ns144;ns146", ";")
    
    For i = LBound(sJuSinId) To UBound(sJuSinId)
        
        If sJuSinId(i) = gUserId Then
            gJuSinBo = True
            Exit For
        Else
            gJuSinBo = False
        End If
        
    Next i
    
    
    '�޴� ���� ����
    
    If gJuSinBo Then
        Toolbar1.Buttons(1).Visible = True  '��������
        Toolbar1.Buttons(2).Visible = True '�ҽ� ��ȯ��
        Toolbar1.Buttons(3).Visible = True 'rule ��ȯ��
        Toolbar1.Buttons(4).Visible = True 'vlookup
        Toolbar1.Buttons(5).Visible = True 'VO Manager
        Toolbar1.Buttons(6).Visible = False 'Report
        Toolbar1.Buttons(7).Visible = True  'mapping
        Toolbar1.Buttons(8).Visible = True  'ui manager
        Toolbar1.Buttons(9).Visible = True  '�ҽ� �ϰ� ����
        
    Else
        Toolbar1.Buttons(1).Visible = True  '��������
        Toolbar1.Buttons(2).Visible = False '�ҽ� ��ȯ��
        Toolbar1.Buttons(3).Visible = True 'rule ��ȯ��
        Toolbar1.Buttons(4).Visible = False 'vlookup
        Toolbar1.Buttons(5).Visible = True 'VO Manager
        Toolbar1.Buttons(6).Visible = False 'Report
        Toolbar1.Buttons(7).Visible = True  'mapping
        Toolbar1.Buttons(8).Visible = False  'ui manager
        Toolbar1.Buttons(9).Visible = False '�ҽ� �ϰ� ����
    End If
    
    If gUserId = "admin" Then
        Toolbar1.Buttons(4).Visible = True 'vlookup
        Toolbar1.Buttons(2).Visible = True '�ҽ� ��ȯ��
        Toolbar1.Buttons(5).Visible = True 'VO Manager
        Toolbar1.Buttons(6).Visible = True 'Report
        Toolbar1.Buttons(7).Visible = True  'mapping
        Toolbar1.Buttons(8).Visible = True  'ui manager
        Toolbar1.Buttons(9).Visible = True  '�ҽ� �ϰ� ����
    End If
    
    
    
    Exit Sub
Err_Handler:
    gfMsgWrn (Err.Description)
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    'Screen.MousePointer = vbHourglass
    

    Set conn = New ADODB.Connection
    conn.ConnectionString = GetConnStringHgDev

    conn.Open

    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
        sSQL = ""
        sSQL = sSQL & vbCrLf & "SELECT  count(*)  AS CNT          "
        sSQL = sSQL & vbCrLf & "  FROM HGDEV_T003           "
        sSQL = sSQL & vbCrLf & " WHERE 1=1"
        sSQL = sSQL & vbCrLf & "   AND ( (  ������  = '" & Trim(gUserNm) & "'  and ���� = '���'  )      "
        sSQL = sSQL & vbCrLf & "      OR (  �����  = '" & Trim(gUserNm) & "'  and ���� = '��ġ�Ϸ�'  )      "
        sSQL = sSQL & vbCrLf & "       ) "
        sSQL = sSQL & vbCrLf & "   AND ������ <> '�����' "
        sSQL = sSQL & vbCrLf & " AND NVL(Ȯ�ο���, 'N') <> 'Y' "

    rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

    
    rsTemp.MoveFirst
    If rsTemp.Fields("CNT").Value > 0 Then
        frmPopUP.Show
    End If
        
    

    Set conn = Nothing
    Set rsTemp = Nothing
    

    Exit Sub
ErrorHandler:
    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
     
     
     
     
     Select Case Button.Index
        Case 1
            If sForm(Button.Index) = False Then
                '��������
                frmMetaConvert.Show
            Else
                frmMetaConvert.SetFocus
            End If
        Case 2
            If sForm(Button.Index) = False Then
                'src converter
                frmSrcConvert.Show
            Else
                frmSrcConvert.SetFocus
            End If
        Case 3  '�꺯ȯ��
            If sForm(Button.Index) = False Then
                frmRule.Show
            Else
                frmRule.SetFocus
            End If
            
        Case 4  'vlookup
            If sForm(Button.Index) = False Then
                frmVlookup.Show
            Else
                frmVlookup.SetFocus
            End If
            
        Case 5  'VO Manager
            If sForm(Button.Index) = False Then
                frmVO.Show
            Else
                frmVO.SetFocus
            End If
            
        Case 6  'OZ Report
            If sForm(Button.Index) = False Then
                frmReport.Show
            Else
                frmReport.SetFocus
            End If
            
        Case 7  'AS-IS MAPPING
            If sForm(Button.Index) = False Then
                frmMethod.Show
            Else
                frmMethod.SetFocus
            End If
            
        Case 8  'UI MANAGER
            If sForm(Button.Index) = False Then
                frmUI.Show
            Else
                frmUI.SetFocus
            End If
            
        Case 9  '���κ���
            If sForm(Button.Index) = False Then
                frmMakeSrc.Show
            Else
                frmMakeSrc.SetFocus
            End If
            
            
            
            
     End Select
     
     sForm(Button.Index) = True
     
End Sub
