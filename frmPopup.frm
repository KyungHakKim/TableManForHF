VERSION 5.00
Begin VB.Form frmPopUP 
   Caption         =   "�˸�"
   ClientHeight    =   4035
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11115
   Icon            =   "frmPopup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   11115
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdSearch2 
      BackColor       =   &H008080FF&
      Caption         =   "Ȯ��"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   3240
      TabIndex        =   1
      Top             =   2460
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "�ű� ������ ��� �Ǿ����ϴ�."
      BeginProperty Font 
         Name            =   "������� ExtraBold"
         Size            =   26.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2340
      TabIndex        =   0
      Top             =   1440
      Width           =   6855
   End
End
Attribute VB_Name = "frmPopUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSearch2_Click()

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
    

    sSQL = ""
    sSQL = sSQL & vbCrLf & " UPDATE HGDEV_T003 "
    sSQL = sSQL & vbCrLf & "    SET Ȯ�ο��� = 'Y'"
    sSQL = sSQL & vbCrLf & " WHERE 1=1 "
    sSQL = sSQL & vbCrLf & "   AND ( (  ������  = '" & Trim(gUserNm) & "'  and ���� = '���'  )      "
    sSQL = sSQL & vbCrLf & "      OR (  �����  = '" & Trim(gUserNm) & "'  and ���� = '��ġ�Ϸ�'  )      "
    sSQL = sSQL & vbCrLf & "       ) "

    sSQL = sSQL & vbCrLf & " AND NVL(Ȯ�ο���, 'N') <> 'Y' "
    
    
    conn.Execute sSQL


    'commit
    conn.CommitTrans
    
    
    
    frmMethod.Show
    frmMethod.SetFocus
    frmMethod.SSTab1.Tab = 1
    
    
   
    '����ȸ
    Set conn = Nothing
    Set rsTemp = Nothing
    
    Unload Me
    
    Exit Sub
ErrorHandler:
    
    conn.RollbackTrans
    
    
    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description
    
    
    




End Sub
