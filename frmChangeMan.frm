VERSION 5.00
Begin VB.Form frmChangeMan 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Change Man"
   ClientHeight    =   1980
   ClientLeft      =   31215
   ClientTop       =   5760
   ClientWidth     =   5430
   Icon            =   "frmChangeMan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1980
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "�ٿ�ε�� ������ �����ϰ� �ֽ��ϴ�. ��� �� ���α׷��� �ٽ� �⵿�˴ϴ�."
      BeginProperty Font 
         Name            =   "����ü"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   900
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   5460
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�ٿ�ε� ���� ������...."
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   3330
   End
   Begin VB.Image imgLoginMain 
      Height          =   1965
      Left            =   0
      MousePointer    =   4  '������
      Picture         =   "frmChangeMan.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5460
   End
End
Attribute VB_Name = "frmChangeMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fileList()  As Variant
Private Sub Form_Activate()

    On Error GoTo ErrorHandler

    Dim i As Integer
Dim ichkcnt  As Integer
    ichkcnt = 0
    On Error GoTo ErrorHandler
    Dim strStep  As String
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")

    Dim orgFile As String
    Call TimeOut(2)
    For i = LBound(fileList) To UBound(fileList)
        orgFile = Mid(Trim(fileList(i)), 1, Len(Trim(fileList(i))) - 4)
        If fs.FileExists(App.Path & "\" & fileList(i)) Then
            If fs.FileExists(App.Path & "\" & orgFile) Then
Retry:
                strStep = "retry"
                fs.DeleteFile App.Path & "\" & orgFile
                strStep = ""
            End If
            strStep = ""
            fs.CopyFile App.Path & "\" & Trim(fileList(i)), App.Path & "\" & orgFile
            fs.DeleteFile App.Path & "\" & Trim(fileList(i))
            
            
        End If
    Next i



ExitHandler:
    Set fs = Nothing
    Dim strExe As String
    strExe = GetSetting("ORAIB_TAB_MAN_BY_HAK", "ORAIB_VERSION", "app_name", "")
    'strExe = App.Path & "\" & strExe
    strExe = strExe

    Shell strExe, vbNormalFocus

    End

ErrorHandler:

    If Err.Number = 70 And strStep = "retry" Then
        Call TimeOut(1)
        ichkcnt = ichkcnt + 1
        If ichkcnt < 4 Then
            Err.Clear
            On Error GoTo ErrorHandler
            GoTo Retry
        End If
    End If
    Set fs = Nothing
    MsgBox Err.Description, vbCritical, "ü������ ����"
End Sub

'�μ��� �ð�(��)��ŭ ��ٸ���.
Private Sub TimeOut(PauseTime As Integer)
    Dim Start, Finish, TotalTime

    Start = Timer   ' ���� �ð��� �����մϴ�.
    Do While Timer < Start + PauseTime
       DoEvents   ' �ٸ� ���ν����� �ѱ�ϴ�.
    Loop
    Finish = Timer   ' ���� �ð��� �����մϴ�.
    TotalTime = Finish - Start   ' ��ü �ð��� ����մϴ�.

End Sub

Private Sub Form_Load()
    Dim i As Integer

    Dim NextLocal
    Dim FullSpec

    FullSpec = App.Path & "\*.tmp"
    i = -1
    NextLocal = Dir(FullSpec, vbDirectory + vbNormal)
    Do While NextLocal <> ""
            FullSpec = App.Path & "\" & NextLocal

        On Error Resume Next
        If (GetAttr(FullSpec) And vbDirectory) = vbDirectory Then

        Else
            i = i + 1
            ReDim Preserve fileList(0 To i)
            fileList(i) = NextLocal
        End If
        NextLocal = Dir
    Loop
   If i < 0 Then
        GoTo ExitHandler
   End If
   Exit Sub
ExitHandler:
    Dim strExe As String
    strExe = GetSetting("ORAIB_TAB_MAN_BY_HAK", "ORAIB_VERSION", "app_name", "")
    'strExe = App.Path & "\" & strExe
    strExe = strExe

    Shell strExe, vbNormalFocus

    End
End Sub

Private Sub imgLoginMain_Click()

End Sub
