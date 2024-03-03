VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSock 
   Caption         =   "Socket Test"
   ClientHeight    =   11475
   ClientLeft      =   6270
   ClientTop       =   2460
   ClientWidth     =   16755
   Icon            =   "frmSock.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11475
   ScaleWidth      =   16755
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   555
      Left            =   8460
      TabIndex        =   5
      Top             =   60
      Width           =   3615
   End
   Begin VB.TextBox txtRecv 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3195
      IMEMode         =   8  '영문
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   4
      Top             =   8220
      Width           =   16635
   End
   Begin VB.TextBox txtSend 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   7395
      IMEMode         =   8  '영문
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   3
      Top             =   720
      Width           =   16635
   End
   Begin VB.TextBox txtPort 
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
      Left            =   2400
      TabIndex        =   2
      Top             =   180
      Width           =   2295
   End
   Begin VB.TextBox txtIP 
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
      Left            =   60
      TabIndex        =   1
      Top             =   180
      Width           =   2295
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "전송"
      Height          =   555
      Left            =   4920
      TabIndex        =   0
      Top             =   60
      Width           =   3615
   End
   Begin MSWinsockLib.Winsock winSock 
      Left            =   15600
      Top             =   180
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    'winSock.RemoteHostIP =

    
    'winSock.RemotePort =

    
    
 
End Sub

Private Sub cmdClear_Click()
    txtSend.Text = ""
    txtRecv.Text = ""
End Sub

Private Sub cmdSend_Click()
    winSock.Close
    winSock.Connect txtIP.Text, txtPort.Text

End Sub

Private Sub Form_Load()
    txtPort.Text = "20000"
    txtIP.Text = winSock.LocalIP
    
    
    
    Dim strMsg As String


strMsg = strMsg & "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
strMsg = strMsg & "<teleMsg>" & vbCrLf
strMsg = strMsg & "    <header>" & vbCrLf
strMsg = strMsg & "        <globId>20130320HomePage00001844343917357900</globId>" & vbCrLf
strMsg = strMsg & "        <rcvmesgSrvId>OCEL_I0200116</rcvmesgSrvId>" & vbCrLf
strMsg = strMsg & "        <rsltRcvmesgSrvId></rsltRcvmesgSrvId>" & vbCrLf
strMsg = strMsg & "        <mciIntfId>OCEL_I0200116</mciIntfId>" & vbCrLf
strMsg = strMsg & "        <mciIntfIdVrsn></mciIntfIdVrsn>" & vbCrLf
strMsg = strMsg & "        <dlKndCd>D</dlKndCd>" & vbCrLf
strMsg = strMsg & "        <trgtSysCd>EOT</trgtSysCd>" & vbCrLf
strMsg = strMsg & "        <respSysNodeNo></respSysNodeNo>" & vbCrLf
strMsg = strMsg & "        <outKndCd>000</outKndCd>" & vbCrLf
strMsg = strMsg & "        <tgrmCntSq></tgrmCntSq>" & vbCrLf
strMsg = strMsg & "        <frstTrnmIp>172.15.15.47</frstTrnmIp>" & vbCrLf
strMsg = strMsg & "        <frstTrnmMac></frstTrnmMac>" & vbCrLf
strMsg = strMsg & "        <bltBrNo></bltBrNo>" & vbCrLf
strMsg = strMsg & "        <bltTmlNo></bltTmlNo>" & vbCrLf
strMsg = strMsg & "        <mesgTrcNo></mesgTrcNo>" & vbCrLf
strMsg = strMsg & "        <envInfoDvcd>D</envInfoDvcd>" & vbCrLf
strMsg = strMsg & "        <frstTrnmSysCd>CEL</frstTrnmSysCd>" & vbCrLf
strMsg = strMsg & "        <trnmSysCd>CEL</trnmSysCd>" & vbCrLf
strMsg = strMsg & "        <demdRespDvcd>Q</demdRespDvcd>" & vbCrLf
strMsg = strMsg & "        <syncDvCd>S</syncDvCd>" & vbCrLf
strMsg = strMsg & "        <mesgDemdDttm>20130320184434391</mesgDemdDttm>" & vbCrLf
strMsg = strMsg & "        <mesgRespDttm></mesgRespDttm>" & vbCrLf
strMsg = strMsg & "        <trttRsltDvcd></trttRsltDvcd>" & vbCrLf
strMsg = strMsg & "        <errSysCd></errSysCd>" & vbCrLf
strMsg = strMsg & "        <errStrdCd></errStrdCd>" & vbCrLf
strMsg = strMsg & "        <errStrdMesg></errStrdMesg>" & vbCrLf
strMsg = strMsg & "        <emNo>GUEST</emNo>" & vbCrLf
strMsg = strMsg & "        <deptCd>999</deptCd>" & vbCrLf
strMsg = strMsg & "        <resveLtrs></resveLtrs>" & vbCrLf
strMsg = strMsg & "    </header>" & vbCrLf
strMsg = strMsg & "  <body>" & vbCrLf
strMsg = strMsg & "         <CUST_NM></CUST_NM>" & vbCrLf
strMsg = strMsg & "         <NBRN>8502182188419</NBRN>" & vbCrLf
strMsg = strMsg & "   </body>" & vbCrLf
strMsg = strMsg & "   <end>" & vbCrLf
strMsg = strMsg & "        <teleMsgEndCd>$$</teleMsgEndCd>" & vbCrLf
strMsg = strMsg & "   </end>" & vbCrLf
strMsg = strMsg & "</teleMsg>" & vbCrLf
    
    txtSend.Text = strMsg
    
    
End Sub

Private Sub winSock_Connect()
    
    
    
    Dim strSend As String
    Dim ilen As Integer
    
    strSend = Trim(txtSend.Text)
    strSend = Replace(strSend, vbCrLf, "")
    
    ilen = Len(strSend) + 2
    
    Dim slen As String
    Dim ilenlen As Integer
    
    ilenlen = Len(CStr(ilen))
    slen = CStr(ilen)
    ilen = ilenlen
    If ilen = 1 Then
        slen = "00000" & slen
    ElseIf ilen = 2 Then
        slen = "0000" & slen
    ElseIf ilen = 3 Then
        slen = "000" & slen
    ElseIf ilen = 4 Then
        slen = "00" & slen
    ElseIf ilen = 5 Then
        slen = "0" & slen
    End If




    winSock.SendData slen & "OK" & strSend
    

End Sub

Private Sub winSock_DataArrival(ByVal bytesTotal As Long)
    Dim sData As String
    winSock.GetData sData
    'MsgBox sData
    txtRecv.Text = sData
    
    MsgBox "recv sucess!!!"
End Sub

Private Sub winSock_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description
End Sub


Private Sub winSock_SendComplete()
    MsgBox "send sucess!!!"
    
End Sub


