VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMakeSrc 
   Caption         =   "���κ��� �ҽ� �ϰ� ����"
   ClientHeight    =   12255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   23175
   Icon            =   "frmMakeSrc.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12255
   ScaleWidth      =   23175
   Begin TabDlg.SSTab SSTab1 
      Height          =   11955
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   22815
      _ExtentX        =   40243
      _ExtentY        =   21087
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Bo(Main)"
      TabPicture(0)   =   "frmMakeSrc.frx":048A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdMakeSrc"
      Tab(0).Control(1)=   "cmdExcel"
      Tab(0).Control(2)=   "fpSrc"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "BO(Popup)"
      TabPicture(1)   =   "frmMakeSrc.frx":04A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "SVC"
      TabPicture(2)   =   "frmMakeSrc.frx":04C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdMakeSrcSVC"
      Tab(2).Control(1)=   "cmdExcelSVC"
      Tab(2).Control(2)=   "fpSrcSVC"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "�����"
      TabPicture(3)   =   "frmMakeSrc.frx":04DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fpBG"
      Tab(3).Control(1)=   "cmbBGMake"
      Tab(3).Control(2)=   "cmdBGExcel"
      Tab(3).Control(3)=   "Command1"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "frmMakeSrc.frx":04FA
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "fpRequest"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Command2"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      Begin VB.CommandButton Command2 
         Caption         =   "�۾�"
         Height          =   555
         Left            =   600
         TabIndex        =   12
         Top             =   480
         Width           =   3315
      End
      Begin VB.CommandButton Command1 
         Caption         =   "����� biz �ҽ� ����"
         Height          =   555
         Left            =   -67920
         TabIndex        =   10
         Top             =   11160
         Width           =   3315
      End
      Begin VB.CommandButton cmdBGExcel 
         Caption         =   "����� Excel �б�"
         Height          =   555
         Left            =   -74520
         TabIndex        =   9
         Top             =   11160
         Width           =   3315
      End
      Begin VB.CommandButton cmbBGMake 
         Caption         =   "����� SVC �ҽ� ����"
         Height          =   555
         Left            =   -71280
         TabIndex        =   8
         Top             =   11160
         Width           =   3315
      End
      Begin VB.CommandButton cmdMakeSrcSVC 
         Caption         =   "SVC �ҽ� ����"
         Height          =   795
         Left            =   -71220
         TabIndex        =   6
         Top             =   11040
         Width           =   3315
      End
      Begin VB.CommandButton cmdExcelSVC 
         Caption         =   "SVC Excel �б�"
         Height          =   795
         Left            =   -74640
         TabIndex        =   5
         Top             =   11040
         Width           =   3315
      End
      Begin VB.CommandButton cmdMakeSrc 
         Caption         =   "�ҽ� ����"
         Height          =   795
         Left            =   -71340
         TabIndex        =   3
         Top             =   10920
         Width           =   3315
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "Excel �б�"
         Height          =   795
         Left            =   -74760
         TabIndex        =   1
         Top             =   10920
         Width           =   3315
      End
      Begin FPSpreadADO.fpSpread fpSrc 
         Height          =   10395
         Left            =   -74760
         TabIndex        =   2
         Top             =   360
         Width           =   21315
         _Version        =   393216
         _ExtentX        =   37597
         _ExtentY        =   18336
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
         SpreadDesigner  =   "frmMakeSrc.frx":0516
      End
      Begin FPSpreadADO.fpSpread fpSrcSVC 
         Height          =   10395
         Left            =   -74880
         TabIndex        =   4
         Top             =   600
         Width           =   21315
         _Version        =   393216
         _ExtentX        =   37597
         _ExtentY        =   18336
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
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
         Protect         =   0   'False
         SpreadDesigner  =   "frmMakeSrc.frx":1E1C0
      End
      Begin FPSpreadADO.fpSpread fpBG 
         Height          =   10395
         Left            =   -74640
         TabIndex        =   7
         Top             =   720
         Width           =   21315
         _Version        =   393216
         _ExtentX        =   37597
         _ExtentY        =   18336
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
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
         Protect         =   0   'False
         SpreadDesigner  =   "frmMakeSrc.frx":3BE97
      End
      Begin FPSpreadADO.fpSpread fpRequest 
         Height          =   10395
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   21315
         _Version        =   393216
         _ExtentX        =   37597
         _ExtentY        =   18336
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
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
         Protect         =   0   'False
         SpreadDesigner  =   "frmMakeSrc.frx":59B6E
      End
   End
End
Attribute VB_Name = "frmMakeSrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim iROW_CNT As Integer
Dim sClassLname As String
Dim sClassPname As String
Dim sPackageDotName As String
Dim sPackageSlashName As String
Dim sUserId As String
Dim sUserName As String
Dim sRevision As String

Dim s�޼ҵ�id, s�޼ҵ��, sINBVO, sOUTBVO As String
Dim sVO������, sVOid, sVO��, sVO��� As String
Dim sSVC�޼ҵ�id, sSVCInVOid, sSVCOutVOid As String

'����� ���� ���� ����
Dim s����ڼ���ID, s�����BizID, s�����SVC��, s�����Biz��� As String


Dim sPname, sLname, sPath As String

'" & s����ڼ���ID & "
'" & s�����SVC�� & "
'" & s�����Biz��� & "
'" & fService3() & "

Private Function fService3() As String
    
    fService3 = LCase(Left(s����ڼ���ID, 1)) & Mid(s����ڼ���ID, 2, 3)


End Function

'����� �ҽ� ���� ��ư Ŭ��
Private Sub cmbBGMake_Click()
On Error GoTo Err_Handler
With fpBG
    Dim stmFileControl As Stream
    Dim sPrintString As String
    
    
        
    iROW_CNT = 0
    
    
    Dim sFilename As String
    iROW_CNT = iROW_CNT + 1
        
LOOP_START:
    
    
    
    Set stmFileControl = New Stream
    
    
    stmFileControl.Open
    stmFileControl.Position = 0
    stmFileControl.Charset = "UTF-8"
    
    
    
    .Row = iROW_CNT
    .Col = 1: s����ڼ���ID = .Value
    .Col = 2: s�����Biz��� = .Value
    s�����BizID = .Value & "Biz"
    .Col = 3: s�����SVC�� = .Value
    

    
    'file name ����
    sFilename = s����ڼ���ID & ".so"
    
    
    '����� ���ڿ� ����
    sPrintString = sMake�����SVC
    
    
    'stmFileControl.WriteText (sPrintString)
    'stmFileControl.SaveToFile "c:\src\" & sFileName, adSaveCreateOverWrite
    'stmFileControl.Close
            
    'Set stmFileControl = Nothing
                    
                    
    Dim objStreamUTF8 As Stream
    Set objStreamUTF8 = New Stream
    
    
    Dim objStreamUTF8NoBOM As Stream
    Set objStreamUTF8NoBOM = New Stream
    
    
    With objStreamUTF8
        .Charset = "UTF-8"
        .Open
        .WriteText sPrintString
        .Position = 0
        .Type = adTypeText
        .Position = 3
    End With
    
    With objStreamUTF8NoBOM
        .Type = adTypeBinary
        .Open
        objStreamUTF8.CopyTo objStreamUTF8NoBOM
        .SaveToFile "c:\src\" & sFilename, adSaveCreateOverWrite
        
    End With
    
    objStreamUTF8.Close
    objStreamUTF8NoBOM.Close
    
    Set objStreamUTF8 = Nothing
    Set objStreamUTF8NoBOM = Nothing
                    
                    
                    
                    
                    
                    
    .Row = iROW_CNT
    .Col = 1
    
    iROW_CNT = iROW_CNT + 1
    .Row = iROW_CNT
    '���� �ƴϸ� �ݺ��Ѵ�.
    If UCase(Trim(.Value)) <> "END" Then
        'iROW_CNT = iROW_CNT + 1
        GoTo LOOP_START
    End If
End With

    gfMsgInf "�Ϸ�Ǿ����ϴ�."
    Exit Sub
Err_Handler:
    gfMsgWrn (Err.Description)
End Sub




Private Function sMake�����SVC() As String
Dim sSrc As String
Dim sMethodNm As String



With fpBG
sSrc = ""
sSrc = sSrc & "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
sSrc = sSrc & vbCrLf & "<ns14:serviceObject xmlns:ns2=""http://www.tmax.co.kr/proobject/resource"" xmlns:ns3=""http://www.tmax.co.kr/proobject/testcase"" "
sSrc = sSrc & "xmlns:ns4=""http://www.tmax.co.kr/proobject/message"" xmlns:ns5=""http://www.tmax.co.kr/proobject/dto"" "
sSrc = sSrc & "xmlns:ns6=""http://www.tmaxsoft.co.kr/proobject/testsuite"" xmlns:ns7=""http://www.tmax.co.kr/proobject/dto/validator"" "
sSrc = sSrc & "xmlns:ns8=""http://www.tmax.co.kr/proobject/serviceobject-automatic"" xmlns:ns9=""http://www.tmax.co.kr/proobject/sourcecode"" "
sSrc = sSrc & "xmlns:ns10=""http://www.tmax.co.kr/proobject/flow"" xmlns:ns11=""http://www.tmax.co.kr/proobject/dataobjectfactory"" "
sSrc = sSrc & "xmlns:ns12=""http://www.tmax.co.kr/proobject/queryobject"" xmlns:ns13=""http://www.tmax.co.kr/proobject/bizobject"" "
sSrc = sSrc & "xmlns:ns14=""http://www.tmax.co.kr/proobject/serviceobject"" xmlns:ns15=""http://www.tmax.co.kr/proobject/dataobject"" "
sSrc = sSrc & "xmlns:ns16=""http://www.tmax.co.kr/proobject/taskobject"" xmlns:ns17=""http://www.tmax.co.kr/proobject/jobobject"" "
sSrc = sSrc & "xmlns:ns18=""http://www.tmax.co.kr/proobject/servicegroup"" xmlns:ns19=""http://www.tmax.co.kr/proobject/aspectj"" "
sSrc = sSrc & "xmlns:ns20=""http://www.tmax.co.kr/proobject/mapping"" xmlns:ns21=""http://www.tmax.co.kr/proobject/application/runtime"" "
sSrc = sSrc & "xmlns:ns22=""http://www.tmaxsoft.co.kr/proobject/testOperation"" xmlns:ns23=""http://www.tmax.co.kr/proobject/restriction_codes"" "
sSrc = sSrc & "xmlns:ns24=""http://www.example.org/externalObjectConfig"" xmlns:ns25=""http://www.tmax.co.kr/proobject/application"" "
sSrc = sSrc & "xmlns:ns26=""http://www.tmax.co.kr/proobject/contents"" xmlns:ns27=""http://www.example.org/testData"" "
sSrc = sSrc & "xmlns:ns28=""http://www.tmaxsoft.co.kr/proobject/testresult"" xmlns:ns29=""http://www.tmax.co.kr/proobject/siteConfig"" "
sSrc = sSrc & "xmlns:ns30=""http://www.tmax.co.kr/proobject/serverConfig"" xmlns:ns31=""http://www.tmax.co.kr/proobject/probuilder_config"" "
sSrc = sSrc & "xmlns:ns32=""http://www.tmax.co.kr/proobject/property"" xmlns:ns33=""http://www.tmax.co.kr/proobject/serviceGroup"" "
sSrc = sSrc & "identifier=""general"" "
sSrc = sSrc & "name=""" & s����ڼ���ID & """ "
sSrc = sSrc & "resourceId=""" & s����ڼ���ID & """ "
sSrc = sSrc & "logicalName=""" & s�����SVC�� & """ "
sSrc = sSrc & "physicalName=""" & s����ڼ���ID & """ "
sSrc = sSrc & "resourceType=""SERVICE_OBJECT"" "
sSrc = sSrc & "resourcePackage=""hf.cor.hg.bg.svc." & fService3() & """ "
sSrc = sSrc & "creator=""ns042"" "  '���� ����
sSrc = sSrc & "revision=""1"" "
sSrc = sSrc & "description=""@���ϸ� : " & s����ڼ���ID & "&#xD;&#xA;"
sSrc = sSrc & "@���� : " & s�����SVC�� & " wframe ����&#xD;&#xA;"
sSrc = sSrc & "@�ۼ��� : ��ټ�(ns042)&#xD;&#xA;--------------------------------------------------------------------&#xD;&#xA; "  '���⺯��
sSrc = sSrc & "* �������� : "
sSrc = sSrc & "������ : ��û��(SR ��ȣ) : "
sSrc = sSrc & "������ �޼ҵ� : ��������&#xD;&#xA;--------------------------------------------------------------------&#xD;&#xA;* "
sSrc = sSrc & "�ۼ����� : 2020.04.03 : "
sSrc = sSrc & "�ۼ��� : ��ټ�(ns042),  "   '���⺯��
sSrc = sSrc & "�ۼ����� : "
sSrc = sSrc & "�����ۼ�"" "
sSrc = sSrc & "resourcePath=""hf/cor/hg/bg/svc/" & fService3() & """ "
sSrc = sSrc & "modifier=""ns042"" "         '���⺯��
sSrc = sSrc & "created=""2020-04-06 19:04:28"" "
sSrc = sSrc & "lastUpdated=""2020-04-06 19:04:28""> "
sSrc = sSrc & vbCrLf & "    <ns14:inputDo xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:type=""ns15:dataObjectType"" resourceId=""urn:proobject:hf.cor.hg.bg.svc." & fService3() & ":" & s����ڼ���ID & """ physicalName=""DummyVo"" resourceType=""DATA_OBJECT"" serviceGroupName=""hf2_cor_service"" resourcePackage=""hf.cor._templet.vo"" revision=""0"" resourcePath=""hf/cor/_templet/vo""/>"
sSrc = sSrc & vbCrLf & "    <ns14:outputDo xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:type=""ns15:dataObjectType"" resourceId=""urn:proobject:hf.cor.hg.bg.svc." & fService3() & ":" & s����ڼ���ID & """ physicalName=""DummyVo"" resourceType=""DATA_OBJECT"" serviceGroupName=""hf2_cor_service"" resourcePackage=""hf.cor._templet.vo"" revision=""0"" resourcePath=""hf/cor/_templet/vo""/>"
sSrc = sSrc & vbCrLf & "    <ns14:opCodeEnable>true</ns14:opCodeEnable>"
sSrc = sSrc & vbCrLf & "    <ns14:opCode>"

                                                                                         
sSrc = sSrc & vbCrLf & "        <ns14:opCode opCode=""sel" & s�����Biz��� & """ svcName=""sel" & s�����Biz��� & """ inputDO=""hf.cor.hg.bg.vo.svo." & fService3() & ".HgbgSel" & s�����Biz��� & "InSvo"" outputDO=""hf.cor.hg.bg.vo.svo." & fService3() & ".HgbgSel" & s�����Biz��� & "OutSvo""/>"

sPname = "HgbgSel" & s�����Biz��� & "InSvo"
sLname = s�����SVC�� & " ��ȸ InSvo"
sPath = Replace("hf.cor.hg.bg.vo.svo." & fService3(), ".", "/")
Call subMake�����VO


sPname = "HgbgSel" & s�����Biz��� & "OutSvo"
sLname = s�����SVC�� & " ��ȸ OutSvo"
sPath = Replace("hf.cor.hg.bg.vo.svo." & fService3(), ".", "/")
Call subMake�����VO

sSrc = sSrc & vbCrLf & "        <ns14:opCode opCode=""prcs" & s�����Biz��� & """ svcName=""prcs" & s�����Biz��� & """ inputDO=""hf.cor.hg.bg.vo.svo." & fService3() & ".HgbgPrcs" & s�����Biz��� & "InSvo"" outputDO=""hf.cor.hg.bg.vo.svo." & fService3() & ".HgbgPrcs" & s�����Biz��� & "OutSvo""/>"

sPname = "HgbgPrcs" & s�����Biz��� & "InSvo"
sLname = s�����SVC�� & " ó�� InSvo"
sPath = Replace("hf.cor.hg.bg.vo.svo." & fService3(), ".", "/")
Call subMake�����VO



sPname = "HgbgPrcs" & s�����Biz��� & "OutSvo"
sLname = s�����SVC�� & " ó�� OutSvo"
sPath = Replace("hf.cor.hg.bg.vo.svo." & fService3(), ".", "/")
Call subMake�����VO



sSrc = sSrc & vbCrLf & "    </ns14:opCode>"
sSrc = sSrc & vbCrLf & "    <ns14:member>"
sSrc = sSrc & vbCrLf & "        <ns14:variable kind=""DTO"" name=""input"" arrayKind=""none"" size=""0"" scope=""member"" initObject="""" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"" inject=""NONE"">"
sSrc = sSrc & vbCrLf & "            <ns10:do physicalName=""DummyVo"" resourcePath=""hf/cor/_templet/vo"" inOutType=""INPUT""/>"
sSrc = sSrc & vbCrLf & "        </ns14:variable>"
sSrc = sSrc & vbCrLf & "        <ns14:variable kind=""DTO"" name=""output"" arrayKind=""none"" size=""0"" scope=""member"" initObject="""" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"" inject=""NONE"">"
sSrc = sSrc & vbCrLf & "            <ns10:do physicalName=""DummyVo"" resourcePath=""hf/cor/_templet/vo"" inOutType=""OUTPUT""/>"
sSrc = sSrc & vbCrLf & "        </ns14:variable>"
sSrc = sSrc & vbCrLf & "        <ns14:variable kind=""OBJECT"" name=""bizHgbg" & s�����Biz��� & """ scope=""member"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"" inject=""NONE"">"
sSrc = sSrc & vbCrLf & "            <ns10:userDefined>hf.cor.hg.bg.biz.sbiz." & fService3() & ".Hgbg" & s�����Biz��� & "Biz</ns10:userDefined>"
sSrc = sSrc & vbCrLf & "        </ns14:variable>"
sSrc = sSrc & vbCrLf & "    </ns14:member>"
sSrc = sSrc & vbCrLf & "    <ns14:wholeImport>"
sSrc = sSrc & vbCrLf & "        <ns14:importType>com.tmax.proobject.model.service.FlexibleServiceObject</ns14:importType>"
sSrc = sSrc & vbCrLf & "        <ns14:importType>com.tmax.proobject.core.ServiceObject</ns14:importType>"
sSrc = sSrc & vbCrLf & "        <ns14:importType>com.tmax.proobject.model.business.BusinessObject</ns14:importType>"
sSrc = sSrc & vbCrLf & "        <ns14:importType>com.tmax.proobject.model.dataobject.DataObject</ns14:importType>"
sSrc = sSrc & vbCrLf & "        <ns14:importType>com.tmax.proobject.engine.servicemanager.*</ns14:importType>"
sSrc = sSrc & vbCrLf & "        <ns14:importType>com.tmax.proobject.logger.ProObjectLogger</ns14:importType>"
sSrc = sSrc & vbCrLf & "        <ns14:importType>com.tmax.proobject.logger.application.ServiceLogger</ns14:importType>"
sSrc = sSrc & vbCrLf & "        <ns14:importType>com.tmax.proobject.model.dto.*</ns14:importType>"
sSrc = sSrc & vbCrLf & "        <ns14:importType>com.tmax.proobject.core.DataSource</ns14:importType>"
sSrc = sSrc & vbCrLf & "        <ns14:importType>javax.inject.Inject</ns14:importType>"
sSrc = sSrc & vbCrLf & "        <ns14:importType>com.tmax.proobject.engine.waitobject.WaitObject</ns14:importType>"
sSrc = sSrc & vbCrLf & "        <ns14:importType>java.math.BigDecimal</ns14:importType>"
sSrc = sSrc & vbCrLf & "        <ns14:importType>java.util.List</ns14:importType>"
sSrc = sSrc & vbCrLf & "        <ns14:importType>java.util.ArrayList</ns14:importType>"
sSrc = sSrc & vbCrLf & "        <ns14:importType>java.util.logging.Level</ns14:importType>"
sSrc = sSrc & vbCrLf & "        <ns14:importType>java.io.File</ns14:importType>"
sSrc = sSrc & vbCrLf & "        <ns14:importType>com.tmax.proobject.model.transaction.TransactionType</ns14:importType>"
sSrc = sSrc & vbCrLf & "        <ns14:importType>com.tmax.proobject.service.RCallReturn</ns14:importType>"
sSrc = sSrc & vbCrLf & "        <ns14:importType>com.tmax.proobject.dataobject.util.DataDefinitionExecutor</ns14:importType>"
sSrc = sSrc & vbCrLf & "        <ns14:importType>hf.cor._templet.vo.DummyVo</ns14:importType>"
sSrc = sSrc & vbCrLf & "        <ns14:importType>hf.cor.hg.bg.vo.svo." & fService3() & ".HgbgSel" & s�����Biz��� & "OutSvo</ns14:importType>"
sSrc = sSrc & vbCrLf & "        <ns14:importType>hf.cor.hg.bg.vo.svo." & fService3() & ".HgbgPrcs" & s�����Biz��� & "OutSvo</ns14:importType>"
sSrc = sSrc & vbCrLf & "    </ns14:wholeImport>"
sSrc = sSrc & vbCrLf & "    <ns14:serviceMethod methodName=""sel" & s�����Biz��� & """ displayName=""" & s�����SVC�� & """ modifier=""public"" comments="""
sSrc = sSrc & " @�޼���� : sel" & s�����Biz��� & "&#xD;&#xA;"
sSrc = sSrc & " @���� : " & s�����SVC�� & "&#xD;&#xA;"
sSrc = sSrc & " @�Է� : HgbgSel" & s�����Biz��� & "InSvo&#xD;&#xA;"
sSrc = sSrc & " @��� : HgbgSel" & s�����Biz��� & "OutSvo"">"
sSrc = sSrc & vbCrLf & "        <ns10:flowModule>"
sSrc = sSrc & vbCrLf & "            <ns10:entryModule>"
sSrc = sSrc & vbCrLf & "                <ns10:name>" & s�����SVC�� & "</ns10:name>"
sSrc = sSrc & vbCrLf & "                <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:id>0</ns10:id>"
sSrc = sSrc & vbCrLf & "                    <ns10:id>2</ns10:id>"
sSrc = sSrc & vbCrLf & "                    <ns10:id>4</ns10:id>"
sSrc = sSrc & vbCrLf & "                </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                <ns10:variable kind=""DTO"" name=""outBvo"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <ns10:do physicalName=""Hgbg" & s�����Biz��� & "01OutBvo"" logicalName=""" & s�����SVC�� & " OutBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"
sSrc = sSrc & vbCrLf & "                </ns10:variable>"
sSrc = sSrc & vbCrLf & "                <ns10:variable kind=""DTO"" name=""inBvo"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <ns10:do physicalName=""Hgbg" & s�����Biz��� & "01InBvo"" logicalName=""" & s�����SVC�� & " InBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"
sSrc = sSrc & vbCrLf & "                </ns10:variable>"
sSrc = sSrc & vbCrLf & "                <ns10:variable kind=""DTO"" name=""outSvo"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <ns10:do physicalName=""HgbgSel" & s�����Biz��� & "OutSvo"" logicalName=""" & s�����SVC�� & " OutSvo"" resourcePath=""hf/cor/hg/bg/vo/svo/" & fService3() & """/>"
sSrc = sSrc & vbCrLf & "                </ns10:variable>"
sSrc = sSrc & vbCrLf & "            </ns10:entryModule>"
sSrc = sSrc & vbCrLf & "            <ns10:innerModules>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""block"" nodeId=""0"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>���� ���� ó��</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>���� ���� ó��</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>1</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition/>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""block"" nodeId=""2"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>���� �� ó��</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>���� �� ó��</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>6</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition/>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""block"" nodeId=""4"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>���� ���� ó��</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>���� ���� ó��</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>5</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition/>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "            </ns10:innerModules>"
sSrc = sSrc & vbCrLf & "            <ns10:xorModules/>"
sSrc = sSrc & vbCrLf & "            <ns10:virtualModules>"
sSrc = sSrc & vbCrLf & "                <ns10:virtualModule nodeId=""1"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>�Է°� Ȯ��</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>�Է°� Ȯ��</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:code>            logger.info(""#########################"" );&#xD;"
sSrc = sSrc & vbCrLf & "            logger.info(""Svc InSvo \n"" + inSvo.toString());&#xD;"
sSrc = sSrc & vbCrLf & "            logger.info(""#########################"" );&#xD;"
sSrc = sSrc & vbCrLf & "            inBvo.setRvewInqVo(inSvo.getRvewInqVo());</ns10:code>"
sSrc = sSrc & vbCrLf & "                </ns10:virtualModule>"
sSrc = sSrc & vbCrLf & "                <ns10:virtualModule nodeId=""5"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>�������</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>�������</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:code>         &#xD;"
sSrc = sSrc & vbCrLf & "            return outSvo;</ns10:code>"
sSrc = sSrc & vbCrLf & "                </ns10:virtualModule>"
sSrc = sSrc & vbCrLf & "            </ns10:virtualModules>"
sSrc = sSrc & vbCrLf & "            <ns10:loopModules/>"
sSrc = sSrc & vbCrLf & "            <ns10:bizMethodCalls>"
sSrc = sSrc & vbCrLf & "                <ns10:bizMethodCall nodeId=""6"">"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:bizInstanceInfo scope=""member"" variableName=""hgbg" & s�����Biz��� & "Biz"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                        <ns10:classInfo classPackageName=""hf.cor.hg.bg.biz.sbiz." & fService3() & """ className=""Hgbg" & s�����Biz��� & "Biz""/>"
sSrc = sSrc & vbCrLf & "                    </ns10:bizInstanceInfo>"
sSrc = sSrc & vbCrLf & "                    <ns10:method methodName=""sel" & s�����Biz��� & """ displayName=""" & s�����SVC�� & """>"
sSrc = sSrc & vbCrLf & "                        <ns13:arguments>"
sSrc = sSrc & vbCrLf & "                            <ns10:argument>"
sSrc = sSrc & vbCrLf & "<ns10:variable kind=""DTO"" name=""inBvo"">"
sSrc = sSrc & vbCrLf & "    <ns10:do physicalName=""Hgbg" & s�����Biz��� & "01InBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"
sSrc = sSrc & vbCrLf & "</ns10:variable>"
sSrc = sSrc & vbCrLf & "<ns10:sequence>0</ns10:sequence>"
sSrc = sSrc & vbCrLf & "                            </ns10:argument>"
sSrc = sSrc & vbCrLf & "                        </ns13:arguments>"
sSrc = sSrc & vbCrLf & "                        <ns13:returnType kind=""DTO"">"
sSrc = sSrc & vbCrLf & "                            <ns10:do physicalName=""Hgbg" & s�����Biz��� & "01OutBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"
sSrc = sSrc & vbCrLf & "                        </ns13:returnType>"
sSrc = sSrc & vbCrLf & "                    </ns10:method>"
sSrc = sSrc & vbCrLf & "                    <ns10:in>"
sSrc = sSrc & vbCrLf & "                        <ns10:transformation/>"
sSrc = sSrc & vbCrLf & "                        <ns10:byPass>"
sSrc = sSrc & vbCrLf & "                            <ns10:memberVariable name=""inBvo"" order=""0""/>"
sSrc = sSrc & vbCrLf & "                        </ns10:byPass>"
sSrc = sSrc & vbCrLf & "                    </ns10:in>"
sSrc = sSrc & vbCrLf & "                    <ns10:out>"
sSrc = sSrc & vbCrLf & "                        <ns10:transformation/>"
sSrc = sSrc & vbCrLf & "                        <ns10:byPass>"
sSrc = sSrc & vbCrLf & "                            <ns10:memberVariable name=""outBvo"" order=""0""/>"
sSrc = sSrc & vbCrLf & "                        </ns10:byPass>"
sSrc = sSrc & vbCrLf & "                    </ns10:out>"
sSrc = sSrc & vbCrLf & "                    <ns10:preAssign/>"
sSrc = sSrc & vbCrLf & "                    <ns10:postAssign/>"
sSrc = sSrc & vbCrLf & "                </ns10:bizMethodCall>"
sSrc = sSrc & vbCrLf & "            </ns10:bizMethodCalls>"
sSrc = sSrc & vbCrLf & "            <ns10:serviceCalls/>"
sSrc = sSrc & vbCrLf & "            <ns10:assignModules/>"
sSrc = sSrc & vbCrLf & "            <ns10:deferredModules/>"
sSrc = sSrc & vbCrLf & "            <ns10:dataObjectCalls/>"
sSrc = sSrc & vbCrLf & "            <ns10:replyCalls/>"
sSrc = sSrc & vbCrLf & "        </ns10:flowModule>"
sSrc = sSrc & vbCrLf & "        <ns14:arguments>"
sSrc = sSrc & vbCrLf & "            <ns10:argument>"
sSrc = sSrc & vbCrLf & "                <ns10:variable kind=""DTO"" name=""inSvo"">"
sSrc = sSrc & vbCrLf & "                    <ns10:do physicalName=""HgbgSel" & s�����Biz��� & "InSvo"" resourcePath=""hf/cor/hg/bg/vo/svo/" & fService3() & """/>"


sPname = "HgbgSel" & s�����Biz��� & "InSvo"
sLname = s�����Biz��� & " InSvo"
sPath = Replace("hf.cor.hg.bg.vo.svo." & fService3(), ".", "/")
Call subMake�����VO


sSrc = sSrc & vbCrLf & "                </ns10:variable>"
sSrc = sSrc & vbCrLf & "                <ns10:sequence>0</ns10:sequence>"
sSrc = sSrc & vbCrLf & "            </ns10:argument>"
sSrc = sSrc & vbCrLf & "        </ns14:arguments>"
sSrc = sSrc & vbCrLf & "        <ns14:returnType kind=""DTO"" name=""arg0"">"
sSrc = sSrc & vbCrLf & "            <ns10:do physicalName=""HgbgSel" & s�����Biz��� & "OutSvo"" resourcePath=""hf/cor/hg/bg/vo/svo/" & fService3() & """/>"
sSrc = sSrc & vbCrLf & "        </ns14:returnType>"
sSrc = sSrc & vbCrLf & "        <ns14:classInfo classPackageName=""hf/cor/hg/bg/svc/" & fService3() & """ className=""" & s����ڼ���ID & """/>"
sSrc = sSrc & vbCrLf & "        <ns14:annotations/>"
sSrc = sSrc & vbCrLf & "    </ns14:serviceMethod>"
sSrc = sSrc & vbCrLf & "    <ns14:serviceMethod methodName=""prcs" & s�����Biz��� & """ displayName=""" & s�����SVC�� & """ modifier=""public"" comments="""
sSrc = sSrc & " @�޼���� : sel" & s�����Biz��� & "&#xD;&#xA;"
sSrc = sSrc & " @���� : " & s�����SVC�� & "&#xD;&#xA;"
sSrc = sSrc & " @�Է� : HgbgSel" & s�����Biz��� & "InSvo&#xD;&#xA;"
sSrc = sSrc & " @��� : HgbgSel" & s�����Biz��� & "OutSvo"">"
sSrc = sSrc & vbCrLf & "        <ns10:flowModule>"
sSrc = sSrc & vbCrLf & "            <ns10:entryModule>"
sSrc = sSrc & vbCrLf & "                <ns10:name>" & s�����SVC�� & "</ns10:name>"
sSrc = sSrc & vbCrLf & "                <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:id>0</ns10:id>"
sSrc = sSrc & vbCrLf & "                    <ns10:id>2</ns10:id>"
sSrc = sSrc & vbCrLf & "                    <ns10:id>3</ns10:id>"
sSrc = sSrc & vbCrLf & "                </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                <ns10:variable kind=""DTO"" name=""outBvo"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <ns10:do physicalName=""Hgbg" & s�����Biz��� & "02OutBvo"" logicalName=""" & s�����SVC�� & " OutBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"
sSrc = sSrc & vbCrLf & "                </ns10:variable>"
sSrc = sSrc & vbCrLf & "                <ns10:variable kind=""DTO"" name=""inBvo"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <ns10:do physicalName=""Hgbg" & s�����Biz��� & "02InBvo"" logicalName=""" & s�����SVC�� & " InBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"
sSrc = sSrc & vbCrLf & "                </ns10:variable>"
sSrc = sSrc & vbCrLf & "                <ns10:variable kind=""DTO"" name=""outSvo"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <ns10:do physicalName=""HgbgPrcs" & s�����Biz��� & "OutSvo"" logicalName=""" & s�����SVC�� & "�� OutSvo"" resourcePath=""hf/cor/hg/bg/vo/svo/" & fService3() & """/>"
sSrc = sSrc & vbCrLf & "                </ns10:variable>"
sSrc = sSrc & vbCrLf & "            </ns10:entryModule>"
sSrc = sSrc & vbCrLf & "            <ns10:innerModules>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""block"" nodeId=""0"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>���� ���� ó��</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>���� ���� ó��</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>1</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition/>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""block"" nodeId=""2"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>���� �� ó��</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>���� �� ó��</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>5</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition/>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""block"" nodeId=""3"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>���� ���� ó��</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>���� ���� ó��</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>4</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition/>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "            </ns10:innerModules>"
sSrc = sSrc & vbCrLf & "            <ns10:xorModules/>"
sSrc = sSrc & vbCrLf & "            <ns10:virtualModules>"
sSrc = sSrc & vbCrLf & "                <ns10:virtualModule nodeId=""1"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>�Է°� Ȯ��</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>�Է°� Ȯ��</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:code>            logger.info(""#########################"" );&#xD;"
sSrc = sSrc & vbCrLf & "            logger.info(""Svc InSvo \n"" + inSvo.toString());&#xD;"
sSrc = sSrc & vbCrLf & "            logger.info(""#########################"" );&#xD;"
sSrc = sSrc & vbCrLf & "            </ns10:code>"
sSrc = sSrc & vbCrLf & "                </ns10:virtualModule>"
sSrc = sSrc & vbCrLf & "                <ns10:virtualModule nodeId=""4"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>�������</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>�������</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:code>         &#xD;"
sSrc = sSrc & vbCrLf & "            return outSvo;</ns10:code>"
sSrc = sSrc & vbCrLf & "                </ns10:virtualModule>"
sSrc = sSrc & vbCrLf & "            </ns10:virtualModules>"
sSrc = sSrc & vbCrLf & "            <ns10:loopModules/>"
sSrc = sSrc & vbCrLf & "            <ns10:bizMethodCalls>"
sSrc = sSrc & vbCrLf & "                <ns10:bizMethodCall nodeId=""5"">"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:bizInstanceInfo scope=""member"" variableName=""hgbg" & s�����Biz��� & "Biz"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                        <ns10:classInfo classPackageName=""hf.cor.hg.bg.biz.sbiz." & fService3() & """ className=""Hgbg" & s�����Biz��� & "Biz""/>"
sSrc = sSrc & vbCrLf & "                    </ns10:bizInstanceInfo>"
sSrc = sSrc & vbCrLf & "                    <ns10:method methodName=""prcs" & s�����Biz��� & """ displayName=""" & s�����SVC�� & """>"
sSrc = sSrc & vbCrLf & "                        <ns13:arguments>"
sSrc = sSrc & vbCrLf & "                            <ns10:argument>"
sSrc = sSrc & vbCrLf & "<ns10:variable kind=""DTO"" name=""inBvo"">"
sSrc = sSrc & vbCrLf & "    <ns10:do physicalName=""HgbgPrcs" & s�����Biz��� & "02InBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"
sSrc = sSrc & vbCrLf & "</ns10:variable>"
sSrc = sSrc & vbCrLf & "<ns10:sequence>0</ns10:sequence>"
sSrc = sSrc & vbCrLf & "                            </ns10:argument>"
sSrc = sSrc & vbCrLf & "                        </ns13:arguments>"
sSrc = sSrc & vbCrLf & "                        <ns13:returnType kind=""DTO"">"
sSrc = sSrc & vbCrLf & "                            <ns10:do physicalName=""Hgbg" & s�����Biz��� & "OutBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"
sSrc = sSrc & vbCrLf & "                        </ns13:returnType>"
sSrc = sSrc & vbCrLf & "                    </ns10:method>"
sSrc = sSrc & vbCrLf & "                    <ns10:in>"
sSrc = sSrc & vbCrLf & "                        <ns10:transformation/>"
sSrc = sSrc & vbCrLf & "                        <ns10:byPass>"
sSrc = sSrc & vbCrLf & "                            <ns10:memberVariable name=""inBvo"" order=""0""/>"
sSrc = sSrc & vbCrLf & "                        </ns10:byPass>"
sSrc = sSrc & vbCrLf & "                    </ns10:in>"
sSrc = sSrc & vbCrLf & "                    <ns10:out>"
sSrc = sSrc & vbCrLf & "                        <ns10:transformation/>"
sSrc = sSrc & vbCrLf & "                        <ns10:byPass>"
sSrc = sSrc & vbCrLf & "                            <ns10:memberVariable name=""outBvo"" order=""0""/>"
sSrc = sSrc & vbCrLf & "                        </ns10:byPass>"
sSrc = sSrc & vbCrLf & "                    </ns10:out>"
sSrc = sSrc & vbCrLf & "                    <ns10:preAssign/>"
sSrc = sSrc & vbCrLf & "                    <ns10:postAssign/>"
sSrc = sSrc & vbCrLf & "                </ns10:bizMethodCall>"
sSrc = sSrc & vbCrLf & "            </ns10:bizMethodCalls>"
sSrc = sSrc & vbCrLf & "            <ns10:serviceCalls/>"
sSrc = sSrc & vbCrLf & "            <ns10:assignModules/>"
sSrc = sSrc & vbCrLf & "            <ns10:deferredModules/>"
sSrc = sSrc & vbCrLf & "            <ns10:dataObjectCalls/>"
sSrc = sSrc & vbCrLf & "            <ns10:replyCalls/>"
sSrc = sSrc & vbCrLf & "        </ns10:flowModule>"
sSrc = sSrc & vbCrLf & "        <ns14:arguments>"
sSrc = sSrc & vbCrLf & "            <ns10:argument>"
sSrc = sSrc & vbCrLf & "                <ns10:variable kind=""DTO"" name=""inSvo"">"
sSrc = sSrc & vbCrLf & "                    <ns10:do physicalName=""Hgbg" & s�����Biz��� & "InSvo"" resourcePath=""hf/cor/hg/bg/vo/svo/" & fService3() & """/>"
sSrc = sSrc & vbCrLf & "                </ns10:variable>"
sSrc = sSrc & vbCrLf & "                <ns10:sequence>0</ns10:sequence>"
sSrc = sSrc & vbCrLf & "            </ns10:argument>"
sSrc = sSrc & vbCrLf & "        </ns14:arguments>"
sSrc = sSrc & vbCrLf & "        <ns14:returnType kind=""DTO"" name=""arg0"">"
sSrc = sSrc & vbCrLf & "            <ns10:do physicalName=""HgbgPrcs" & s�����Biz��� & "OutSvo"" resourcePath=""hf/cor/hg/bg/vo/svo/" & fService3() & """/>"
sSrc = sSrc & vbCrLf & "        </ns14:returnType>"
sSrc = sSrc & vbCrLf & "        <ns14:classInfo classPackageName=""hf/cor/hg/bg/svc/" & fService3() & """ className=""" & s����ڼ���ID & """/>"
sSrc = sSrc & vbCrLf & "        <ns14:annotations/>"
sSrc = sSrc & vbCrLf & "    </ns14:serviceMethod>"
sSrc = sSrc & vbCrLf & "    <ns14:interfaces/>"
sSrc = sSrc & vbCrLf & "</ns14:serviceObject>"


                                                                                                                                                                                                                              
End With
sMake�����SVC = sSrc
End Function


Private Function sMake�����biz() As String
Dim sSrc As String
Dim sMethodNm As String
sSrc = ""
sSrc = sSrc & "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
sSrc = sSrc & vbCrLf & "<ns13:bizObject xmlns:ns2=""http://www.tmax.co.kr/proobject/resource"" xmlns:ns3=""http://www.tmax.co.kr/proobject/testcase"" "
sSrc = sSrc & " xmlns:ns4=""http://www.tmax.co.kr/proobject/message"" xmlns:ns5=""http://www.tmax.co.kr/proobject/dto"" xmlns:ns6=""http://www.tmaxsoft.co.kr/proobject/testsuite"" "
sSrc = sSrc & " xmlns:ns7=""http://www.tmax.co.kr/proobject/dto/validator"" xmlns:ns8=""http://www.tmax.co.kr/proobject/serviceobject-automatic"" "
sSrc = sSrc & " xmlns:ns9=""http://www.tmax.co.kr/proobject/sourcecode"" xmlns:ns10=""http://www.tmax.co.kr/proobject/flow"" xmlns:ns11=""http://www.tmax.co.kr/proobject/dataobjectfactory"" "
sSrc = sSrc & " xmlns:ns12=""http://www.tmax.co.kr/proobject/queryobject"" xmlns:ns13=""http://www.tmax.co.kr/proobject/bizobject"" "
sSrc = sSrc & " xmlns:ns14=""http://www.tmax.co.kr/proobject/serviceobject"" xmlns:ns15=""http://www.tmax.co.kr/proobject/dataobject"" "
sSrc = sSrc & " xmlns:ns16=""http://www.tmax.co.kr/proobject/taskobject"" xmlns:ns17=""http://www.tmax.co.kr/proobject/jobobject"" "
sSrc = sSrc & " xmlns:ns18=""http://www.tmax.co.kr/proobject/servicegroup"" xmlns:ns19=""http://www.tmax.co.kr/proobject/aspectj"" xmlns:ns20=""http://www.tmax.co.kr/proobject/mapping"" "
sSrc = sSrc & " xmlns:ns21=""http://www.tmax.co.kr/proobject/application/runtime"" xmlns:ns22=""http://www.tmaxsoft.co.kr/proobject/testOperation"" xmlns:ns23=""http://www.tmax.co.kr/proobject/restriction_codes"" xmlns:ns24=""http://www.example.org/externalObjectConfig"" "
sSrc = sSrc & " xmlns:ns25=""http://www.tmax.co.kr/proobject/application"" xmlns:ns26=""http://www.tmax.co.kr/proobject/contents"" "
sSrc = sSrc & " xmlns:ns27=""http://www.example.org/testData"" xmlns:ns28=""http://www.tmaxsoft.co.kr/proobject/testresult"" "
sSrc = sSrc & " xmlns:ns29=""http://www.tmax.co.kr/proobject/siteConfig"" xmlns:ns30=""http://www.tmax.co.kr/proobject/serverConfig"" "
sSrc = sSrc & " xmlns:ns31=""http://www.tmax.co.kr/proobject/probuilder_config"" xmlns:ns32=""http://www.tmax.co.kr/proobject/property"" "
sSrc = sSrc & " xmlns:ns33=""http://www.tmax.co.kr/proobject/serviceGroup"" identifier=""general"" "
sSrc = sSrc & " resourceId=""00006c97f18200ccdc629500f3a02c0b"" "
sSrc = sSrc & " logicalName=""" & s�����SVC�� & """ "
sSrc = sSrc & " physicalName=""Hgbg" & s�����Biz��� & "Biz"" "
sSrc = sSrc & " resourceType=""BIZ_OBJECT"" "
sSrc = sSrc & " resourcePackage=""hf.cor.hg.bg.biz.sbiz." & fService3() & """ "
sSrc = sSrc & " status=""DEV"" "
sSrc = sSrc & " creator=""ns042"" "
sSrc = sSrc & " owner=""ns042"" "
sSrc = sSrc & " revision=""1"" "
sSrc = sSrc & " description=""" & s�����SVC�� & " "
sSrc = sSrc & " wFrame"" resourcePath=""hf/cor/hg/bg/biz/sbiz/" & fService3() & """ "
sSrc = sSrc & " modifier=""ns042"" created=""2020-04-06 19:02:53"" lastUpdated=""2020-04-06 19:02:53"">"


sSrc = sSrc & vbCrLf & "    <ns13:bizMethod methodName=""sel" & s�����Biz��� & """ displayName=""" & s�����SVC�� & " ��ȸ"" modifier=""public"" comments=""@�޼���� : sel" & s�����Biz��� & "&#xD;&#xA;@���� : " & s�����SVC�� & " ��ȸ&#xD;&#xA;@�Է� : .Hgbg" & s�����Biz��� & "01InBvo&#xD;&#xA;@��� : Hgbg" & s�����Biz��� & "01OutBvo"">"
sSrc = sSrc & vbCrLf & "        <ns10:flowModule>"
sSrc = sSrc & vbCrLf & "            <ns10:entryModule>"
sSrc = sSrc & vbCrLf & "                <ns10:name>" & s�����SVC�� & " ��ȸ</ns10:name>"
sSrc = sSrc & vbCrLf & "                <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:id>0</ns10:id>"
sSrc = sSrc & vbCrLf & "                    <ns10:id>2</ns10:id>"
sSrc = sSrc & vbCrLf & "                    <ns10:id>4</ns10:id>"
sSrc = sSrc & vbCrLf & "                </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                <ns10:variable kind=""DTO"" name=""outBvo"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <ns10:do physicalName=""Hgbg" & s�����Biz��� & "01OutBvo"" logicalName=""" & s�����SVC�� & " ��ȸ OutBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"

sPname = "Hgbg" & s�����Biz��� & "01OutBvo"
sLname = s�����SVC�� & " ��ȸ OutBvo"
sPath = "hf/cor/hg/bg/vo/bvo/sbvo/" & fService3()
Call subMake�����VO


sSrc = sSrc & vbCrLf & "                </ns10:variable>"
sSrc = sSrc & vbCrLf & "            </ns10:entryModule>"
sSrc = sSrc & vbCrLf & "            <ns10:innerModules>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""block"" nodeId=""0"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>���� ���� ó��</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>���� ���� ó��</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>1</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition/>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""block"" nodeId=""2"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>���� �� ó��</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>���� �� ó��</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>9</ns10:id>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>8</ns10:id>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>3</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition/>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""block"" nodeId=""4"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>���� ���� ó��</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>���� ���� ó��</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>5</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition/>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""if"" nodeId=""6"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>��ȸ���� = ��ȸ</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>��ȸ������ ��ȸ�ϰ��</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>10</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition>"
sSrc = sSrc & vbCrLf & "                        <ns10:code>""1"".equals(inBvo.getRvewInqVo().getInqDivNo())</ns10:code>"
sSrc = sSrc & vbCrLf & "                    </ns10:condition>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""if"" nodeId=""7"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>��ȸ���� = �������</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>��ȸ������ ��������� ��� ��ȸ</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>11</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition/>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "            </ns10:innerModules>"
sSrc = sSrc & vbCrLf & "            <ns10:xorModules>"
sSrc = sSrc & vbCrLf & "                <ns10:xorModule nodeId=""8"">"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>6</ns10:id>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>7</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                </ns10:xorModule>"
sSrc = sSrc & vbCrLf & "            </ns10:xorModules>"
sSrc = sSrc & vbCrLf & "            <ns10:virtualModules>"
sSrc = sSrc & vbCrLf & "                <ns10:virtualModule nodeId=""1"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>�Է°� Ȯ��</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>�Է°� Ȯ��</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:code>                    logger.info(""#########################"" );&#xD;"
sSrc = sSrc & vbCrLf & "                   logger.info(""inBvo \n"" + inBvo.toString());&#xD;"
'sSrc = sSrc & vbCrLf & "                   logger.info(""#########################"" );    &#xD;"
'sSrc = sSrc & vbCrLf & "                    if(StringUtil.isEmpty(inBvo.getRvewInqVo().getCsno())){&#xD;"
'sSrc = sSrc & vbCrLf & "                        throw new Exception(""����ȣ"" + ""��(��) �ʼ��Է��Դϴ�."");                     &#xD;"
'sSrc = sSrc & vbCrLf & "                        //throw new UserException(""E00001"", e.getMessage());&#xD;"
'sSrc = sSrc & vbCrLf & "                    }&#xD;"
'sSrc = sSrc & vbCrLf & "                    if(StringUtil.isEmpty(inBvo.getRvewInqVo().getGrntNo())){&#xD;"
'sSrc = sSrc & vbCrLf & "                        throw new Exception(""������ȣ"" + ""��(��) �ʼ��Է��Դϴ�."");                     &#xD;"
'sSrc = sSrc & vbCrLf & "                        //throw new UserException(""E00001"", e.getMessage());&#xD;"
'sSrc = sSrc & vbCrLf & "                    }&#xD;"
'sSrc = sSrc & vbCrLf & "                    if(inBvo.getRvewInqVo().getRvewSeq() == 0){&#xD;"
'sSrc = sSrc & vbCrLf & "                        throw new Exception(""�ɻ����"" + ""��(��) �ʼ��Է��Դϴ�."");                     &#xD;"
'sSrc = sSrc & vbCrLf & "                        //throw new UserException(""E00001"", e.getMessage());&#xD;"
'sSrc = sSrc & vbCrLf & "                    } "
sSrc = sSrc & vbCrLf & "                </ns10:code>"
sSrc = sSrc & vbCrLf & "                </ns10:virtualModule>"
sSrc = sSrc & vbCrLf & "                <ns10:virtualModule nodeId=""3"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>" & s�����SVC�� & " ��ȸ</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>" & s�����SVC�� & " ��ȸ</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:code></ns10:code>"
sSrc = sSrc & vbCrLf & "                </ns10:virtualModule>"
sSrc = sSrc & vbCrLf & "                <ns10:virtualModule nodeId=""5"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>�������</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>�������</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:code>                     return outBvo;</ns10:code>"
sSrc = sSrc & vbCrLf & "                </ns10:virtualModule>"
sSrc = sSrc & vbCrLf & "                <ns10:virtualModule nodeId=""9"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>��������</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:code></ns10:code>"
sSrc = sSrc & vbCrLf & "                </ns10:virtualModule>"
sSrc = sSrc & vbCrLf & "                <ns10:virtualModule nodeId=""10"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>��ȸ</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>��ȸ</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:code>                         "
'sSrc = sSrc & vbCrLf & "                        TbHgj052mRvewetprbrf001Dvo hgj052mRvewetprbrf001Dvo = new TbHgj052mRvewetprbrf001Dvo();        // ���ýſ뺸��_�ɻ���ü����⺻&#xD;"
'sSrc = sSrc & vbCrLf & "                        TbHgj053mGrntrequnsell001Dvo hgj053mGrntrequnsell001Dvo = new TbHgj053mGrntrequnsell001Dvo();  // ���ýſ뺸��_������û����̺о�⺻&#xD;"
'sSrc = sSrc & vbCrLf & "                        TbHgj054mCnstprgsbzpl001Dvo hgj054mCnstprgsbzpl001Dvo = new TbHgj054mCnstprgsbzpl001Dvo();     // ���ýſ뺸��_�ð���������߻����⺻&#xD;"
'sSrc = sSrc & vbCrLf & "                                           &#xD;"
'sSrc = sSrc & vbCrLf & "                        // ���ýſ뺸��_�ɻ���ü����⺻ ��ȸ&#xD;"
'sSrc = sSrc & vbCrLf & "                        DataObjectMapper&lt;TbHgj052mRvewetprbrf001Dvo&gt; mapper = new DataObjectMapper&lt;TbHgj052mRvewetprbrf001Dvo&gt;( ""po7_app"", TbHgj052mRvewetprbrf001Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper.setQuery(TbHgj052mRvewetprbrf001Dao.SEL001);&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper.setParameter(""csno"", inBvo.getRvewInqVo().getCsno());         // ����ȣ                  &#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper.setParameter(""rvewSeq"", inBvo.getRvewInqVo().getRvewSeq());       // �ɻ����&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper.setParameter(""rqerDvcd"", ""01"");         // ��û�α����ڵ�(01:��û��)&#xD;"
'sSrc = sSrc & vbCrLf & "                        &#xD;"
'sSrc = sSrc & vbCrLf & "                        hgj052mRvewetprbrf001Dvo = mapper.get();       // ������û��� ��ü���� ��ȸ&#xD;"
'sSrc = sSrc & vbCrLf & "                        outBvo.setBzepBrfVo(hgj052mRvewetprbrf001Dvo);&#xD;"
'sSrc = sSrc & vbCrLf & "                        &#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper.setParameter(""rqerDvcd"", ""02"");         // ��û�α����ڵ�(02:�ð���)&#xD;"
'sSrc = sSrc & vbCrLf & "                        hgj052mRvewetprbrf001Dvo = mapper.get();       // �ð��� ��ü���� ��ȸ&#xD;"
'sSrc = sSrc & vbCrLf & "                        outBvo.setCncoBzepBrfVo(hgj052mRvewetprbrf001Dvo);&#xD;"
'sSrc = sSrc & vbCrLf & "                        &#xD;"
'sSrc = sSrc & vbCrLf & "                         &#xD;"
'sSrc = sSrc & vbCrLf & "                        // ���ýſ뺸��_������û����̺о�⺻ ��ȸ &#xD;"
'sSrc = sSrc & vbCrLf & "                        // ������û��� �̺о�� ��ȸ&#xD;"
'sSrc = sSrc & vbCrLf & "                        DataObjectMapper&lt;TbHgj053mGrntrequnsell001Dvo&gt; mapper2 = new DataObjectMapper&lt;TbHgj053mGrntrequnsell001Dvo&gt;( ""po7_app"", TbHgj053mGrntrequnsell001Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper2.setQuery(TbHgj053mGrntrequnsell001Dao.SELLIST001);&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper2.setParameter(""csno"", inBvo.getRvewInqVo().getCsno());            // ����ȣ                  &#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper2.setParameter(""rvewSeq"", inBvo.getRvewInqVo().getRvewSeq());          // �ɻ����                  &#xD;"
'sSrc = sSrc & vbCrLf & "                        outBvo.setUnSellVoList(mapper2.getForwardList());&#xD;"
'sSrc = sSrc & vbCrLf & "                            &#xD;"
'sSrc = sSrc & vbCrLf & "                        // ���ýſ뺸��_�ð���������߻����⺻ ��ȸ&#xD;"
'sSrc = sSrc & vbCrLf & "                        // �ð���� �������� ����� ��ȸ&#xD;"
'sSrc = sSrc & vbCrLf & "                        DataObjectMapper&lt;TbHgj054mCnstprgsbzpl001Dvo&gt; mapper3 = new DataObjectMapper&lt;TbHgj054mCnstprgsbzpl001Dvo&gt;( ""po7_app"", TbHgj054mCnstprgsbzpl001Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper3.setQuery(TbHgj054mCnstprgsbzpl001Dao.SELLIST001);&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper3.setParameter(""csno"", inBvo.getRvewInqVo().getCsno());            // ����ȣ                  &#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper3.setParameter(""rvewSeq"", inBvo.getRvewInqVo().getRvewSeq());          // �ɻ����                  &#xD;"
'sSrc = sSrc & vbCrLf & "                        outBvo.setCnstPrgsBzplVoList(mapper3.getForwardList());    &#xD;"
'sSrc = sSrc & vbCrLf & "                        &#xD;"
'sSrc = sSrc & vbCrLf & "                        // ������ȣ��  �ش��ϴ� �ð��� ����ȣ ��ȸ&#xD;"
'sSrc = sSrc & vbCrLf & "                        Hgbg" & s�����Biz��� & "001Dvo " & s�����Biz��� & "001Dvo = new Hgbg" & s�����Biz��� & "001Dvo();&#xD;"
'sSrc = sSrc & vbCrLf & "                        DataObjectMapper&lt;Hgbg" & s�����Biz��� & "001Dvo&gt; mapper4 = new DataObjectMapper&lt;Hgbg" & s�����Biz��� & "001Dvo&gt;( ""po7_app"", Hgbg" & s�����Biz��� & "001Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper4.setQuery(Hgbg" & s�����Biz��� & "001Dao.SEL001);&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper4.setParameter(""grntNo"", inBvo.getRvewInqVo().getGrntNo());            // ������ȣ                                  &#xD;"
'sSrc = sSrc & vbCrLf & "                        " & s�����Biz��� & "001Dvo = mapper4.get(); &#xD;"
'sSrc = sSrc & vbCrLf & "                        outBvo.setCncoVo(" & s�����Biz��� & "001Dvo);  &#xD;"
'sSrc = sSrc & vbCrLf & "                             &#xD;"
'sSrc = sSrc & vbCrLf & "                        // ������û��� ��ü���� ������ ���� ���&#xD;"
'sSrc = sSrc & vbCrLf & "                        if(hgj052mRvewetprbrf001Dvo == null){&#xD;"
'sSrc = sSrc & vbCrLf & "                            // ������� ��ȸ&#xD;"
'sSrc = sSrc & vbCrLf & "                            " & s�����Biz��� & "001Dvo = new Hgbg" & s�����Biz��� & "001Dvo();&#xD;"
'sSrc = sSrc & vbCrLf & "                            DataObjectMapper&lt;Hgbg" & s�����Biz��� & "001Dvo&gt; mapper5 = new DataObjectMapper&lt;Hgbg" & s�����Biz��� & "001Dvo&gt;( ""po7_app"", Hgbg" & s�����Biz��� & "001Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                            mapper5.setQuery(Hgbg" & s�����Biz��� & "001Dao.SEL002);&#xD;"
'sSrc = sSrc & vbCrLf & "                            mapper5.setParameter(""grntNo"", inBvo.getRvewInqVo().getGrntNo());            // ������ȣ                                  &#xD;"
'sSrc = sSrc & vbCrLf & "                            " & s�����Biz��� & "001Dvo = mapper5.get();&#xD;"
'sSrc = sSrc & vbCrLf & "                            &#xD;"
'sSrc = sSrc & vbCrLf & "                            // ������� ��ȸ����� ���� ���&#xD;"
'sSrc = sSrc & vbCrLf & "                            if(" & s�����Biz��� & "001Dvo != null){&#xD;"
'sSrc = sSrc & vbCrLf & "                                String strEtno = """";         // ���ü��ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                                String strFinKncd = """";      // �繫�����ڵ�&#xD;"
'sSrc = sSrc & vbCrLf & "                                String strFinFormCd = """";        // �繫�����ڵ�&#xD;"
'sSrc = sSrc & vbCrLf & "                                String strRschSbjtCd = """";       // ��������ڵ�&#xD;"
'sSrc = sSrc & vbCrLf & "                                &#xD;"
'sSrc = sSrc & vbCrLf & "                                strEtno = " & s�����Biz��� & "001Dvo.getEtno();    // ���ü��ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                                strFinKncd = ""01"";                               // �繫�����ڵ�(01: ��������ǥ)&#xD;"
'sSrc = sSrc & vbCrLf & "                                strRschSbjtCd = ""100000"";                        // �繫�����ڵ�(100000: �ڻ��Ѱ�)&#xD;"
'sSrc = sSrc & vbCrLf & "                                &#xD;"
'sSrc = sSrc & vbCrLf & "                            }&#xD;"
'sSrc = sSrc & vbCrLf & "                                                        &#xD;"
'sSrc = sSrc & vbCrLf & "                        }          "
sSrc = sSrc & vbCrLf & "                </ns10:code>"
sSrc = sSrc & vbCrLf & "                </ns10:virtualModule>"
sSrc = sSrc & vbCrLf & "                <ns10:virtualModule nodeId=""11"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>������� ��ȸ</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>������� ��ȸ</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:code>                         "
'sSrc = sSrc & vbCrLf & "                        TbHgj052mRvewetprbrf001Dvo hgj052mRvewetprbrf001Dvo = new TbHgj052mRvewetprbrf001Dvo();        // ���ýſ뺸��_�ɻ���ü����⺻&#xD;"
'sSrc = sSrc & vbCrLf & "                        TbHgj053mGrntrequnsell001Dvo hgj053mGrntrequnsell001Dvo = new TbHgj053mGrntrequnsell001Dvo();  // ���ýſ뺸��_������û����̺о�⺻&#xD;"
'sSrc = sSrc & vbCrLf & "                        TbHgj054mCnstprgsbzpl001Dvo hgj054mCnstprgsbzpl001Dvo = new TbHgj054mCnstprgsbzpl001Dvo();     // ���ýſ뺸��_�ð���������߻����⺻&#xD;"
'sSrc = sSrc & vbCrLf & "                                           &#xD;"
'sSrc = sSrc & vbCrLf & "                        // ���ýſ뺸��_�ɻ���ü����⺻ ��ȸ&#xD;"
'sSrc = sSrc & vbCrLf & "                        DataObjectMapper&lt;TbHgj052mRvewetprbrf001Dvo&gt; mapper = new DataObjectMapper&lt;TbHgj052mRvewetprbrf001Dvo&gt;( ""po7_app"", TbHgj052mRvewetprbrf001Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper.setQuery(TbHgj052mRvewetprbrf001Dao.SEL001);&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper.setParameter(""csno"", inBvo.getRvewInqVo().getCsno());         // ����ȣ                  &#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper.setParameter(""rvewSeq"", inBvo.getRvewInqVo().getRvewSeq());       // �ɻ����&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper.setParameter(""rqerDvcd"", ""01"");         // ��û�α����ڵ�(01:��û��)&#xD;"
'sSrc = sSrc & vbCrLf & "                        &#xD;"
'sSrc = sSrc & vbCrLf & "                        hgj052mRvewetprbrf001Dvo = mapper.get();       // ������û��� ��ü���� ��ȸ&#xD;"
'sSrc = sSrc & vbCrLf & "                        outBvo.setBzepBrfVo(hgj052mRvewetprbrf001Dvo);&#xD;"
'sSrc = sSrc & vbCrLf & "                        &#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper.setParameter(""rqerDvcd"", ""02"");         // ��û�α����ڵ�(02:�ð���)&#xD;"
'sSrc = sSrc & vbCrLf & "                        hgj052mRvewetprbrf001Dvo = mapper.get();       // �ð��� ��ü���� ��ȸ&#xD;"
'sSrc = sSrc & vbCrLf & "                        outBvo.setCncoBzepBrfVo(hgj052mRvewetprbrf001Dvo);&#xD;"
'sSrc = sSrc & vbCrLf & "                        &#xD;"
'sSrc = sSrc & vbCrLf & "                         &#xD;"
'sSrc = sSrc & vbCrLf & "                        // ���ýſ뺸��_������û����̺о�⺻ ��ȸ &#xD;"
'sSrc = sSrc & vbCrLf & "                        // ������û��� �̺о�� ��ȸ&#xD;"
'sSrc = sSrc & vbCrLf & "                        DataObjectMapper&lt;TbHgj053mGrntrequnsell001Dvo&gt; mapper2 = new DataObjectMapper&lt;TbHgj053mGrntrequnsell001Dvo&gt;( ""po7_app"", TbHgj053mGrntrequnsell001Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper2.setQuery(TbHgj053mGrntrequnsell001Dao.SELLIST001);&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper2.setParameter(""csno"", inBvo.getRvewInqVo().getCsno());            // ����ȣ                  &#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper2.setParameter(""rvewSeq"", inBvo.getRvewInqVo().getRvewSeq());          // �ɻ����                  &#xD;"
'sSrc = sSrc & vbCrLf & "                        outBvo.setUnSellVoList(mapper2.getForwardList());&#xD;"
'sSrc = sSrc & vbCrLf & "                            &#xD;"
'sSrc = sSrc & vbCrLf & "                        // ���ýſ뺸��_�ð���������߻����⺻ ��ȸ&#xD;"
'sSrc = sSrc & vbCrLf & "                        // �ð���� �������� ����� ��ȸ&#xD;"
'sSrc = sSrc & vbCrLf & "                        DataObjectMapper&lt;TbHgj054mCnstprgsbzpl001Dvo&gt; mapper3 = new DataObjectMapper&lt;TbHgj054mCnstprgsbzpl001Dvo&gt;( ""po7_app"", TbHgj054mCnstprgsbzpl001Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper3.setQuery(TbHgj054mCnstprgsbzpl001Dao.SELLIST001);&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper3.setParameter(""csno"", inBvo.getRvewInqVo().getCsno());            // ����ȣ                  &#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper3.setParameter(""rvewSeq"", inBvo.getRvewInqVo().getRvewSeq());          // �ɻ����                  &#xD;"
'sSrc = sSrc & vbCrLf & "                        outBvo.setCnstPrgsBzplVoList(mapper3.getForwardList());    &#xD;"
'sSrc = sSrc & vbCrLf & "                        &#xD;"
'sSrc = sSrc & vbCrLf & "                        // ������ȣ��  �ش��ϴ� �ð��� ����ȣ ��ȸ&#xD;"
'sSrc = sSrc & vbCrLf & "                        Hgbg" & s�����Biz��� & "001Dvo " & s�����Biz��� & "001Dvo = new Hgbg" & s�����Biz��� & "001Dvo();&#xD;"
'sSrc = sSrc & vbCrLf & "                        DataObjectMapper&lt;Hgbg" & s�����Biz��� & "001Dvo&gt; mapper4 = new DataObjectMapper&lt;Hgbg" & s�����Biz��� & "001Dvo&gt;( ""po7_app"", Hgbg" & s�����Biz��� & "001Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper4.setQuery(Hgbg" & s�����Biz��� & "001Dao.SEL001);&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper4.setParameter(""grntNo"", inBvo.getRvewInqVo().getGrntNo());            // ������ȣ                                  &#xD;"
'sSrc = sSrc & vbCrLf & "                        " & s�����Biz��� & "001Dvo = mapper4.get(); &#xD;"
'sSrc = sSrc & vbCrLf & "                        outBvo.setCncoVo(" & s�����Biz��� & "001Dvo);  &#xD;"
'sSrc = sSrc & vbCrLf & "                             &#xD;"
'sSrc = sSrc & vbCrLf & "                        // ������û��� ��ü���� ������ ���� ���&#xD;"
'sSrc = sSrc & vbCrLf & "                        if(hgj052mRvewetprbrf001Dvo == null){&#xD;"
'sSrc = sSrc & vbCrLf & "                            // ������� ��ȸ&#xD;"
'sSrc = sSrc & vbCrLf & "                            " & s�����Biz��� & "001Dvo = new Hgbg" & s�����Biz��� & "001Dvo();&#xD;"
'sSrc = sSrc & vbCrLf & "                            DataObjectMapper&lt;Hgbg" & s�����Biz��� & "001Dvo&gt; mapper5 = new DataObjectMapper&lt;Hgbg" & s�����Biz��� & "001Dvo&gt;( ""po7_app"", Hgbg" & s�����Biz��� & "001Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                            mapper5.setQuery(Hgbg" & s�����Biz��� & "001Dao.SEL002);&#xD;"
'sSrc = sSrc & vbCrLf & "                            mapper5.setParameter(""grntNo"", inBvo.getRvewInqVo().getGrntNo());            // ������ȣ                                  &#xD;"
'sSrc = sSrc & vbCrLf & "                            " & s�����Biz��� & "001Dvo = mapper5.get();&#xD;"
'sSrc = sSrc & vbCrLf & "                            &#xD;"
'sSrc = sSrc & vbCrLf & "                            // ������� ��ȸ����� ���� ���&#xD;"
'sSrc = sSrc & vbCrLf & "                            if(" & s�����Biz��� & "001Dvo != null){&#xD;"
'sSrc = sSrc & vbCrLf & "                                String strEtno = " & s�����Biz��� & "001Dvo.getEtno();                 // ���ü��ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                                String strRschBasisDy = " & s�����Biz��� & "001Dvo.getRschBasisDy();   // �����������&#xD;"
'sSrc = sSrc & vbCrLf & "                                &#xD;"
'sSrc = sSrc & vbCrLf & "                                String strFinKncd = """";      // �繫�����ڵ�&#xD;"
'sSrc = sSrc & vbCrLf & "                                String strFinFormCd = """";        // �繫�����ڵ�&#xD;"
'sSrc = sSrc & vbCrLf & "                                String strRschSbjtCd = """";       // ��������ڵ�&#xD;"
'sSrc = sSrc & vbCrLf & "                                &#xD;"
'sSrc = sSrc & vbCrLf & "                                strFinKncd = ""01"";                               // �繫�����ڵ�(01: ��������ǥ)&#xD;"
'sSrc = sSrc & vbCrLf & "                                strRschSbjtCd = ""100000"";                        // �繫�����ڵ�(100000: �ڻ��Ѱ�)&#xD;"
'sSrc = sSrc & vbCrLf & "                                &#xD;"
'sSrc = sSrc & vbCrLf & "                                // ��������Ͽ� ���� �繫�����ڵ� ����&#xD;"
'sSrc = sSrc & vbCrLf & "                                if(DateUtil.getDaysDiff(strRschBasisDy, ""20080503"") &lt; 0){&#xD;"
'sSrc = sSrc & vbCrLf & "                                    strFinFormCd = ""01""; // �繫�����ڵ�(01: 2008�� ���� ���)&#xD;"
'sSrc = sSrc & vbCrLf & "                                }&#xD;"
'sSrc = sSrc & vbCrLf & "                                else if(DateUtil.getDaysDiff(strRschBasisDy, ""20170109"") &lt; 0){&#xD;"
'sSrc = sSrc & vbCrLf & "                                    strFinFormCd = ""02""; // �繫�����ڵ�(01: 2008�� ���� ���)&#xD;"
'sSrc = sSrc & vbCrLf & "                                }&#xD;"
'sSrc = sSrc & vbCrLf & "                                else if(DateUtil.getDaysDiff(strRschBasisDy, ""20181026"") &lt; 0){&#xD;"
'sSrc = sSrc & vbCrLf & "                                    strFinFormCd = ""03""; // �繫�����ڵ�(01: 2017�� ���� ���)&#xD;"
'sSrc = sSrc & vbCrLf & "                                }&#xD;"
'sSrc = sSrc & vbCrLf & "                                else{&#xD;"
'sSrc = sSrc & vbCrLf & "                                    strFinFormCd = ""04""; // �繫�����ڵ�(04: 2018�� ���� ���)&#xD;"
'sSrc = sSrc & vbCrLf & "                                }                                                               &#xD;"
'sSrc = sSrc & vbCrLf & "                                &#xD;"
'sSrc = sSrc & vbCrLf & "                                // ���� ���ڻ�&#xD;"
'sSrc = sSrc & vbCrLf & "                            }&#xD;"
'sSrc = sSrc & vbCrLf & "                                                        &#xD;"
'sSrc = sSrc & vbCrLf & "                        }          "
sSrc = sSrc & vbCrLf & "                 </ns10:code>"
sSrc = sSrc & vbCrLf & "                </ns10:virtualModule>"
sSrc = sSrc & vbCrLf & "            </ns10:virtualModules>"
sSrc = sSrc & vbCrLf & "            <ns10:loopModules/>"
sSrc = sSrc & vbCrLf & "            <ns10:bizMethodCalls/>"
sSrc = sSrc & vbCrLf & "            <ns10:serviceCalls/>"
sSrc = sSrc & vbCrLf & "            <ns10:assignModules/>"
sSrc = sSrc & vbCrLf & "            <ns10:dataObjectCalls/>"
sSrc = sSrc & vbCrLf & "            <ns10:replyCalls/>"
sSrc = sSrc & vbCrLf & "        </ns10:flowModule>"
sSrc = sSrc & vbCrLf & "        <ns13:arguments>"
sSrc = sSrc & vbCrLf & "            <ns10:argument>"
sSrc = sSrc & vbCrLf & "                <ns10:variable kind=""DTO"" name=""inBvo"" comments=""0'th argument"">"
sSrc = sSrc & vbCrLf & "                    <ns10:do physicalName=""Hgbg" & s�����Biz��� & "01InBvo"" logicalName=""" & s�����SVC�� & " ��ȸ InBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"

sPname = "Hgbg" & s�����Biz��� & "01InBvo"
sLname = s�����SVC�� & " ��ȸ InBvo"
sPath = "hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & ""
Call subMake�����VO

sSrc = sSrc & vbCrLf & "                </ns10:variable>"
sSrc = sSrc & vbCrLf & "                <ns10:sequence>0</ns10:sequence>"
sSrc = sSrc & vbCrLf & "            </ns10:argument>"
sSrc = sSrc & vbCrLf & "        </ns13:arguments>"
sSrc = sSrc & vbCrLf & "        <ns13:returnType kind=""DTO"">"
sSrc = sSrc & vbCrLf & "            <ns10:do physicalName=""Hgbg" & s�����Biz��� & "01OutBvo"" logicalName=""" & s�����SVC�� & " ��ȸ OutBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"
sSrc = sSrc & vbCrLf & "        </ns13:returnType>"
sSrc = sSrc & vbCrLf & "        <ns13:classInfo classPackageName=""hf/cor/hg/bg/biz/sbiz/" & fService3() & """ className=""Hgbg" & s�����Biz��� & "Biz""/>"
sSrc = sSrc & vbCrLf & "        <ns13:annotations/>"
sSrc = sSrc & vbCrLf & "    </ns13:bizMethod>"
sSrc = sSrc & vbCrLf & "    <ns13:bizMethod methodName=""prcs" & s�����Biz��� & """ displayName=""" & s�����SVC�� & " ó��"" modifier=""public"" comments=""@�޼���� : prcs" & s�����Biz��� & "&#xD;&#xA;@���� : " & s�����SVC�� & " ó��&#xD;&#xA;@�Է� : .Hgbg" & s�����Biz��� & "02InBvo&#xD;&#xA;@��� : Hgbg" & s�����Biz��� & "02OutBvo"">"
sSrc = sSrc & vbCrLf & "        <ns10:flowModule>"
sSrc = sSrc & vbCrLf & "            <ns10:entryModule>"
sSrc = sSrc & vbCrLf & "                <ns10:name>" & s�����SVC�� & " ó��</ns10:name>"
sSrc = sSrc & vbCrLf & "                <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:id>0</ns10:id>"
sSrc = sSrc & vbCrLf & "                    <ns10:id>2</ns10:id>"
sSrc = sSrc & vbCrLf & "                    <ns10:id>4</ns10:id>"
sSrc = sSrc & vbCrLf & "                </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                <ns10:variable kind=""DTO"" name=""outBvo"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <ns10:do physicalName=""Hgbg" & s�����Biz��� & "02OutBvo"" logicalName=""" & s�����SVC�� & " ó�� OutBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"

sPname = "Hgbg" & s�����Biz��� & "02OutBvo"
sLname = s�����SVC�� & " ó�� OutBvo"
sPath = "hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & ""
Call subMake�����VO

sSrc = sSrc & vbCrLf & "                </ns10:variable>"
sSrc = sSrc & vbCrLf & "            </ns10:entryModule>"
sSrc = sSrc & vbCrLf & "            <ns10:innerModules>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""block"" nodeId=""0"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>���� ���� ó��</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>���� ���� ó��</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>1</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition/>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""block"" nodeId=""2"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>���� �� ó��</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>���� �� ó��</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>3</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition/>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""block"" nodeId=""4"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>���� ���� ó��</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>���� ���� ó��</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>5</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition/>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "            </ns10:innerModules>"
sSrc = sSrc & vbCrLf & "            <ns10:xorModules/>"
sSrc = sSrc & vbCrLf & "            <ns10:virtualModules>"
sSrc = sSrc & vbCrLf & "                <ns10:virtualModule nodeId=""1"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>�Է°� Ȯ��</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>�Է°� Ȯ��</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:code>                    logger.info(""#########################"" );&#xD;"
sSrc = sSrc & vbCrLf & "                   logger.info(""inBvo \n"" + inBvo.toString());&#xD;"
'sSrc = sSrc & vbCrLf & "                   logger.info(""#########################"" );    &#xD;"
'sSrc = sSrc & vbCrLf & "//                  if(StringUtil.isEmpty(inBvo.getRvewInqVo().getCsno())){&#xD;"
'sSrc = sSrc & vbCrLf & "//                      throw new Exception(""����ȣ"" + ""��(��) �ʼ��Է��Դϴ�."");                     &#xD;"
'sSrc = sSrc & vbCrLf & "//                      //throw new UserException(""E00001"", e.getMessage());&#xD;"
'sSrc = sSrc & vbCrLf & "//                  }&#xD;"
'sSrc = sSrc & vbCrLf & "//                  if(StringUtil.isEmpty(inBvo.getRvewInqVo().getGrntNo())){&#xD;"
'sSrc = sSrc & vbCrLf & "//                      throw new Exception(""������ȣ"" + ""��(��) �ʼ��Է��Դϴ�."");                     &#xD;"
'sSrc = sSrc & vbCrLf & "//                      //throw new UserException(""E00001"", e.getMessage());&#xD;"
'sSrc = sSrc & vbCrLf & "//                  }&#xD;"
'sSrc = sSrc & vbCrLf & "//                  if(inBvo.getRvewInqVo().getRvewSeq() == 0){&#xD;"
'sSrc = sSrc & vbCrLf & "//                      throw new Exception(""�ɻ����"" + ""��(��) �ʼ��Է��Դϴ�."");                     &#xD;"
'sSrc = sSrc & vbCrLf & "//                      //throw new UserException(""E00001"", e.getMessage());&#xD;"
'sSrc = sSrc & vbCrLf & "//                  }"
sSrc = sSrc & vbCrLf & "                </ns10:code>"
sSrc = sSrc & vbCrLf & "                </ns10:virtualModule>"
sSrc = sSrc & vbCrLf & "                <ns10:virtualModule nodeId=""3"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>" & s�����SVC�� & " ó��</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>" & s�����SVC�� & " ó��</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:code></ns10:code>"
sSrc = sSrc & vbCrLf & "                </ns10:virtualModule>"
sSrc = sSrc & vbCrLf & "                <ns10:virtualModule nodeId=""5"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>�������</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>�������</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:code>                     return outBvo;</ns10:code>"
sSrc = sSrc & vbCrLf & "                </ns10:virtualModule>"
sSrc = sSrc & vbCrLf & "            </ns10:virtualModules>"
sSrc = sSrc & vbCrLf & "            <ns10:loopModules/>"
sSrc = sSrc & vbCrLf & "            <ns10:bizMethodCalls/>"
sSrc = sSrc & vbCrLf & "            <ns10:serviceCalls/>"
sSrc = sSrc & vbCrLf & "            <ns10:assignModules/>"
sSrc = sSrc & vbCrLf & "            <ns10:dataObjectCalls/>"
sSrc = sSrc & vbCrLf & "            <ns10:replyCalls/>"
sSrc = sSrc & vbCrLf & "        </ns10:flowModule>"
sSrc = sSrc & vbCrLf & "        <ns13:arguments>"
sSrc = sSrc & vbCrLf & "            <ns10:argument>"
sSrc = sSrc & vbCrLf & "                <ns10:variable kind=""DTO"" name=""inBvo"" comments=""0'th argument"">"
sSrc = sSrc & vbCrLf & "                    <ns10:do physicalName=""Hgbg" & s�����Biz��� & "02InBvo"" logicalName=""" & s�����SVC�� & " ó�� InBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"

sPname = "Hgbg" & s�����Biz��� & "02InBvo"
sLname = s�����SVC�� & " ó�� InBvo"
sPath = "hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & ""
Call subMake�����VO

sSrc = sSrc & vbCrLf & "                </ns10:variable>"
sSrc = sSrc & vbCrLf & "                <ns10:sequence>0</ns10:sequence>"
sSrc = sSrc & vbCrLf & "            </ns10:argument>"
sSrc = sSrc & vbCrLf & "        </ns13:arguments>"
sSrc = sSrc & vbCrLf & "        <ns13:returnType kind=""DTO"">"
sSrc = sSrc & vbCrLf & "            <ns10:do physicalName=""Hgbg" & s�����Biz��� & "02OutBvo"" logicalName=""" & s�����SVC�� & " ó�� OutBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"

sPname = "Hgbg" & s�����Biz��� & "02OutBvo"
sLname = s�����SVC�� & " ó�� OutBvo"
sPath = "hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & ""
Call subMake�����VO

sSrc = sSrc & vbCrLf & "        </ns13:returnType>"
sSrc = sSrc & vbCrLf & "        <ns13:classInfo classPackageName=""hf/cor/hg/bg/biz/sbiz/" & fService3() & """ className=""Hgbg" & s�����Biz��� & "Biz""/>"
sSrc = sSrc & vbCrLf & "        <ns13:annotations/>"
sSrc = sSrc & vbCrLf & "    </ns13:bizMethod>"
sSrc = sSrc & vbCrLf & "    <ns13:member/>"
sSrc = sSrc & vbCrLf & "    <ns13:interfaces/>"
sSrc = sSrc & vbCrLf & "    <ns13:wholeImport>"
sSrc = sSrc & vbCrLf & "        <ns13:importType>com.tmax.proobject.core.BizObject</ns13:importType>"
sSrc = sSrc & vbCrLf & "        <ns13:importType>com.tmax.proobject.dataobject.mapper.DataObjectMapper</ns13:importType>"
sSrc = sSrc & vbCrLf & "        <ns13:importType>com.tmax.proobject.logger.ProObjectLogger</ns13:importType>"
sSrc = sSrc & vbCrLf & "        <ns13:importType>com.tmax.proobject.logger.application.ServiceLogger</ns13:importType>"
sSrc = sSrc & vbCrLf & "        <ns13:importType>com.tmax.proobject.model.business.BusinessObject</ns13:importType>"
sSrc = sSrc & vbCrLf & "        <ns13:importType>hf.cor.co.common.util.DateUtil</ns13:importType>"
sSrc = sSrc & vbCrLf & "        <ns13:importType>hf.cor.co.common.util.StringUtil</ns13:importType>"
sSrc = sSrc & vbCrLf & "        <ns13:importType>hf.cor.hg.bg.dao.sdao." & fService3() & ".Hgbg" & s�����Biz��� & "001Dao</ns13:importType>"


sPname = "Hgbg" & s�����Biz��� & "001Dao"
sLname = s�����SVC�� & "001Dao"
sPath = Replace("hf.cor.hg.bg.dao.sdao." & fService3(), ".", "/")
Call subMake�����DAO


sPname = "Hgbg" & s�����Biz��� & "001Dvo"
sLname = s�����SVC�� & " ó�� InBvo"
sPath = "hg/bg/vo/dvo/sdvo/" & fService3() & ""
Call subMake�����VO


sSrc = sSrc & vbCrLf & "        <ns13:importType>hf.cor.hg.bg.dao.tdao.j.TbHgj052mRvewetprbrf001Dao</ns13:importType>"
sSrc = sSrc & vbCrLf & "        <ns13:importType>hf.cor.hg.bg.dao.tdao.j.TbHgj053mGrntrequnsell001Dao</ns13:importType>"
sSrc = sSrc & vbCrLf & "        <ns13:importType>hf.cor.hg.bg.dao.tdao.j.TbHgj054mCnstprgsbzpl001Dao</ns13:importType>"
sSrc = sSrc & vbCrLf & "        <ns13:importType>hf.cor.hg.bg.vo.dvo.sdvo." & fService3() & ".Hgbg" & s�����Biz��� & "001Dvo</ns13:importType>"
sSrc = sSrc & vbCrLf & "        <ns13:importType>hf.cor.hg.bg.vo.dvo.tdvo.j.TbHgj052mRvewetprbrf001Dvo</ns13:importType>"
sSrc = sSrc & vbCrLf & "        <ns13:importType>hf.cor.hg.bg.vo.dvo.tdvo.j.TbHgj053mGrntrequnsell001Dvo</ns13:importType>"
sSrc = sSrc & vbCrLf & "        <ns13:importType>hf.cor.hg.bg.vo.dvo.tdvo.j.TbHgj054mCnstprgsbzpl001Dvo</ns13:importType>"
sSrc = sSrc & vbCrLf & "        <ns13:importType>hf2.cor.fw.sys.util.UserException</ns13:importType>"
sSrc = sSrc & vbCrLf & "    </ns13:wholeImport>"
sSrc = sSrc & vbCrLf & "    <ns13:typeParameters/>"
sSrc = sSrc & vbCrLf & "</ns13:bizObject>"


With fpBG
                                                                                                                                                                                                                
End With
sMake�����biz = sSrc
End Function



'����� ���� �б�
Private Sub cmdBGExcel_Click()
    
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
    x = fpBG.IsExcelFile("c:\src\�����.xls")

    ' If file is Excel file, tell user, import sheet
    ' list, and set result to y
    If x = 1 Then
        'MsgBox "File is an Excel file.", , "File Type"
        y = fpBG.GetExcelSheetList("c:\src\�����.xls", List, listcount, "C:\ILOGFILE_SVC.TXT", handle, True)
        ' If received sheet list, tell user, import file,
        ' and set result to z
        If y = True Then
            'MsgBox "Got sheet list.", , "Status"
            z = fpBG.ImportExcelSheet(handle, 0)
            ' Tell user result based on T/F value of z
            If z = True Then
                MsgBox "Import complete.", , "Result"
            Else
                MsgBox "Import did not succeed.", , "Result"
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

'" & sClassPName & "
'" & sPackageSlashName & "
'
'" & s�޼ҵ�id & "
'" & s�޼ҵ�� & "
'" & sINBVO & "
'" & sOUTBVO & "
'
'" & sVO������ & "
'" & sVOid & "
'" & sVO�� & "
'" & sVO��� & "
'
'" & right(s�޼ҵ�id,4) & "
'
'iROW_CNT = iROW_CNT + 1: subSetMehod
'iROW_CNT = iROW_CNT + 1: subSetVO
Private Sub cmdExcel_Click()
    
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
    x = fpSrcSVC.IsExcelFile("c:\src\src.xls")

    ' If file is Excel file, tell user, import sheet
    ' list, and set result to y
    If x = 1 Then
        'MsgBox "File is an Excel file.", , "File Type"
        y = fpSrc.GetExcelSheetList("c:\src\src.xls", List, listcount, "C:\ILOGFILE.TXT", handle, True)
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

Private Sub cmdExcelSVC_Click()
    
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
    x = fpSrc.IsExcelFile("c:\src\src_SVC.xls")

    ' If file is Excel file, tell user, import sheet
    ' list, and set result to y
    If x = 1 Then
        'MsgBox "File is an Excel file.", , "File Type"
        y = fpSrc.GetExcelSheetList("c:\src\src_SVC.xls", List, listcount, "C:\ILOGFILE_SVC.TXT", handle, True)
        ' If received sheet list, tell user, import file,
        ' and set result to z
        If y = True Then
            'MsgBox "Got sheet list.", , "Status"
            z = fpSrcSVC.ImportExcelSheet(handle, 0)
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

'----------------------------------------
'�ҽ� ���� ��ư Ŭ��
'----------------------------------------
Private Sub cmdMakeSrc_Click()
On Error GoTo Err_Handler
With fpSrc
    Dim stmFileControl As Stream
    Dim sPrintString As String
    
    
        
    iROW_CNT = 1
    
    
    Dim sFilename As String
    iROW_CNT = iROW_CNT + 1
        
LOOP_START:
    
    
    
    Set stmFileControl = New Stream
    
    
    stmFileControl.Open
    stmFileControl.Position = 0
    stmFileControl.Charset = "UTF-8"
    
    
    'file name ����
    .Row = iROW_CNT: .Col = 2
    sFilename = .Value & ".bo"
    
    .Col = 1: sClassLname = .Value
    .Col = 2: sClassPname = .Value
    .Col = 3: sPackageDotName = .Value
    .Col = 4: sUserId = .Value
    .Col = 5: sUserName = .Value
    .Col = 6: sRevision = .Value
    .Col = 7: sPackageSlashName = .Value
    
    
    
    '����� ���ڿ� ����
    sPrintString = sMakeSrcString
    
    
    'stmFileControl.WriteText (sPrintString)
    'stmFileControl.SaveToFile "c:\src\" & sFileName, adSaveCreateOverWrite
    'stmFileControl.Close
            
    'Set stmFileControl = Nothing
                    
                    
    Dim objStreamUTF8 As Stream
    Set objStreamUTF8 = New Stream
    
    
    Dim objStreamUTF8NoBOM As Stream
    Set objStreamUTF8NoBOM = New Stream
    
    
    With objStreamUTF8
        .Charset = "UTF-8"
        .Open
        .WriteText sPrintString
        .Position = 0
        .Type = adTypeText
        .Position = 3
    End With
    
    With objStreamUTF8NoBOM
        .Type = adTypeBinary
        .Open
        objStreamUTF8.CopyTo objStreamUTF8NoBOM
        .SaveToFile "c:\src\" & sFilename, adSaveCreateOverWrite
        
    End With
    
    objStreamUTF8.Close
    objStreamUTF8NoBOM.Close
    
    Set objStreamUTF8 = Nothing
    Set objStreamUTF8NoBOM = Nothing
                    
                    
                    
                    
                    
                    
    .Row = iROW_CNT
    .Col = 1
    
    iROW_CNT = iROW_CNT + 1
    .Row = iROW_CNT
    '���� �ƴϸ� �ݺ��Ѵ�.
    If UCase(Trim(.Value)) <> "END" Then
        iROW_CNT = iROW_CNT + 1
        GoTo LOOP_START
    End If
End With

    gfMsgInf "�Ϸ�Ǿ����ϴ�."
    Exit Sub
Err_Handler:
    gfMsgWrn (Err.Description)
    
End Sub

Private Function sMakeSrcString() As String
Dim sSrc As String
Dim sMethodNm As String



With fpSrc
sSrc = ""



'sSrc = sSrc & "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf
'sSrc = sSrc & "<bo:bizObject "
'sSrc = sSrc & "xmlns:ns2=""http://www.tmax.co.kr/proobject/message"" "
'sSrc = sSrc & "xmlns:ns3=""http://www.tmax.co.kr/proobject/resource"" "
'sSrc = sSrc & "xmlns:ns4=""http://www.tmax.co.kr/proobject/dto"" "
'sSrc = sSrc & "xmlns:ns5=""http://www.tmax.co.kr/proobject/dto/validator"" "
'sSrc = sSrc & "xmlns:ns6=""http://www.tmax.co.kr/proobject/testcase"" "
'sSrc = sSrc & "xmlns:ns7=""http://www.tmax.co.kr/proobject/serviceobject-automatic"" "
'sSrc = sSrc & "xmlns:ns8=""http://www.tmax.co.kr/proobject/sourcecode"" "
'sSrc = sSrc & "xmlns:ns9=""http://www.tmax.co.kr/proobject/dataobjectfactory"" "
'sSrc = sSrc & "xmlns:flow=""http://www.tmax.co.kr/proobject/flow"" "

'sSrc = sSrc & "xmlns:ns11=""http://www.tmax.co.kr/proobject/queryobject"" "
'sSrc = sSrc & "xmlns:ns12=""http://www.tmax.co.kr/proobject/serviceobject"" "
'sSrc = sSrc & "xmlns:bo=""http://www.tmax.co.kr/proobject/bizobject"" "
'sSrc = sSrc & "xmlns:ns14=""http://www.tmax.co.kr/proobject/dataobject"" "
'sSrc = sSrc & "xmlns:ns15=""http://www.tmax.co.kr/proobject/taskobject"" "
'sSrc = sSrc & "xmlns:ns16=""http://www.tmax.co.kr/proobject/jobobject"" "
'sSrc = sSrc & "xmlns:ns17=""http://www.tmax.co.kr/proobject/servicegroup"" "
'sSrc = sSrc & "xmlns:ns18=""http://www.tmax.co.kr/proobject/aspectj"" "
'sSrc = sSrc & "xmlns:ns19=""http://www.tmax.co.kr/proobject/mapping"" "
'sSrc = sSrc & "xmlns:ns20=""http://www.tmax.co.kr/proobject/application/runtime"" "
'sSrc = sSrc & "xmlns:ns21=""http://www.tmaxsoft.co.kr/proobject/testsuite"" "
'sSrc = sSrc & "xmlns:ns22=""http://www.tmaxsoft.co.kr/proobject/testOperation"" "
'sSrc = sSrc & "xmlns:ns23=""http://www.tmax.co.kr/proobject/restriction_codes"" "
'sSrc = sSrc & "xmlns:ns24=""http://www.example.org/externalObjectConfig"" "
'sSrc = sSrc & "xmlns:ns25=""http://www.tmax.co.kr/proobject/application"" "
'sSrc = sSrc & "xmlns:ns26=""http://www.tmax.co.kr/proobject/contents"" "
'sSrc = sSrc & "xmlns:ns27=""http://www.example.org/testData"" "
'sSrc = sSrc & "xmlns:ns28=""http://www.tmaxsoft.co.kr/proobject/testresult"" "
'sSrc = sSrc & "xmlns:ns29=""http://www.tmax.co.kr/proobject/siteConfig"" "
'sSrc = sSrc & "xmlns:ns30=""http://www.tmax.co.kr/proobject/serverConfig"" "
'sSrc = sSrc & "xmlns:ns31=""http://www.tmax.co.kr/proobject/probuilder_config"" "
'sSrc = sSrc & "xmlns:ns32=""http://www.tmax.co.kr/proobject/property"" "
'sSrc = sSrc & "xmlns:ns33=""http://www.tmax.co.kr/proobject/serviceGroup"" "

'sSrc = sSrc & "identifier=""general"" "
'sSrc = sSrc & "resourceId=""" & sClassPName & """ "
'sSrc = sSrc & "logicalName=""" & sClassLName & """ "
'sSrc = sSrc & "physicalName=""" & sClassPName & """ resourceType=""BIZ_OBJECT"" "
'sSrc = sSrc & "resourcePackage=""" & sPackageDotName & """ "
'sSrc = sSrc & "status=""DEV"" creator=""" & sUserId & """ owner=""" & sUserId & """ revision=""" & sRevision & """ "
'sSrc = sSrc & "description=""@���ϸ� : " & sClassPName & "&#xD;&#xA;"
'sSrc = sSrc & "@���� : " & sClassLName & "&#xD;&#xA;"
'sSrc = sSrc & "@�ۼ��� : " & sUserName & "(" & sUserId & ")&#xD;&#xA;--------------------------------------------------------------------&#xD;&#xA;"
'sSrc = sSrc & "�������� : "
'sSrc = sSrc & "������ : "
'sSrc = sSrc & "��û��(SR ��ȣ) "
'sSrc = sSrc & ": ������ �޼ҵ� : "
'sSrc = sSrc & "��������&#xD;&#xA;--------------------------------------------------------------------&#xD;&#xA;"
'sSrc = sSrc & "�ۼ����� : 2020.03.19 : "
'sSrc = sSrc & "�ۼ��� : " & sUserName & "(" & sUserId & "),  "
'sSrc = sSrc & "�ۼ����� : �����ۼ�"" "
'sSrc = sSrc & "resourcePath=""" & sPackageSlashName & """ "
'sSrc = sSrc & "modifier=""" & sUserId & """ "
'sSrc = sSrc & "created=""2020-03-19 15:55:23.393"" lastUpdated=""2020-03-27 17:44:02"">"



sSrc = sSrc & "<?xml version=""1.0"" encoding=""UTF-8""?>"
sSrc = sSrc & vbCrLf & "<bo:bizObject "
sSrc = sSrc & vbCrLf & "xmlns:bo=""http://www.tmax.co.kr/proobject/bizobject"" "
sSrc = sSrc & vbCrLf & "xmlns:flow=""http://www.tmax.co.kr/proobject/flow"" "
sSrc = sSrc & vbCrLf & "resourceId=""" & sClassPname & """ "
sSrc = sSrc & vbCrLf & "logicalName=""" & sClassLname & """ "
sSrc = sSrc & vbCrLf & "physicalName=""" & sClassPname & """ "
sSrc = sSrc & vbCrLf & "resourceType=""BIZ_OBJECT"" "
sSrc = sSrc & vbCrLf & "resourcePackage=""" & sPackageDotName & """ "
sSrc = sSrc & vbCrLf & "status=""DEV"" "
sSrc = sSrc & vbCrLf & "creator=""" & sUserId & """ "
sSrc = sSrc & vbCrLf & "owner=""" & sUserId & """ "
sSrc = sSrc & vbCrLf & "revision=""1"" "
sSrc = sSrc & "description=""@���ϸ� : " & sClassPname & "&#xD;&#xA;"
sSrc = sSrc & "@���� : " & sClassLname & "&#xD;&#xA;"
sSrc = sSrc & "@�ۼ��� : " & sUserName & "(" & sUserId & ")&#xD;&#xA;--------------------------------------------------------------------&#xD;&#xA;"
sSrc = sSrc & "�������� : "
sSrc = sSrc & "������ : "
sSrc = sSrc & "��û��(SR ��ȣ) "
sSrc = sSrc & ": ������ �޼ҵ� : "
sSrc = sSrc & "��������&#xD;&#xA;--------------------------------------------------------------------&#xD;&#xA;"
sSrc = sSrc & "�ۼ����� : 2020.03.19 : "
sSrc = sSrc & "�ۼ��� : " & sUserName & "(" & sUserId & "),  "
sSrc = sSrc & "�ۼ����� : �����ۼ�"" "
sSrc = sSrc & vbCrLf & "resourcePath=""" & sPackageSlashName & """ "
sSrc = sSrc & vbCrLf & "modifier=""" & sUserId & """ "
sSrc = sSrc & vbCrLf & "created=""2020-03-19 15:55:23.393"" "
sSrc = sSrc & vbCrLf & "lastUpdated=""2020-03-27 17:44:02"">"



'----------------------------------------------------
'STEP1 ��������
'----------------------------------------------------
iROW_CNT = iROW_CNT + 1: subSetMehod


sSrc = sSrc & vbCrLf & "    <bo:bizMethod methodName=""" & s�޼ҵ�id & """ "
sSrc = sSrc & "displayName=""" & s�޼ҵ�� & """ modifier=""public"" "
sSrc = sSrc & "comments=""@�޼���� : " & s�޼ҵ�id & "&#xD;&#xA;"
sSrc = sSrc & "@���� : " & s�޼ҵ�� & "&#xD;&#xA; "
sSrc = sSrc & "@�Է� : " & sINBVO & "&#xD;&#xA; "
sSrc = sSrc & "@��� : " & sOUTBVO & """> "



sSrc = sSrc & vbCrLf & "        <flow:flowModule>"
sSrc = sSrc & vbCrLf & "            <flow:entryModule>"
sSrc = sSrc & vbCrLf & "                <flow:name>" & s�޼ҵ�� & "</flow:name>"
sSrc = sSrc & vbCrLf & "                <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:id>0</flow:id>"
sSrc = sSrc & vbCrLf & "                    <flow:id>2</flow:id>"
sSrc = sSrc & vbCrLf & "                    <flow:id>4</flow:id>"
sSrc = sSrc & vbCrLf & "                </flow:nextEntries>"

iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO������ & """ scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO�� & """ resourcePath=""" & sVO��� & """/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"

'vofile ����
Call subMakeVO

iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO������ & """ scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO�� & """ resourcePath=""" & sVO��� & """/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"

'vofile ����
Call subMakeVO

'sSrc = sSrc & vbCrLf & "                <flow:variable kind=""PRIMITIVE"" name=""sTrstTlgmRespCd"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
'sSrc = sSrc & vbCrLf & "                    <flow:primitive>String</flow:primitive>"
'sSrc = sSrc & vbCrLf & "                </flow:variable>"

sSrc = sSrc & vbCrLf & "            </flow:entryModule>"
sSrc = sSrc & vbCrLf & "            <flow:innerModules>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""0"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>��������ó��</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>1</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""2"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>���� �� ó��</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>3</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""4"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>��������ó��</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>5</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "            </flow:innerModules>"
sSrc = sSrc & vbCrLf & "            <flow:xorModules/>"
sSrc = sSrc & vbCrLf & "            <flow:virtualModules>"
sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""1"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>�Է°� Ȯ��</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                 logger.info(""  �١ڡ١ڡ١ڡ١�inBvo  " & s�޼ҵ�id & " = \n[""+inBvo.toString()+""] �١ڡ١ڡ١ڡ١�"" );&#xD;"
'sSrc = sSrc & vbCrLf & "                    tlgmHdrBvo = inBvo.getRecvHdrVo();&#xD;"
'sSrc = sSrc & vbCrLf & "                    //��������&#xD;"
'sSrc = sSrc & vbCrLf & "                    tlgmHdrBvo.setRecvDy(DateUtil.getCurrentDate(2));&#xD;"
'sSrc = sSrc & vbCrLf & "                    //���Žð�&#xD;"
'sSrc = sSrc & vbCrLf & "                    tlgmHdrBvo.setRecvPtm(DateUtil.getCurrentTime(""HHmmss""));&#xD;"
'sSrc = sSrc & vbCrLf & "                    &#xD;"
'sSrc = sSrc & vbCrLf & "                    if(inBvo.getRecvDataVo() == null){&#xD;"
'sSrc = sSrc & vbCrLf & "                         //�����ڷ�����&#xD;"
'sSrc = sSrc & vbCrLf & "                         tlgmHdrBvo.setDtalDataYn(""N"");&#xD;"
'sSrc = sSrc & vbCrLf & "                     } else {&#xD;"
'sSrc = sSrc & vbCrLf & "                         //�����ڷ�����&#xD;"
'sSrc = sSrc & vbCrLf & "                         tlgmHdrBvo.setDtalDataYn(""Y"");&#xD;"
'sSrc = sSrc & vbCrLf & "                         //�ֹι��ι�ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                         tlgmHdrBvo.setJno(inBvo.getRecvDataVo().getJno()); &#xD;"
'sSrc = sSrc & vbCrLf & "                     }&#xD;"
'sSrc = sSrc & vbCrLf & "                     sTrstTlgmRespCd = HgIgCommConst.TRST_TLGM_RESP_CD_0000;</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"



sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""3"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>����ó��</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                  try{&#xD;"
'sSrc = sSrc & vbCrLf & "                         /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                         * ���ŷα� ����&#xD;"
'sSrc = sSrc & vbCrLf & "                         *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                         HgigTrstTlgmMngBiz tlgmMngBiz = new HgigTrstTlgmMngBiz();&#xD;"
'sSrc = sSrc & vbCrLf & "                         HgigTrstTlgmMng01InBvo recvLogInBvo = new HgigTrstTlgmMng01InBvo();&#xD;"
'sSrc = sSrc & vbCrLf & "                         recvLogInBvo.setTrstTlgmHdrVo(tlgmHdrBvo);                      //���&#xD;"
'sSrc = sSrc & vbCrLf & "                         &#xD;"
'sSrc = sSrc & vbCrLf & "                         tlgmMngBiz.insRecvLog(recvLogInBvo);&#xD;"
'sSrc = sSrc & vbCrLf & "                         &#xD;"
'sSrc = sSrc & vbCrLf & "                         /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                         * ����ó�� - �������� ��ȸ&#xD;"
'sSrc = sSrc & vbCrLf & "                         *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                         inBvo.setRecvHdrVo(tlgmHdrBvo);&#xD;"
'sSrc = sSrc & vbCrLf & "                         outBvo = this.selListGrntCont(inBvo);&#xD;"
'sSrc = sSrc & vbCrLf & "                         //��������&#xD;"
'sSrc = sSrc & vbCrLf & "                         tlgmHdrBvo.setTrstTlgmRespCd(sTrstTlgmRespCd);&#xD;"
'sSrc = sSrc & vbCrLf & "                     } catch( Exception ex){&#xD;"
'sSrc = sSrc & vbCrLf & "                         if (ex instanceof UserException){&#xD;"
'sSrc = sSrc & vbCrLf & "                             /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                             * �����α� ���� ����&#xD;"
'sSrc = sSrc & vbCrLf & "                             *****************************************************************/&#xD;"
'sSrc = sSrc & vbCrLf & "                             //TODO: ���������� ���� �κ� �߰� �ʿ���.&#xD;"
'sSrc = sSrc & vbCrLf & "                             sTrstTlgmRespCd = ((UserException) ex).getReturnCode();&#xD;"
'sSrc = sSrc & vbCrLf & "                             tlgmHdrBvo.setTrstTlgmRespCd(sTrstTlgmRespCd);&#xD;"
'sSrc = sSrc & vbCrLf & "                             //�������α׷�����&#xD;"
'sSrc = sSrc & vbCrLf & "                             tlgmHdrBvo.setErrPgrmCont(""�����׽�Ʈ"");&#xD;"
'sSrc = sSrc & vbCrLf & "                             //�������μ���&#xD;"
'sSrc = sSrc & vbCrLf & "                             tlgmHdrBvo.setErrLineSeq(0L); &#xD;"
'sSrc = sSrc & vbCrLf & "                             //�����޽�������&#xD;"
'sSrc = sSrc & vbCrLf & "                             tlgmHdrBvo.setErrMsgCont(((UserException) ex).getMessage()); &#xD;"
'sSrc = sSrc & vbCrLf & "                         } else {&#xD;"
'sSrc = sSrc & vbCrLf & "                             //TODO:&#xD;"
'sSrc = sSrc & vbCrLf & "                             //throw new UserException(""90000"");&#xD;"
'sSrc = sSrc & vbCrLf & "                             //throw new LiveException(""zz.error.90000"", ex);&#xD;"
'sSrc = sSrc & vbCrLf & "                             throw ex;&#xD;"
'sSrc = sSrc & vbCrLf & "                         }&#xD;"
'sSrc = sSrc & vbCrLf & "                     }&#xD;"
'sSrc = sSrc & vbCrLf & "                     /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                     * �����۽�&#xD;"
'sSrc = sSrc & vbCrLf & "                     *****************************************************************/&#xD;"
'sSrc = sSrc & vbCrLf & "                     HgigGrntContInq04InBvo sendInBvo = new HgigGrntContInq04InBvo();&#xD;"
'sSrc = sSrc & vbCrLf & "                     sendInBvo.setSendHdrVo(tlgmHdrBvo);&#xD;"
'sSrc = sSrc & vbCrLf & "                     sendInBvo.setSendDataVo(outBvo.getGrntContVo());&#xD;"
'sSrc = sSrc & vbCrLf & "                     this.sendTlgm1602(sendInBvo);</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"


sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""5"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>return</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <flow:code>&#xD;"
sSrc = sSrc & vbCrLf & "                     logger.info(""  �١ڡ١ڡ١ڡ١�outBvo  " & s�޼ҵ�id & " = \n [""+outBvo.toString()+""] �١ڡ١ڡ١ڡ١�"" );&#xD;"
sSrc = sSrc & vbCrLf & "                     return outBvo;</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "            </flow:virtualModules>"
sSrc = sSrc & vbCrLf & "            <flow:loopModules/>"
sSrc = sSrc & vbCrLf & "            <flow:bizMethodCalls/>"
sSrc = sSrc & vbCrLf & "            <flow:serviceCalls/>"
sSrc = sSrc & vbCrLf & "            <flow:assignModules/>"
sSrc = sSrc & vbCrLf & "            <flow:dataObjectCalls/>"
sSrc = sSrc & vbCrLf & "            <flow:replyCalls/>"
sSrc = sSrc & vbCrLf & "        </flow:flowModule>"



iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "        <bo:arguments>"
sSrc = sSrc & vbCrLf & "            <flow:argument>"
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO������ & """ comments=""0'th argument"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO�� & """ resourcePath=""" & sVO��� & """/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"
sSrc = sSrc & vbCrLf & "                <flow:sequence>0</flow:sequence>"
sSrc = sSrc & vbCrLf & "            </flow:argument>"
sSrc = sSrc & vbCrLf & "        </bo:arguments>"

'vofile ����
Call subMakeVO

iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "        <bo:returnType kind=""DTO"">"
sSrc = sSrc & vbCrLf & "            <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO�� & """ resourcePath=""" & sVO��� & """/>"
sSrc = sSrc & vbCrLf & "        </bo:returnType>"
sSrc = sSrc & vbCrLf & "        <bo:classInfo classPackageName=""" & sPackageSlashName & """ className=""" & sClassPname & """/>"
sSrc = sSrc & vbCrLf & "        <bo:annotations/>"
sSrc = sSrc & vbCrLf & "    </bo:bizMethod>"
'vofile ����
Call subMakeVO


'----------------------------------------------
'STEP2 ���� �۽�
'----------------------------------------------
iROW_CNT = iROW_CNT + 1: subSetMehod

sSrc = sSrc & vbCrLf & "    <bo:bizMethod methodName=""" & s�޼ҵ�id & """ "
sSrc = sSrc & "displayName=""" & s�޼ҵ�� & """ modifier=""public"" "
sSrc = sSrc & "comments=""@�޼���� : " & s�޼ҵ�id & "&#xD;&#xA;"
sSrc = sSrc & "@���� : " & s�޼ҵ�� & "&#xD;&#xA; "
sSrc = sSrc & "@�Է� : " & sINBVO & "&#xD;&#xA; "
sSrc = sSrc & "@��� : " & sOUTBVO & """> "
sSrc = sSrc & vbCrLf & "        <flow:flowModule>"
sSrc = sSrc & vbCrLf & "            <flow:entryModule>"
sSrc = sSrc & vbCrLf & "                <flow:name>" & s�޼ҵ�� & "</flow:name>"
sSrc = sSrc & vbCrLf & "                <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:id>0</flow:id>"
sSrc = sSrc & vbCrLf & "                    <flow:id>2</flow:id>"
sSrc = sSrc & vbCrLf & "                    <flow:id>4</flow:id>"
sSrc = sSrc & vbCrLf & "                </flow:nextEntries>"




'VO
iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO������ & """ scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO�� & """ resourcePath=""" & sVO��� & """/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"

'vofile ����
Call subMakeVO

'VO
iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO������ & """ scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO�� & """ resourcePath=""" & sVO��� & """/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"
'vofile ����
Call subMakeVO


sSrc = sSrc & vbCrLf & "            </flow:entryModule>"
sSrc = sSrc & vbCrLf & "            <flow:innerModules>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""0"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>��������ó��</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>1</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""2"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>���� �� ó��</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>3</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""4"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>��������ó��</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>5</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "            </flow:innerModules>"
sSrc = sSrc & vbCrLf & "            <flow:xorModules/>"

sSrc = sSrc & vbCrLf & "            <flow:virtualModules>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""1"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>�Է°� Ȯ��</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                  logger.info(""  �١ڡ١ڡ١ڡ١�inBvo  " & s�޼ҵ�id & " = \n[""+inBvo.toString()+""] �١ڡ١ڡ١ڡ١�"" );&#xD;"
'sSrc = sSrc & vbCrLf & "                 tlgmHdrBvo = inBvo.getSendHdrVo();&#xD;"
'sSrc = sSrc & vbCrLf & "                 //��Ź�����ڵ�&#xD;"
'sSrc = sSrc & vbCrLf & "                 tlgmHdrBvo.setTrstTlgmCd(""" & Right(s�޼ҵ�id, 4) & """);</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"


sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""3"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>�۽�ó��</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                   &#xD;"
'sSrc = sSrc & vbCrLf & "                  //���������ΰ��&#xD;"
'sSrc = sSrc & vbCrLf & "                  if(!StringUtil.isEmpty(tlgmHdrBvo.getTrstTlgmRespCd())&#xD;"
'sSrc = sSrc & vbCrLf & "                  &amp;&amp; !HgIgCommConst.TRST_TLGM_RESP_CD_0000.equals(tlgmHdrBvo.getTrstTlgmRespCd())){&#xD;"
'sSrc = sSrc & vbCrLf & "                      /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                      * �����α� ����&#xD;"
'sSrc = sSrc & vbCrLf & "                      *****************************************************************/&#xD;"
'sSrc = sSrc & vbCrLf & "                      HgigTrstTlgmMngBiz tlgmMngBiz = new HgigTrstTlgmMngBiz();&#xD;"
'sSrc = sSrc & vbCrLf & "                      HgigTrstTlgmMng03InBvo sendLogInBvo = new HgigTrstTlgmMng03InBvo();&#xD;"
'sSrc = sSrc & vbCrLf & "                      sendLogInBvo.setTrstTlgmHdrVo(tlgmHdrBvo);&#xD;"
'sSrc = sSrc & vbCrLf & "                      &#xD;"
'sSrc = sSrc & vbCrLf & "                      tlgmMngBiz.insErrLog(sendLogInBvo);&#xD;"
'sSrc = sSrc & vbCrLf & "                  }&#xD;"
'sSrc = sSrc & vbCrLf & "                  /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                  * �۽ŷα� ����&#xD;"
'sSrc = sSrc & vbCrLf & "                  *****************************************************************/&#xD;"
'sSrc = sSrc & vbCrLf & "                  HgigTrstTlgmMngBiz tlgmMngBiz = new HgigTrstTlgmMngBiz();&#xD;"
'sSrc = sSrc & vbCrLf & "                  HgigTrstTlgmMng02InBvo sendLogInBvo = new HgigTrstTlgmMng02InBvo();&#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  //���&#xD;"
'sSrc = sSrc & vbCrLf & "                  sendLogInBvo.setTrstTlgmHdrVo(tlgmHdrBvo);&#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  tlgmMngBiz.insSendLog(sendLogInBvo);&#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                  * �����۽�ó��&#xD;"
'sSrc = sSrc & vbCrLf & "                  *****************************************************************/&#xD;"
'sSrc = sSrc & vbCrLf & "                  //TODO:</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"


sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""5"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>return</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <flow:code>&#xD;"
sSrc = sSrc & vbCrLf & "                  logger.info(""  �١ڡ١ڡ١ڡ١�outBvo  " & s�޼ҵ�id & " = \n [""+outBvo.toString()+""] �١ڡ١ڡ١ڡ١�"" );&#xD;"
sSrc = sSrc & vbCrLf & "                  return outBvo;</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"


sSrc = sSrc & vbCrLf & "            </flow:virtualModules>"

sSrc = sSrc & vbCrLf & "            <flow:loopModules/>"
sSrc = sSrc & vbCrLf & "            <flow:bizMethodCalls/>"
sSrc = sSrc & vbCrLf & "            <flow:serviceCalls/>"
sSrc = sSrc & vbCrLf & "            <flow:assignModules/>"
sSrc = sSrc & vbCrLf & "            <flow:dataObjectCalls/>"
sSrc = sSrc & vbCrLf & "            <flow:replyCalls/>"
sSrc = sSrc & vbCrLf & "        </flow:flowModule>"


sSrc = sSrc & vbCrLf & "        <bo:arguments>"
sSrc = sSrc & vbCrLf & "            <flow:argument>"

iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO������ & """ comments=""0'th argument"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO�� & """ resourcePath=""hf/cor/hg/ig/vo/bvo/sbvo""/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"
sSrc = sSrc & vbCrLf & "                <flow:sequence>0</flow:sequence>"
sSrc = sSrc & vbCrLf & "            </flow:argument>"
sSrc = sSrc & vbCrLf & "        </bo:arguments>"
'vofile ����
Call subMakeVO


iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "        <bo:returnType kind=""DTO"">"
sSrc = sSrc & vbCrLf & "            <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO�� & """ resourcePath=""" & sVO��� & """/>"
sSrc = sSrc & vbCrLf & "        </bo:returnType>"
sSrc = sSrc & vbCrLf & "        <bo:classInfo classPackageName=""" & sPackageSlashName & """ className=""" & sClassPname & """/>"
sSrc = sSrc & vbCrLf & "        <bo:annotations/>"
sSrc = sSrc & vbCrLf & "    </bo:bizMethod>"

'vofile ����
Call subMakeVO

'----------------------------------------------------
'STEP3 ���ų��� ��ȸ
'----------------------------------------------------
iROW_CNT = iROW_CNT + 1: subSetMehod

sSrc = sSrc & vbCrLf & "    <bo:bizMethod methodName=""" & s�޼ҵ�id & """ "
sSrc = sSrc & "displayName=""" & s�޼ҵ�� & """ modifier=""public"" "
sSrc = sSrc & "comments=""@�޼���� : " & s�޼ҵ�id & "&#xD;&#xA;"
sSrc = sSrc & "@���� : " & s�޼ҵ�� & "&#xD;&#xA; "
sSrc = sSrc & "@�Է� : " & sINBVO & "&#xD;&#xA; "
sSrc = sSrc & "@��� : " & sOUTBVO & """> "
sSrc = sSrc & vbCrLf & "        <flow:flowModule>"
sSrc = sSrc & vbCrLf & "            <flow:entryModule>"
sSrc = sSrc & vbCrLf & "                <flow:name>" & s�޼ҵ�� & "</flow:name>"
sSrc = sSrc & vbCrLf & "                <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:id>0</flow:id>"
sSrc = sSrc & vbCrLf & "                    <flow:id>2</flow:id>"
sSrc = sSrc & vbCrLf & "                    <flow:id>4</flow:id>"
sSrc = sSrc & vbCrLf & "                </flow:nextEntries>"


'VO
iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO������ & """ scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO�� & """ resourcePath=""" & sVO��� & """/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"

'vofile ����
Call subMakeVO

'VO
iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO������ & """ scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO�� & """ resourcePath=""" & sVO��� & """/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"

'vofile ����
Call subMakeVO

'VO
iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO������ & """ scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO�� & """ resourcePath=""" & sVO��� & """/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"

'vofile ����
Call subMakeVO


sSrc = sSrc & vbCrLf & "            </flow:entryModule>"
sSrc = sSrc & vbCrLf & "            <flow:innerModules>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""0"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>��������ó��</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>1</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""2"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>���� �� ó��</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>3</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""4"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>��������ó��</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>5</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "            </flow:innerModules>"
sSrc = sSrc & vbCrLf & "            <flow:xorModules/>"


sSrc = sSrc & vbCrLf & "            <flow:virtualModules>"
sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""1"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>�Է°� Ȯ��</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                  logger.info(""  �١ڡ١ڡ١ڡ١�inBvo   " & s�޼ҵ�id & " = \n[""+inBvo.toString()+""] �١ڡ١ڡ١ڡ١�"" );&#xD;"
'sSrc = sSrc & vbCrLf & "                 &#xD;"
'sSrc = sSrc & vbCrLf & "                 if(inBvo.getInqComnVo() == null){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""��ȸ����"" + ""��(��) �ʼ��Է��Դϴ�."");&#xD;"
'sSrc = sSrc & vbCrLf & "                 }&#xD;"
'sSrc = sSrc & vbCrLf & "                 //��������&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(StringUtil.isEmpty(inBvo.getInqComnVo().getRecvDy())){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""��������"" + ""��(��) �ʼ��Է��Դϴ�."");&#xD;"
'sSrc = sSrc & vbCrLf & "                     &#xD;"
'sSrc = sSrc & vbCrLf & "                 }&#xD;"
'sSrc = sSrc & vbCrLf & "                 //���Žð�&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(StringUtil.isEmpty(inBvo.getInqComnVo().getRecvPtm())){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""���Žð�"" + ""��(��) �ʼ��Է��Դϴ�."");&#xD;"
'sSrc = sSrc & vbCrLf & "                 }&#xD;"
'sSrc = sSrc & vbCrLf & "                 //��������ڵ�&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(StringUtil.isEmpty(inBvo.getInqComnVo().getOrgGiroCd())){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""��������ڵ�"" + ""��(��) �ʼ��Է��Դϴ�."");&#xD;"
'sSrc = sSrc & vbCrLf & "                 }&#xD;"
'sSrc = sSrc & vbCrLf & "                 //��Ź�����ڵ�&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(StringUtil.isEmpty(inBvo.getInqComnVo().getTrstTlgmCd())){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""��Ź�����ڵ�"" + ""��(��) �ʼ��Է��Դϴ�."");&#xD;"
'sSrc = sSrc & vbCrLf & "                 }&#xD;"
'sSrc = sSrc & vbCrLf & "                 //����������ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(StringUtil.isEmpty(inBvo.getInqComnVo().getTlgmMngNo())){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""����������ȣ"" + ""��(��) �ʼ��Է��Դϴ�."");&#xD;"
'sSrc = sSrc & vbCrLf & "                 }&#xD;"
'sSrc = sSrc & vbCrLf & "                 //��������&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(inBvo.getInqComnVo().getTlgmSeq() &lt;= 0){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""��������"" + ""��(��) �ʼ��Է��Դϴ�."");&#xD;"
'sSrc = sSrc & vbCrLf & "                 }</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""3"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>���ų��� ��ȸ</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                   DataObjectMapper&lt;TbHgc601lHstrstrecv001Dvo&gt; mapper = new DataObjectMapper&lt;TbHgc601lHstrstrecv001Dvo&gt;( ""po7_app"", TbHgc601lHstrstrecv001Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setQuery(TbHgc601lHstrstrecv001Dao.SEL001);&#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  //��������&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""recvDy"", inBvo.getInqComnVo().getRecvDy());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //���Žð�&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""recvPtm"", inBvo.getInqComnVo().getRecvPtm());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //��������ڵ�&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""orgGiroCd"", inBvo.getInqComnVo().getOrgGiroCd());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //��Ź�����ڵ�&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""trstTlgmCd"", inBvo.getInqComnVo().getTrstTlgmCd());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //����������ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""tlgmMngNo"", inBvo.getInqComnVo().getTlgmMngNo());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //��������&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""tlgmSeq"", inBvo.getInqComnVo().getTlgmSeq());&#xD;"
'sSrc = sSrc & vbCrLf & "&#xD;"
'sSrc = sSrc & vbCrLf & "                  TbHgc601lHstrstrecv001Dvo reDvo = mapper.get();&#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                  * HEAD ����&#xD;"
'sSrc = sSrc & vbCrLf & "                  *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                  //�������� &#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setRecvDy(reDvo.getRecvDy());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //���Žð� &#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setRecvPtm(reDvo.getRecvPtm());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //��������ڵ� &#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setOrgGiroCd(reDvo.getOrgGiroCd());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //�۽����� &#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setSendDy(reDvo.getSendDy());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //�۽Žð� &#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setSendPtm(reDvo.getSendPtm());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //��Ź�����ڵ�&#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setTrstTlgmCd(reDvo.getTrstTlgmCd());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //��Ź�����   &#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setCnsgOrgNm(reDvo.getCnsgOrgNm());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //����������ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setTlgmMngNo(reDvo.getTlgmMngNo());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //��Ź���������ڵ�&#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setTrstTlgmRespCd(reDvo.getTrstTlgmRespCd());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //���ŵ����ͳ����ȣȭ&#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setRecvDataContEnc(reDvo.getRecvDataContEnc());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //��������&#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setTlgmLen(reDvo.getTlgmLen());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //�������μ���&#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setErrLineSeq(reDvo.getErrLineSeq());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //�������α׷�����&#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setErrPgrmCont(reDvo.getErrPgrmCont());&#xD;"
'sSrc = sSrc & vbCrLf & "&#xD;"
'sSrc = sSrc & vbCrLf & "                  /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                  * DATA ����&#xD;"
'sSrc = sSrc & vbCrLf & "                  *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                  //�ֹι�ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                  dataVo.setJno(reDvo.getRecvDataContEnc().trim());</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""5"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>return</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                   outBvo.setRecvHdrVo(hdrVo);&#xD;"
'sSrc = sSrc & vbCrLf & "                  outBvo.setRecvDataVo(dataVo);&#xD;"
'sSrc = sSrc & vbCrLf & "                  logger.info(""  �١ڡ١ڡ١ڡ١�outBvo   " & s�޼ҵ�id & " = \n [""+outBvo.toString()+""] �١ڡ١ڡ١ڡ١�"" );&#xD;"
'sSrc = sSrc & vbCrLf & "                  return outBvo;</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "            </flow:virtualModules>"

sSrc = sSrc & vbCrLf & "            <flow:loopModules/>"
sSrc = sSrc & vbCrLf & "            <flow:bizMethodCalls/>"
sSrc = sSrc & vbCrLf & "            <flow:serviceCalls/>"
sSrc = sSrc & vbCrLf & "            <flow:assignModules/>"
sSrc = sSrc & vbCrLf & "            <flow:dataObjectCalls/>"
sSrc = sSrc & vbCrLf & "            <flow:replyCalls/>"
sSrc = sSrc & vbCrLf & "        </flow:flowModule>"
sSrc = sSrc & vbCrLf & "        <bo:arguments>"
sSrc = sSrc & vbCrLf & "            <flow:argument>"

iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO������ & """ comments=""0'th argument"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO�� & """ resourcePath=""hf/cor/hg/ig/vo/bvo/sbvo""/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"
sSrc = sSrc & vbCrLf & "                <flow:sequence>0</flow:sequence>"
sSrc = sSrc & vbCrLf & "            </flow:argument>"
sSrc = sSrc & vbCrLf & "        </bo:arguments>"

'vofile ����
Call subMakeVO

iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "        <bo:returnType kind=""DTO"">"
sSrc = sSrc & vbCrLf & "            <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO�� & """ resourcePath=""" & sVO��� & """/>"
sSrc = sSrc & vbCrLf & "        </bo:returnType>"
sSrc = sSrc & vbCrLf & "        <bo:classInfo classPackageName=""" & sPackageSlashName & """ className=""" & sClassPname & """/>"
sSrc = sSrc & vbCrLf & "        <bo:annotations/>"
sSrc = sSrc & vbCrLf & "    </bo:bizMethod>"

'vofile ����
Call subMakeVO


'----------------------------------------------------
'STEP4 �������� ��ȸ biz ó��
'----------------------------------------------------


iROW_CNT = iROW_CNT + 1: subSetMehod

sSrc = sSrc & vbCrLf & "    <bo:bizMethod methodName=""" & s�޼ҵ�id & """ "
sSrc = sSrc & "displayName=""" & s�޼ҵ�� & """ modifier=""public"" "
sSrc = sSrc & "comments=""@�޼���� : " & s�޼ҵ�id & "&#xD;&#xA;"
sSrc = sSrc & "@���� : " & s�޼ҵ�� & "&#xD;&#xA; "
sSrc = sSrc & "@�Է� : " & sINBVO & "&#xD;&#xA; "
sSrc = sSrc & "@��� : " & sOUTBVO & """> "
sSrc = sSrc & vbCrLf & "        <flow:flowModule>"
sSrc = sSrc & vbCrLf & "            <flow:entryModule>"
sSrc = sSrc & vbCrLf & "                <flow:name>" & s�޼ҵ�� & "</flow:name>"
sSrc = sSrc & vbCrLf & "                <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:id>0</flow:id>"
sSrc = sSrc & vbCrLf & "                    <flow:id>2</flow:id>"
sSrc = sSrc & vbCrLf & "                    <flow:id>4</flow:id>"
sSrc = sSrc & vbCrLf & "                </flow:nextEntries>"


'VO
iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO������ & """ scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO�� & """ resourcePath=""" & sVO��� & """/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"

'vofile ����
Call subMakeVO

''VO
'sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""custInfoVo"" arrayKind=""none"" size=""0"" scope=""local"" initObject="""" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"" inject=""NONE"" static=""false"" final=""false"">"
'sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""HgigGrntContInq001Dvo"" logicalName=""����������ȸ001Dvo"" resourcePath=""hf/cor/hg/ig/vo/dvo/sdvo""/>"
'sSrc = sSrc & vbCrLf & "                </flow:variable>"
'
''VO
'sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""grntContList"" arrayKind=""variable"" size=""0"" scope=""local"" initObject="""" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"" inject=""NONE"" static=""false"" final=""false"">"
'sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""HgigGrntContInq002Dvo"" logicalName=""������û������ȸDvo"" resourcePath=""hf/cor/hg/ig/vo/dvo/sdvo""/>"
'sSrc = sSrc & vbCrLf & "                </flow:variable>"
'
''VO
'sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""grnt1601List"" arrayKind=""variable"" size=""0"" scope=""local"" initObject="""" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"" inject=""NONE"" static=""false"" final=""false"">"
'sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""HgigTrstTlgmSendData1601GrntBvo"" logicalName=""��Ź�����۽ŵ�����1601��������VO"" resourcePath=""hf/cor/hg/ig/vo/bvo/comn""/>"
'sSrc = sSrc & vbCrLf & "                </flow:variable>"
'
''VO
'sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""weak1601List"" arrayKind=""variable"" size=""0"" scope=""local"" initObject="""" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"" inject=""NONE"" static=""false"" final=""false"">"
'sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""HgigTrstTlgmSendData1601WeakBvo"" logicalName=""��Ź�����۽ŵ�����1601�ν��ڷ�Vo"" resourcePath=""hf/cor/hg/ig/vo/bvo/comn""/>"
'sSrc = sSrc & vbCrLf & "                </flow:variable>"
'
''VO
'sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""send1601Vo"" arrayKind=""none"" size=""0"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"" inject=""NONE"" static=""false"" final=""false"">"
'sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""HgigTrstTlgmSendData1601Bvo"" logicalName=""��Ź�����۽ŵ�����1601Vo"" resourcePath=""hf/cor/hg/ig/vo/bvo/comn""/>"
'sSrc = sSrc & vbCrLf & "                </flow:variable>"
'
''VO
'sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""yth1601List"" arrayKind=""variable"" size=""0"" scope=""local"" initObject="""" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"" final=""false"">"
'sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""HgigTrstTlgmSendData1601YthBvo"" logicalName=""��Ź�����۽ŵ�����1601û�����Vo"" resourcePath=""hf/cor/hg/ig/vo/bvo/comn""/>"
'sSrc = sSrc & vbCrLf & "                </flow:variable>"


sSrc = sSrc & vbCrLf & "            </flow:entryModule>"
sSrc = sSrc & vbCrLf & "            <flow:innerModules>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""0"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>��������ó��</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>1</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""2"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>���� �� ó��</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>3</flow:id>"
sSrc = sSrc & vbCrLf & "                        <flow:id>7</flow:id>"
sSrc = sSrc & vbCrLf & "                        <flow:id>10</flow:id>"
sSrc = sSrc & vbCrLf & "                        <flow:id>13</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""4"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>��������ó��</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>5</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""if"" ifElseWhenLast=""false"" nodeId=""15"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>������</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:comments>���Һ������� �Ϲݰ�, ������ �Ǵ�</flow:comments>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>11</flow:id>"
sSrc = sSrc & vbCrLf & "                        <flow:id>18</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition>"
sSrc = sSrc & vbCrLf & "                        <flow:code>!StringUtil.isEmpty(custInfoVo.getCtrlBrofCd())</flow:code>"
sSrc = sSrc & vbCrLf & "                    </flow:condition>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""7"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>����������ȸ</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>6</flow:id>"
sSrc = sSrc & vbCrLf & "                        <flow:id>8</flow:id>"
sSrc = sSrc & vbCrLf & "                        <flow:id>9</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""13"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>��������</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>20</flow:id>"
sSrc = sSrc & vbCrLf & "                        <flow:id>17</flow:id>"
sSrc = sSrc & vbCrLf & "                        <flow:id>12</flow:id>"
sSrc = sSrc & vbCrLf & "                        <flow:id>15</flow:id>"
sSrc = sSrc & vbCrLf & "                        <flow:id>19</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "            </flow:innerModules>"
sSrc = sSrc & vbCrLf & "            <flow:xorModules/>"

sSrc = sSrc & vbCrLf & "            <flow:virtualModules>"
sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""1"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>�Է°� Ȯ��</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'Src = sSrc & vbCrLf & "                    <flow:code>                  logger.info(""  �١ڡ١ڡ١ڡ١�inBvo  selListGrntCont = \n[""+inBvo.toString()+""] �١ڡ١ڡ١ڡ١�"" );&#xD;"
'Src = sSrc & vbCrLf & "                 &#xD;"
'sSrc = sSrc & vbCrLf & "                 //��������ڵ�&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(StringUtil.isEmpty(inBvo.getRecvHdrVo().getOrgGiroCd())){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""��������ڵ�"" + ""��(��) �ʼ��Է��Դϴ�.""); &#xD;"
'sSrc = sSrc & vbCrLf & "                     //LiveException ex = new LiveException(""hg.info.00921"", pVOs);&#xD;"
'sSrc = sSrc & vbCrLf & "                 }&#xD;"
'sSrc = sSrc & vbCrLf & "                 String sJno = inBvo.getRecvDataVo().getJno();&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(sJno.length() != 13 || ""0000000000000"".equals(sJno)){&#xD;"
'sSrc = sSrc & vbCrLf & "                     //TODO: ����ó�� ����&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new UserException(""1234"", ""�ֹι�ȣ"");&#xD;"
'sSrc = sSrc & vbCrLf & "                     //LiveException ex = new LiveException(""hg.info.00829"", pVOs);&#xD;"
'sSrc = sSrc & vbCrLf & "                 }</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""5"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>return</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                   //��������&#xD;"
'sSrc = sSrc & vbCrLf & "                  send1601Vo.setGrntContList(grnt1601List);&#xD;"
'sSrc = sSrc & vbCrLf & "                  //�ν��ڷ����⳻��&#xD;"
'sSrc = sSrc & vbCrLf & "                  send1601Vo.setWeakDataList(weak1601List);&#xD;"
'sSrc = sSrc & vbCrLf & "                  //û������ڱݺ�������&#xD;"
'sSrc = sSrc & vbCrLf & "                  send1601Vo.setYthMmrtList(yth1601List);&#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  outBvo.setGrntContVo(send1601Vo);&#xD;"
'sSrc = sSrc & vbCrLf & "                  logger.info(""  �١ڡ١ڡ١ڡ١�outBvo  selListGrntCont = \n [""+outBvo.toString()+""] �١ڡ١ڡ١ڡ١�"" );&#xD;"
'sSrc = sSrc & vbCrLf & "                  return outBvo;</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"


sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""6"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>�� ������û���� ��ȸ</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                       /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                      * ������û���� ��ȸ(MAX 20��)&#xD;"
'sSrc = sSrc & vbCrLf & "                      *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                      DataObjectMapper&lt;HgigGrntContInq002Dvo&gt; mapper = new DataObjectMapper&lt;HgigGrntContInq002Dvo&gt;( ""po7_app"", HgigGrntContInq002Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                      mapper.setQuery(HgigGrntContInq001Dao.SELLIST001);&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                      //�ֹι�ȣ TODO: ��ȣȭó��&#xD;"
'sSrc = sSrc & vbCrLf & "                      mapper.setParameter(""csno"", custInfoVo.getCsno());&#xD;"
'sSrc = sSrc & vbCrLf & "                      &#xD;"
'sSrc = sSrc & vbCrLf & "                      for(HgigGrntContInq002Dvo grntContVo : mapper.getForwardList()){&#xD;"
'sSrc = sSrc & vbCrLf & "                          grntContList.add(grntContVo);&#xD;"
'sSrc = sSrc & vbCrLf & "                      }</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""8"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>��������ں��� ��û���� ��ȸ</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                       /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                      * ��������ں������� ��û���� ��ȸ&#xD;"
'sSrc = sSrc & vbCrLf & "                      *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                      DataObjectMapper&lt;HgigGrntContInq002Dvo&gt; mapper = new DataObjectMapper&lt;HgigGrntContInq002Dvo&gt;( ""po7_app"", HgigGrntContInq002Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                      mapper.setQuery(HgigGrntContInq001Dao.SELLIST002);&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                      //����ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                      mapper.setParameter(""csno"", custInfoVo.getCsno());&#xD;"
'sSrc = sSrc & vbCrLf & "                      &#xD;"
'sSrc = sSrc & vbCrLf & "                      for(HgigGrntContInq002Dvo JintGrntContVo : mapper.getForwardList()){&#xD;"
'sSrc = sSrc & vbCrLf & "                          if(grntContList.size() &gt; 20){&#xD;"
'sSrc = sSrc & vbCrLf & "                              break;&#xD;"
'sSrc = sSrc & vbCrLf & "                          }&#xD;"
'sSrc = sSrc & vbCrLf & "                          grntContList.add(JintGrntContVo);&#xD;"
'sSrc = sSrc & vbCrLf & "                      }</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""9"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>������û���� �߰����� ��ȸ</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                       //�����ܾ�&#xD;"
'sSrc = sSrc & vbCrLf & "                      BigDecimal bdGrntRamt = BigDecimal.ZERO;&#xD;"
'sSrc = sSrc & vbCrLf & "                      hf.cor.hg.ig.vo.bvo.comn.HgigTrstTlgmSendData1601GrntBvo grnt1601Vo = null;&#xD;"
'sSrc = sSrc & vbCrLf & "                      for (int iGrntCont = 0; iGrntCont &lt; grntContList.size(); iGrntCont++) {&#xD;"
'sSrc = sSrc & vbCrLf & "                          grnt1601Vo = new hf.cor.hg.ig.vo.bvo.comn.HgigTrstTlgmSendData1601GrntBvo();&#xD;"
'sSrc = sSrc & vbCrLf & "                          //�����ܾ�&#xD;"
'sSrc = sSrc & vbCrLf & "                          bdGrntRamt = grntContList.get(iGrntCont).getGrntRamt();&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                          /* 1.����(����)���� */&#xD;"
'sSrc = sSrc & vbCrLf & "                          if(bdGrntRamt.compareTo(BigDecimal.ZERO) &gt; 0){&#xD;"
'sSrc = sSrc & vbCrLf & "                              //��������(&lt;-���ʺ�������)&#xD;"
'sSrc = sSrc & vbCrLf & "                              grnt1601Vo.setGrntDy(grntContList.get(iGrntCont).getOgnlGrntDy());&#xD;"
'sSrc = sSrc & vbCrLf & "                          } else {&#xD;"
'sSrc = sSrc & vbCrLf & "                              /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                              * ������������ ��ȸ&#xD;"
'sSrc = sSrc & vbCrLf & "                              *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                              DataObjectMapper&lt;HgigGrntContInq003Dvo&gt; mapper = new DataObjectMapper&lt;HgigGrntContInq003Dvo&gt;( ""po7_app"", HgigGrntContInq003Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                                  &#xD;"
'sSrc = sSrc & vbCrLf & "                              mapper.setQuery(HgigGrntContInq001Dao.SEL002);&#xD;"
'sSrc = sSrc & vbCrLf & "                                  &#xD;"
'sSrc = sSrc & vbCrLf & "                              //������ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                              mapper.setParameter(""grntNo"", grntContList.get(iGrntCont).getGrntNo());&#xD;"
'sSrc = sSrc & vbCrLf & "                              &#xD;"
'sSrc = sSrc & vbCrLf & "                              HgigGrntContInq003Dvo etcInfoVo = mapper.get(); &#xD;"
'sSrc = sSrc & vbCrLf & "                              //��������(&lt;-����������ǥ����)&#xD;"
'sSrc = sSrc & vbCrLf & "                              grnt1601Vo.setGrntDy(etcInfoVo.getExecCnclEntDy());&#xD;"
'sSrc = sSrc & vbCrLf & "                          }&#xD;"
'sSrc = sSrc & vbCrLf & "                          /* 2.������ȣ */&#xD;"
'sSrc = sSrc & vbCrLf & "                          grnt1601Vo.setGrntNo(grntContList.get(iGrntCont).getGrntNo());&#xD;"
'sSrc = sSrc & vbCrLf & "                          /* 3.���������ڵ� */&#xD;"
'sSrc = sSrc & vbCrLf & "                         grnt1601Vo.setGrntDvcd(grntContList.get(iGrntCont).getGrntDvcd());&#xD;"
'sSrc = sSrc & vbCrLf & "                         &#xD;"
'sSrc = sSrc & vbCrLf & "                         /* 4.�����ܾ� */&#xD;"
'sSrc = sSrc & vbCrLf & "                         //�ֽź�������û�����ڵ�(07:�߱�)&#xD;"
'sSrc = sSrc & vbCrLf & "                         if(""07"".equals(grntContList.get(iGrntCont).getHgGrntReqStcd())){&#xD;"
'sSrc = sSrc & vbCrLf & "                             //�����ܾ�&#xD;"
'sSrc = sSrc & vbCrLf & "                             grnt1601Vo.setGrntRamt(grntContList.get(iGrntCont).getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                         } else {&#xD;"
'sSrc = sSrc & vbCrLf & "                             //�����ܾ�(&lt;-������û�ݾ�)&#xD;"
'sSrc = sSrc & vbCrLf & "                             grnt1601Vo.setGrntRamt(grntContList.get(iGrntCont).getGrntReqAmt());&#xD;"
'sSrc = sSrc & vbCrLf & "                         }&#xD;"
'sSrc = sSrc & vbCrLf & "                         &#xD;"
'sSrc = sSrc & vbCrLf & "                         /* 5.��ä�������(&lt;-��������) */&#xD;"
'sSrc = sSrc & vbCrLf & "                         grnt1601Vo.setMdbtSbjtNm(grntContList.get(iGrntCont).getLoanSbjtNm());&#xD;"
'sSrc = sSrc & vbCrLf & "                         &#xD;"
'sSrc = sSrc & vbCrLf & "                          /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                          * �����������ȣ ��ȸ&#xD;"
'sSrc = sSrc & vbCrLf & "                          *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                          DataObjectMapper&lt;HgigGrntContInq003Dvo&gt; mapper = new DataObjectMapper&lt;HgigGrntContInq003Dvo&gt;( ""po7_app"", HgigGrntContInq003Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                              &#xD;"
'sSrc = sSrc & vbCrLf & "                          mapper.setQuery(HgigGrntContInq001Dao.SEL003);&#xD;"
'sSrc = sSrc & vbCrLf & "                              &#xD;"
'sSrc = sSrc & vbCrLf & "                          //������ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                          mapper.setParameter(""grntNo"", grntContList.get(iGrntCont).getGrntNo());&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                          HgigGrntContInq003Dvo etcInfoVo = mapper.get(); &#xD;"
'sSrc = sSrc & vbCrLf & "                         /* 6.�����������ȣ */&#xD;"
'sSrc = sSrc & vbCrLf & "                         grnt1601Vo.setTlocZpcd(etcInfoVo.getTlocZpcd());&#xD;"
'sSrc = sSrc & vbCrLf & "                         /* 7.���������ڵ�(&lt;-�������ó�����ڵ�) */&#xD;"
'sSrc = sSrc & vbCrLf & "                         grnt1601Vo.setBankGiroCd(grntContList.get(iGrntCont).getGrntCtpcGiroCd());       &#xD;"
'sSrc = sSrc & vbCrLf & "                         &#xD;"
'sSrc = sSrc & vbCrLf & "                         /* �������� �߰� */&#xD;"
'sSrc = sSrc & vbCrLf & "                         grnt1601List.add(grnt1601Vo);&#xD;"
'sSrc = sSrc & vbCrLf & "                      }</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""11"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>�� �����ܾ� ��ȸ</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                           /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                          * �����ܾ� ��ȸ&#xD;"
'sSrc = sSrc & vbCrLf & "                          *****************************************************************/&#xD;"
'sSrc = sSrc & vbCrLf & "                          DataObjectMapper&lt;HgigGrntContInq002Dvo&gt; mapper = new DataObjectMapper&lt;HgigGrntContInq002Dvo&gt;( ""po7_app"", HgigGrntContInq002Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                              &#xD;"
'sSrc = sSrc & vbCrLf & "                          mapper.setQuery(HgigGrntContInq001Dao.SELLIST004);&#xD;"
'sSrc = sSrc & vbCrLf & "                              &#xD;"
'sSrc = sSrc & vbCrLf & "                          //����ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                          mapper.setParameter(""csno"", custInfoVo.getCsno());&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                          //�������ݾ�             10.�������(����,����,����)&#xD;"
'sSrc = sSrc & vbCrLf & "                          BigDecimal bdDrctAcqrAmt = BigDecimal.ZERO;&#xD;"
'sSrc = sSrc & vbCrLf & "                          //�����ߵ��ݾ�         11.�����ߵ���&#xD;"
'sSrc = sSrc & vbCrLf & "                          BigDecimal bdDrctHlfwAmt = BigDecimal.ZERO;&#xD;"
'sSrc = sSrc & vbCrLf & "                          //���������ݾ�         12.��������&#xD;"
'sSrc = sSrc & vbCrLf & "                          BigDecimal bdDrctRentAmt = BigDecimal.ZERO;&#xD;"
'sSrc = sSrc & vbCrLf & "                          //��Ź���ݾ�         13.��Ź���(����)&#xD;"
'sSrc = sSrc & vbCrLf & "                          BigDecimal bdTrstAcqrAmt = BigDecimal.ZERO;&#xD;"
'sSrc = sSrc & vbCrLf & "                          //��Ź�ߵ��ݾ�         14.��Ź�ߵ���&#xD;"
'sSrc = sSrc & vbCrLf & "                          BigDecimal bdTrstHlfwAmt = BigDecimal.ZERO;&#xD;"
'sSrc = sSrc & vbCrLf & "                          //��Ź�����ݾ�         15.��Ź����&#xD;"
'sSrc = sSrc & vbCrLf & "                          BigDecimal bdTrstRentAmt = BigDecimal.ZERO;&#xD;"
'sSrc = sSrc & vbCrLf & "                          for(HgigGrntContInq002Dvo grntRamtVo : mapper.getForwardList()){&#xD;"
'sSrc = sSrc & vbCrLf & "                              //������Ź�����ڵ�(01:����,02:��Ź)&#xD;"
'sSrc = sSrc & vbCrLf & "                              if(""01"".equals(grntRamtVo.getDrctTrstDvcd())){&#xD;"
'sSrc = sSrc & vbCrLf & "                                  //TODO: �����ڵ� ����&#xD;"
'sSrc = sSrc & vbCrLf & "                                  //������氳���ڱ�  CodeHandler.isValidTypeCode(""T41013"", ""410003"", guarntDvcd)&#xD;"
'sSrc = sSrc & vbCrLf & "                                  if (""������氳���ڱ�"".equals(grntRamtVo.getGrntDvcd())) {&#xD;"
'sSrc = sSrc & vbCrLf & "                                      bdDrctAcqrAmt = bdDrctAcqrAmt.add(grntRamtVo.getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                                  } &#xD;"
'sSrc = sSrc & vbCrLf & "                                  //�����ߵ���  CodeHandler.isValidTypeCode(""T41013"", ""410004"", guarntDvcd)&#xD;"
'sSrc = sSrc & vbCrLf & "                                  else if (""�����ߵ���"".equals(grntRamtVo.getGrntDvcd())) {&#xD;"
'sSrc = sSrc & vbCrLf & "                                      bdDrctHlfwAmt = bdDrctAcqrAmt.add(grntRamtVo.getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                                  }&#xD;"
'sSrc = sSrc & vbCrLf & "                                  //���������ڱ�  CodeHandler.isValidTypeCode(""T41013"", ""410005"", guarntDvcd)&#xD;"
'sSrc = sSrc & vbCrLf & "                                  else if (""��������"".equals(grntRamtVo.getGrntDvcd())) {&#xD;"
'sSrc = sSrc & vbCrLf & "                                      bdDrctRentAmt = bdDrctAcqrAmt.add(grntRamtVo.getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                                  }&#xD;"
'sSrc = sSrc & vbCrLf & "                              }&#xD;"
'sSrc = sSrc & vbCrLf & "                              //��Ź�� &#xD;"
'sSrc = sSrc & vbCrLf & "                              else {&#xD;"
'sSrc = sSrc & vbCrLf & "                                  //TODO: �����ڵ� ����&#xD;"
'sSrc = sSrc & vbCrLf & "                                  //��Ź���ݾ�  CodeHandler.isValidTypeCode(""T41013"", ""410003"", guarntDvcd)&#xD;"
'sSrc = sSrc & vbCrLf & "                                  if (""��Ź���ݾ�"".equals(grntRamtVo.getGrntDvcd())) {&#xD;"
'sSrc = sSrc & vbCrLf & "                                      bdTrstAcqrAmt = bdDrctAcqrAmt.add(grntRamtVo.getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                                  } &#xD;"
'sSrc = sSrc & vbCrLf & "                                  //��Ź�ߵ��ݾ�  CodeHandler.isValidTypeCode(""T41013"", ""410004"", guarntDvcd)&#xD;"
'sSrc = sSrc & vbCrLf & "                                  else if (""��Ź�ߵ��ݾ�"".equals(grntRamtVo.getGrntDvcd())) {&#xD;"
'sSrc = sSrc & vbCrLf & "                                      bdTrstHlfwAmt = bdDrctAcqrAmt.add(grntRamtVo.getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                                  }&#xD;"
'sSrc = sSrc & vbCrLf & "                                  //��Ź�����ݾ�  CodeHandler.isValidTypeCode(""T41013"", ""410005"", guarntDvcd)&#xD;"
'sSrc = sSrc & vbCrLf & "                                  else if (""��Ź�����ݾ�"".equals(grntRamtVo.getGrntDvcd())) {&#xD;"
'sSrc = sSrc & vbCrLf & "                                      bdTrstRentAmt = bdDrctAcqrAmt.add(grntRamtVo.getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                                  }&#xD;"
'sSrc = sSrc & vbCrLf & "                              }&#xD;"
'sSrc = sSrc & vbCrLf & "                          }&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                          //�������ݾ�             10.�������(����,����,����)&#xD;"
'sSrc = sSrc & vbCrLf & "                          send1601Vo.setDrctAcqrAmt(bdDrctAcqrAmt);&#xD;"
'sSrc = sSrc & vbCrLf & "                          //�����ߵ��ݾ�         11.�����ߵ���&#xD;"
'sSrc = sSrc & vbCrLf & "                          send1601Vo.setDrctHlfwAmt(bdDrctHlfwAmt);&#xD;"
'sSrc = sSrc & vbCrLf & "                          //���������ݾ�         12.��������&#xD;"
'sSrc = sSrc & vbCrLf & "                          send1601Vo.setDrctRentAmt(bdDrctRentAmt);&#xD;"
'sSrc = sSrc & vbCrLf & "                          //��Ź���ݾ�         13.��Ź���(����)&#xD;"
'sSrc = sSrc & vbCrLf & "                          send1601Vo.setTrstAcqrAmt(bdTrstAcqrAmt);&#xD;"
'sSrc = sSrc & vbCrLf & "                          //��Ź�ߵ��ݾ�         14.��Ź�ߵ���&#xD;"
'sSrc = sSrc & vbCrLf & "                          send1601Vo.setTrstHlfwAmt(bdTrstHlfwAmt);&#xD;"
'sSrc = sSrc & vbCrLf & "                          //��Ź�����ݾ�         15.��Ź����&#xD;"
'sSrc = sSrc & vbCrLf & "                          send1601Vo.setTrstRentAmt(bdTrstRentAmt);</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"


sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""17"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>�����ܾ� ���� �ʱⰪ SET</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                       /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                      * �����ܾ� ���� �ʱⰪ SET&#xD;"
'sSrc = sSrc & vbCrLf & "                      *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                      //�������ݾ�             10.�������(����,����,����)&#xD;"
'sSrc = sSrc & vbCrLf & "                      send1601Vo.setDrctAcqrAmt(BigDecimal.ZERO);&#xD;"
'sSrc = sSrc & vbCrLf & "                      //�����ߵ��ݾ�         11.�����ߵ���&#xD;"
'sSrc = sSrc & vbCrLf & "                      send1601Vo.setDrctHlfwAmt(BigDecimal.ZERO);&#xD;"
'sSrc = sSrc & vbCrLf & "                      //���������ݾ�         12.��������&#xD;"
'sSrc = sSrc & vbCrLf & "                      send1601Vo.setDrctRentAmt(BigDecimal.ZERO);&#xD;"
'sSrc = sSrc & vbCrLf & "                      //��Ź���ݾ�         13.��Ź���(����)&#xD;"
'sSrc = sSrc & vbCrLf & "                      send1601Vo.setTrstAcqrAmt(BigDecimal.ZERO);&#xD;"
'sSrc = sSrc & vbCrLf & "                      //��Ź�ߵ��ݾ�         14.��Ź�ߵ���&#xD;"
'sSrc = sSrc & vbCrLf & "                      send1601Vo.setTrstHlfwAmt(BigDecimal.ZERO);&#xD;"
'sSrc = sSrc & vbCrLf & "                      //��Ź�����ݾ�         15.��Ź����&#xD;"
'sSrc = sSrc & vbCrLf & "                      send1601Vo.setTrstRentAmt(BigDecimal.ZERO);&#xD;"
'sSrc = sSrc & vbCrLf & "                      //�Ժ��ѱݾ�           16.�Ժ��ѱݾ�&#xD;"
'sSrc = sSrc & vbCrLf & "                      send1601Vo.setSrtpTotAmt(BigDecimal.ZERO);</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""18"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>��������� �����ܾ� ��ȸ</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                           /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                          * ��������� �����ܾ� ��ȸ&#xD;"
'sSrc = sSrc & vbCrLf & "                          *****************************************************************/&#xD;"
'sSrc = sSrc & vbCrLf & "                          DataObjectMapper&lt;HgigGrntContInq002Dvo&gt; mapper = new DataObjectMapper&lt;HgigGrntContInq002Dvo&gt;( ""po7_app"", HgigGrntContInq002Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                              &#xD;"
'sSrc = sSrc & vbCrLf & "                          mapper.setQuery(HgigGrntContInq001Dao.SELLIST005);&#xD;"
'sSrc = sSrc & vbCrLf & "                              &#xD;"
'sSrc = sSrc & vbCrLf & "                          //����ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                          mapper.setParameter(""csno"", custInfoVo.getCsno());&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                          //�������ݾ�             10.�������(����,����,����)&#xD;"
'sSrc = sSrc & vbCrLf & "                          BigDecimal bdDrctAcqrAmt = send1601Vo.getDrctAcqrAmt();&#xD;"
'sSrc = sSrc & vbCrLf & "                          //�����ߵ��ݾ�         11.�����ߵ���&#xD;"
'sSrc = sSrc & vbCrLf & "                          BigDecimal bdDrctHlfwAmt = send1601Vo.getDrctHlfwAmt();&#xD;"
'sSrc = sSrc & vbCrLf & "                          //���������ݾ�         12.��������&#xD;"
'sSrc = sSrc & vbCrLf & "                          BigDecimal bdDrctRentAmt = send1601Vo.getDrctRentAmt();&#xD;"
'sSrc = sSrc & vbCrLf & "                          //��Ź���ݾ�         13.��Ź���(����)&#xD;"
'sSrc = sSrc & vbCrLf & "                          BigDecimal bdTrstAcqrAmt = send1601Vo.getTrstAcqrAmt();&#xD;"
'sSrc = sSrc & vbCrLf & "                          //��Ź�ߵ��ݾ�         14.��Ź�ߵ���&#xD;"
'sSrc = sSrc & vbCrLf & "                          BigDecimal bdTrstHlfwAmt = send1601Vo.getTrstHlfwAmt();&#xD;"
'sSrc = sSrc & vbCrLf & "                          //��Ź�����ݾ�         15.��Ź����&#xD;"
'sSrc = sSrc & vbCrLf & "                          BigDecimal bdTrstRentAmt = send1601Vo.getTrstRentAmt();&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                          for(HgigGrntContInq002Dvo jintGrntRamtVo : mapper.getForwardList()){&#xD;"
'sSrc = sSrc & vbCrLf & "                              //������Ź�����ڵ�(01:����,02:��Ź)&#xD;"
'sSrc = sSrc & vbCrLf & "                              if(""01"".equals(jintGrntRamtVo.getDrctTrstDvcd())){&#xD;"
'sSrc = sSrc & vbCrLf & "                                  //TODO: �����ڵ� ����&#xD;"
'sSrc = sSrc & vbCrLf & "                                  //������氳���ڱ�  CodeHandler.isValidTypeCode(""T41013"", ""410003"", guarntDvcd)&#xD;"
'sSrc = sSrc & vbCrLf & "                                  if (""������氳���ڱ�"".equals(jintGrntRamtVo.getGrntDvcd())) {&#xD;"
'sSrc = sSrc & vbCrLf & "                                      bdDrctAcqrAmt = bdDrctAcqrAmt.add(jintGrntRamtVo.getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                                  } &#xD;"
'sSrc = sSrc & vbCrLf & "                                  //�����ߵ���  CodeHandler.isValidTypeCode(""T41013"", ""410004"", guarntDvcd)&#xD;"
'sSrc = sSrc & vbCrLf & "                                  else if (""�����ߵ���"".equals(jintGrntRamtVo.getGrntDvcd())) {&#xD;"
'sSrc = sSrc & vbCrLf & "                                      bdDrctHlfwAmt = bdDrctAcqrAmt.add(jintGrntRamtVo.getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                                  }&#xD;"
'sSrc = sSrc & vbCrLf & "                                  //���������ڱ�  CodeHandler.isValidTypeCode(""T41013"", ""410005"", guarntDvcd)&#xD;"
'sSrc = sSrc & vbCrLf & "                                  else if (""��������"".equals(jintGrntRamtVo.getGrntDvcd())) {&#xD;"
'sSrc = sSrc & vbCrLf & "                                      bdDrctRentAmt = bdDrctAcqrAmt.add(jintGrntRamtVo.getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                                  }&#xD;"
'sSrc = sSrc & vbCrLf & "                              }&#xD;"
'sSrc = sSrc & vbCrLf & "                              //��Ź�� &#xD;"
'sSrc = sSrc & vbCrLf & "                              else {&#xD;"
'sSrc = sSrc & vbCrLf & "                                  //TODO: �����ڵ� ����&#xD;"
'sSrc = sSrc & vbCrLf & "                                  //��Ź���ݾ�  CodeHandler.isValidTypeCode(""T41013"", ""410003"", guarntDvcd)&#xD;"
'sSrc = sSrc & vbCrLf & "                                  if (""��Ź���ݾ�"".equals(jintGrntRamtVo.getGrntDvcd())) {&#xD;"
'sSrc = sSrc & vbCrLf & "                                      bdTrstAcqrAmt = bdDrctAcqrAmt.add(jintGrntRamtVo.getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                                  } &#xD;"
'sSrc = sSrc & vbCrLf & "                                  //��Ź�ߵ��ݾ�  CodeHandler.isValidTypeCode(""T41013"", ""410004"", guarntDvcd)&#xD;"
'sSrc = sSrc & vbCrLf & "                                  else if (""��Ź�ߵ��ݾ�"".equals(jintGrntRamtVo.getGrntDvcd())) {&#xD;"
'sSrc = sSrc & vbCrLf & "                                      bdTrstHlfwAmt = bdDrctAcqrAmt.add(jintGrntRamtVo.getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                                  }&#xD;"
'sSrc = sSrc & vbCrLf & "                                  //��Ź�����ݾ�  CodeHandler.isValidTypeCode(""T41013"", ""410005"", guarntDvcd)&#xD;"
'sSrc = sSrc & vbCrLf & "                                  else if (""��Ź�����ݾ�"".equals(jintGrntRamtVo.getGrntDvcd())) {&#xD;"
'sSrc = sSrc & vbCrLf & "                                      bdTrstRentAmt = bdDrctAcqrAmt.add(jintGrntRamtVo.getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                                  }&#xD;"
'sSrc = sSrc & vbCrLf & "                              }&#xD;"
'sSrc = sSrc & vbCrLf & "                          }&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                          //�������ݾ�             10.�������(����,����,����)&#xD;"
'sSrc = sSrc & vbCrLf & "                          send1601Vo.setDrctAcqrAmt(bdDrctAcqrAmt);&#xD;"
'sSrc = sSrc & vbCrLf & "                          //�����ߵ��ݾ�         11.�����ߵ���&#xD;"
'sSrc = sSrc & vbCrLf & "                          send1601Vo.setDrctHlfwAmt(bdDrctHlfwAmt);&#xD;"
'sSrc = sSrc & vbCrLf & "                          //���������ݾ�         12.��������&#xD;"
'sSrc = sSrc & vbCrLf & "                          send1601Vo.setDrctRentAmt(bdDrctRentAmt);&#xD;"
'sSrc = sSrc & vbCrLf & "                          //��Ź���ݾ�         13.��Ź���(����)&#xD;"
'sSrc = sSrc & vbCrLf & "                          send1601Vo.setTrstAcqrAmt(bdTrstAcqrAmt);&#xD;"
'sSrc = sSrc & vbCrLf & "                          //��Ź�ߵ��ݾ�         14.��Ź�ߵ���&#xD;"
'sSrc = sSrc & vbCrLf & "                          send1601Vo.setTrstHlfwAmt(bdTrstHlfwAmt);&#xD;"
'sSrc = sSrc & vbCrLf & "                          //��Ź�����ݾ�         15.��Ź����&#xD;"
'sSrc = sSrc & vbCrLf & "                          send1601Vo.setTrstRentAmt(bdTrstRentAmt);</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"


sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""12"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>�Ժ��ѱݾ� ��ȸ</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                       /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                      * �Ժ��ѱݾ� ��ȸ&#xD;"
'sSrc = sSrc & vbCrLf & "                      *****************************************************************/&#xD;"
'sSrc = sSrc & vbCrLf & "                      DataObjectMapper&lt;HgigGrntContInq003Dvo&gt; mapper = new DataObjectMapper&lt;HgigGrntContInq003Dvo&gt;( ""po7_app"", HgigGrntContInq003Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                      mapper.setQuery(HgigGrntContInq001Dao.SEL004);&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                      //����ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                      mapper.setParameter(""csno"", custInfoVo.getCsno());&#xD;"
'sSrc = sSrc & vbCrLf & "                      &#xD;"
'sSrc = sSrc & vbCrLf & "                      HgigGrntContInq003Dvo etcInfoVo = mapper.get(); &#xD;"
'sSrc = sSrc & vbCrLf & "                      &#xD;"
'sSrc = sSrc & vbCrLf & "                      //�Ժ��ѱݾ�&#xD;"
'sSrc = sSrc & vbCrLf & "                      send1601Vo.setSrtpTotAmt(etcInfoVo.getSrtpTotAmt());                             </flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""19"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>û��������� ��û������ȸ</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                       /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                      * û��������� ��û������ȸ(MAX 10��)&#xD;"
'sSrc = sSrc & vbCrLf & "                      *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                         DataObjectMapper&lt;HgigGrntContInq002Dvo&gt; mapper = new DataObjectMapper&lt;HgigGrntContInq002Dvo&gt;( ""po7_app"", HgigGrntContInq002Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                              &#xD;"
'sSrc = sSrc & vbCrLf & "                          mapper.setQuery(HgigGrntContInq001Dao.SELLIST006);&#xD;"
'sSrc = sSrc & vbCrLf & "                              &#xD;"
'sSrc = sSrc & vbCrLf & "                          //����ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                          mapper.setParameter(""csno"", custInfoVo.getCsno());&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                          hf.cor.hg.ig.vo.bvo.comn.HgigTrstTlgmSendData1601YthBvo yth1601Vo = null;&#xD;"
'sSrc = sSrc & vbCrLf & "                          for(HgigGrntContInq002Dvo ythMmrtGrntVo : mapper.getForwardList()){&#xD;"
'sSrc = sSrc & vbCrLf & "                              yth1601Vo = new hf.cor.hg.ig.vo.bvo.comn.HgigTrstTlgmSendData1601YthBvo();&#xD;"
'sSrc = sSrc & vbCrLf & "                              //1.������ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                              yth1601Vo.setGrntNo(ythMmrtGrntVo.getGrntNo());&#xD;"
'sSrc = sSrc & vbCrLf & "                              //2.���Ҵ������ݾ� �հ�&#xD;"
'sSrc = sSrc & vbCrLf & "                              yth1601Vo.setLoanTotAmt(ythMmrtGrntVo.getLoanTotAmt());&#xD;"
'sSrc = sSrc & vbCrLf & "                              //3.�����ܾ�&#xD;"
'sSrc = sSrc & vbCrLf & "                              yth1601Vo.setGrntRamt(ythMmrtGrntVo.getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                              //4.���������ڵ�&#xD;"
'sSrc = sSrc & vbCrLf & "                              yth1601Vo.setBankGiroCd(ythMmrtGrntVo.getNvlGrntCtpcGiroCd());&#xD;"
'sSrc = sSrc & vbCrLf & "                              &#xD;"
'sSrc = sSrc & vbCrLf & "                              yth1601List.add(yth1601Vo);&#xD;"
'sSrc = sSrc & vbCrLf & "                          }</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""20"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>������ SET</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                       /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                      * �� ���� SET&#xD;"
'sSrc = sSrc & vbCrLf & "                      *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                     //1.����ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                     send1601Vo.setCsno(custInfoVo.getCsno());&#xD;"
'sSrc = sSrc & vbCrLf & "                     //2.�ֹι�ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                     send1601Vo.setJno(custInfoVo.getJcno());&#xD;"
'sSrc = sSrc & vbCrLf & "                     //3.����&#xD;"
'sSrc = sSrc & vbCrLf & "                     send1601Vo.setCustNm(custInfoVo.getCustNm());&#xD;"
'sSrc = sSrc & vbCrLf & "                     //4.������ �����ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                     send1601Vo.setZpcd(custInfoVo.getHomeZpcd());&#xD;"
'sSrc = sSrc & vbCrLf & "                     //5.������ �ּ�&#xD;"
'sSrc = sSrc & vbCrLf & "                     send1601Vo.setAddr(custInfoVo.getHomeAddr());&#xD;"
'sSrc = sSrc & vbCrLf & "                     //6.������ȭ��ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                     send1601Vo.setHomeTlno(custInfoVo.getHomeTlno());&#xD;"
'sSrc = sSrc & vbCrLf & "                     //7.�޴�����ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                     send1601Vo.setMbno(custInfoVo.getMbno());&#xD;"
'sSrc = sSrc & vbCrLf & "                     //8.������������&#xD;"
'sSrc = sSrc & vbCrLf & "                     send1601Vo.setCtrlBrNm(custInfoVo.getCtrlBrofCdNm());&#xD;"
'sSrc = sSrc & vbCrLf & "                     //9.��������������ȭ��ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                     send1601Vo.setBrTlno(custInfoVo.getBrTlno());&#xD;"
'sSrc = sSrc & vbCrLf & "                     //17.������ȣ�Ǽ�&#xD;"
'sSrc = sSrc & vbCrLf & "                     send1601Vo.setGrntCnt(grnt1601List.size());</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""3"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>������ ��ȸ</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                   /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                  * ������ ��ȸ&#xD;"
'sSrc = sSrc & vbCrLf & "                  *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                  DataObjectMapper&lt;HgigGrntContInq001Dvo&gt; mapper = new DataObjectMapper&lt;HgigGrntContInq001Dvo&gt;( ""po7_app"", HgigGrntContInq001Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                      &#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setQuery(HgigGrntContInq001Dao.SEL001);&#xD;"
'sSrc = sSrc & vbCrLf & "                      &#xD;"
'sSrc = sSrc & vbCrLf & "                  //�ֹι��ι�ȣ TODO: ��ȣȭó��&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""jcno"", inBvo.getRecvDataVo().getJno());&#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  custInfoVo = mapper.get(); &#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  //������������ ����&#xD;"
'sSrc = sSrc & vbCrLf & "                  if(custInfoVo == null){&#xD;"
'sSrc = sSrc & vbCrLf & "                      hf.cor.hg.ig.vo.bvo.comn.HgigTrstTlgmSendData1601Bvo reVo = new hf.cor.hg.ig.vo.bvo.comn.HgigTrstTlgmSendData1601Bvo();&#xD;"
'sSrc = sSrc & vbCrLf & "                      reVo.setJno(inBvo.getRecvDataVo().getJno());&#xD;"
'sSrc = sSrc & vbCrLf & "                      outBvo.setGrntContVo(reVo);&#xD;"
'sSrc = sSrc & vbCrLf & "                      return outBvo;&#xD;"
'sSrc = sSrc & vbCrLf & "                      // �ش� ���� ��ϵǾ� ���� �ʽ��ϴ�&#xD;"
'sSrc = sSrc & vbCrLf & "                      //throw new LiveException(""hg.info.00512"", pVOs);&#xD;"
'sSrc = sSrc & vbCrLf & "                  }</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""10"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>�ν��ڷ� �����ڳ��� ��ȸ</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                   /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                  * �ν��ڷ� �����ڳ��� ��ȸ(MAX 10��)&#xD;"
'sSrc = sSrc & vbCrLf & "                  *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                  DataObjectMapper&lt;HgigGrntContInq004Dvo&gt; mapper = new DataObjectMapper&lt;HgigGrntContInq004Dvo&gt;( ""po7_app"", HgigGrntContInq004Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                      &#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setQuery(HgigGrntContInq001Dao.SELLIST003);&#xD;"
'sSrc = sSrc & vbCrLf & "                      &#xD;"
'sSrc = sSrc & vbCrLf & "                  //����ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""csno"", custInfoVo.getCsno());&#xD;"
'sSrc = sSrc & vbCrLf & "                  hf.cor.hg.ig.vo.bvo.comn.HgigTrstTlgmSendData1601WeakBvo weak1601Vo = null;&#xD;"
'sSrc = sSrc & vbCrLf & "                  for(HgigGrntContInq004Dvo weakDataVo : mapper.getForwardList()){&#xD;"
'sSrc = sSrc & vbCrLf & "                      weak1601Vo = new hf.cor.hg.ig.vo.bvo.comn.HgigTrstTlgmSendData1601WeakBvo();&#xD;"
'sSrc = sSrc & vbCrLf & "                      &#xD;"
'sSrc = sSrc & vbCrLf & "                      //�������&#xD;"
'sSrc = sSrc & vbCrLf & "                      weak1601Vo.setRegDy(weakDataVo.getInptDy());&#xD;"
'sSrc = sSrc & vbCrLf & "                      //�ν��ڷ������ڵ�&#xD;"
'sSrc = sSrc & vbCrLf & "                      weak1601Vo.setWeakDataKncd1(weakDataVo.getWeakDataKncd1());&#xD;"
'sSrc = sSrc & vbCrLf & "                      //�ν��ڷ������ڵ�&#xD;"
'sSrc = sSrc & vbCrLf & "                      weak1601Vo.setWeakDataKncd2(weakDataVo.getWeakDataKncd2());&#xD;"
'sSrc = sSrc & vbCrLf & "                      //�ν��ڷ������ڵ�&#xD;"
'sSrc = sSrc & vbCrLf & "                      weak1601Vo.setWeakDataKncd3(weakDataVo.getWeakDataKncd3());&#xD;"
'sSrc = sSrc & vbCrLf & "                      //�����&#xD;"
'sSrc = sSrc & vbCrLf & "                      weak1601Vo.setOrgNm(weakDataVo.getPrsCtrlBrofCdNm());&#xD;"
'sSrc = sSrc & vbCrLf & "                      &#xD;"
'sSrc = sSrc & vbCrLf & "                      weak1601List.add(weak1601Vo);&#xD;"
'sSrc = sSrc & vbCrLf & "                  }</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "            </flow:virtualModules>"
sSrc = sSrc & vbCrLf & "            <flow:loopModules/>"
sSrc = sSrc & vbCrLf & "            <flow:bizMethodCalls/>"
sSrc = sSrc & vbCrLf & "            <flow:serviceCalls/>"
sSrc = sSrc & vbCrLf & "            <flow:assignModules/>"
sSrc = sSrc & vbCrLf & "            <flow:dataObjectCalls/>"
sSrc = sSrc & vbCrLf & "            <flow:replyCalls/>"
sSrc = sSrc & vbCrLf & "        </flow:flowModule>"
sSrc = sSrc & vbCrLf & "        <bo:arguments>"
sSrc = sSrc & vbCrLf & "            <flow:argument>"


iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO������ & """ comments=""0'th argument"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO�� & """ resourcePath=""hf/cor/hg/ig/vo/bvo/sbvo""/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"
sSrc = sSrc & vbCrLf & "                <flow:sequence>0</flow:sequence>"
sSrc = sSrc & vbCrLf & "            </flow:argument>"
sSrc = sSrc & vbCrLf & "        </bo:arguments>"

'vofile ����
Call subMakeVO

iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "        <bo:returnType kind=""DTO"">"
sSrc = sSrc & vbCrLf & "            <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO�� & """ resourcePath=""" & sVO��� & """/>"
sSrc = sSrc & vbCrLf & "        </bo:returnType>"
sSrc = sSrc & vbCrLf & "        <bo:classInfo classPackageName=""" & sPackageSlashName & """ className=""" & sClassPname & """/>"
sSrc = sSrc & vbCrLf & "        <bo:annotations/>"
sSrc = sSrc & vbCrLf & "    </bo:bizMethod>"

'vofile ����
Call subMakeVO

'----------------------------------------------------
'STEP5 �۽ų�����ȸ
'----------------------------------------------------
iROW_CNT = iROW_CNT + 1: subSetMehod

sSrc = sSrc & vbCrLf & "    <bo:bizMethod methodName=""" & s�޼ҵ�id & """ "
sSrc = sSrc & "displayName=""" & s�޼ҵ�� & """ modifier=""public"" "
sSrc = sSrc & "comments=""@�޼���� : " & s�޼ҵ�id & "&#xD;&#xA;"
sSrc = sSrc & "@���� : " & s�޼ҵ�� & "&#xD;&#xA; "
sSrc = sSrc & "@�Է� : " & sINBVO & "&#xD;&#xA; "
sSrc = sSrc & "@��� : " & sOUTBVO & """> "
sSrc = sSrc & vbCrLf & "        <flow:flowModule>"
sSrc = sSrc & vbCrLf & "            <flow:entryModule>"
sSrc = sSrc & vbCrLf & "                <flow:name>" & s�޼ҵ�� & "</flow:name>"
sSrc = sSrc & vbCrLf & "                <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:id>0</flow:id>"
sSrc = sSrc & vbCrLf & "                    <flow:id>2</flow:id>"
sSrc = sSrc & vbCrLf & "                    <flow:id>4</flow:id>"
sSrc = sSrc & vbCrLf & "                </flow:nextEntries>"



'VO
iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO������ & """ scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO�� & """ resourcePath=""" & sVO��� & """/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"

'vofile ����
Call subMakeVO

'VO
iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO������ & """ scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO�� & """ resourcePath=""" & sVO��� & """/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"

'vofile ����
Call subMakeVO

'VO
iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO������ & """ scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO�� & """ resourcePath=""" & sVO��� & """/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"

'vofile ����
Call subMakeVO


sSrc = sSrc & vbCrLf & "            </flow:entryModule>"
sSrc = sSrc & vbCrLf & "            <flow:innerModules>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""0"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>��������ó��</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>1</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""2"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>���� �� ó��</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>3</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""4"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>��������ó��</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>5</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "            </flow:innerModules>"
sSrc = sSrc & vbCrLf & "            <flow:xorModules/>"

sSrc = sSrc & vbCrLf & "            <flow:virtualModules>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""1"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>�Է°� Ȯ��</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                  logger.info(""  �١ڡ١ڡ١ڡ١�inBvo  " & s�޼ҵ�id & " = \n[""+inBvo.toString()+""] �١ڡ١ڡ١ڡ١�"" );&#xD;"
'sSrc = sSrc & vbCrLf & "                 &#xD;"
'sSrc = sSrc & vbCrLf & "                 if(inBvo.getInqComnVo() == null){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""��ȸ����"" + ""��(��) �ʼ��Է��Դϴ�."");&#xD;"
'sSrc = sSrc & vbCrLf & "                 }&#xD;"
'sSrc = sSrc & vbCrLf & "                 //��������&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(StringUtil.isEmpty(inBvo.getInqComnVo().getRecvDy())){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""��������"" + ""��(��) �ʼ��Է��Դϴ�."");&#xD;"
'sSrc = sSrc & vbCrLf & "                     &#xD;"
'sSrc = sSrc & vbCrLf & "                 }&#xD;"
'sSrc = sSrc & vbCrLf & "                 //���Žð�&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(StringUtil.isEmpty(inBvo.getInqComnVo().getRecvPtm())){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""���Žð�"" + ""��(��) �ʼ��Է��Դϴ�."");&#xD;"
'sSrc = sSrc & vbCrLf & "                 }&#xD;"
'sSrc = sSrc & vbCrLf & "                 //��������ڵ�&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(StringUtil.isEmpty(inBvo.getInqComnVo().getOrgGiroCd())){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""��������ڵ�"" + ""��(��) �ʼ��Է��Դϴ�."");&#xD;"
'sSrc = sSrc & vbCrLf & "                 }&#xD;"
'sSrc = sSrc & vbCrLf & "                 //��Ź�����ڵ�&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(StringUtil.isEmpty(inBvo.getInqComnVo().getTrstTlgmCd())){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""��Ź�����ڵ�"" + ""��(��) �ʼ��Է��Դϴ�."");&#xD;"
'sSrc = sSrc & vbCrLf & "                 }&#xD;"
'sSrc = sSrc & vbCrLf & "                 //����������ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(StringUtil.isEmpty(inBvo.getInqComnVo().getTlgmMngNo())){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""����������ȣ"" + ""��(��) �ʼ��Է��Դϴ�."");&#xD;"
'sSrc = sSrc & vbCrLf & "                 }&#xD;"
'sSrc = sSrc & vbCrLf & "                 //��������&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(inBvo.getInqComnVo().getTlgmSeq() &lt;= 0){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""��������"" + ""��(��) �ʼ��Է��Դϴ�."");&#xD;"
'sSrc = sSrc & vbCrLf & "                 }</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""3"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>�۽ų��� ��ȸ</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                   /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                  * �۽ų��� ��ȸ&#xD;"
'sSrc = sSrc & vbCrLf & "                  *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                  DataObjectMapper&lt;TbHgc602lHstrstsend001Dvo&gt; mapper = new DataObjectMapper&lt;TbHgc602lHstrstsend001Dvo&gt;( ""po7_app"", TbHgc602lHstrstsend001Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setQuery(TbHgc602lHstrstsend001Dao.SEL001);&#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  //��������&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""recvDy"", inBvo.getInqComnVo().getRecvDy());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //���Žð�&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""recvPtm"", inBvo.getInqComnVo().getRecvPtm());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //��������ڵ�&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""orgGiroCd"", inBvo.getInqComnVo().getOrgGiroCd());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //��Ź�����ڵ�&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""trstTlgmCd"", inBvo.getInqComnVo().getTrstTlgmCd());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //����������ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""tlgmMngNo"", inBvo.getInqComnVo().getTlgmMngNo());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //��������&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""tlgmSeq"", inBvo.getInqComnVo().getTlgmSeq());&#xD;"
'sSrc = sSrc & vbCrLf & "&#xD;"
'sSrc = sSrc & vbCrLf & "                  TbHgc602lHstrstsend001Dvo reDvo = mapper.get();&#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  if(reDvo == null){&#xD;"
'sSrc = sSrc & vbCrLf & "                      //TODO:&#xD;"
'sSrc = sSrc & vbCrLf & "                      throw new UserException(""�۽ų��� ��ȸ����� �����ϴ�."");&#xD;"
'sSrc = sSrc & vbCrLf & "                  }&#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                  * HEAD ����&#xD;"
'sSrc = sSrc & vbCrLf & "                  *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                  //��������ڵ� &#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setOrgGiroCd(reDvo.getOrgGiroCd());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //�۽����� &#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setSendDy(reDvo.getSendDy());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //�۽Žð� &#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setSendPtm(reDvo.getSendPtm());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //��Ź�����   &#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setCnsgOrgNm(reDvo.getCnsgOrgNm());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //����������ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setTlgmMngNo(reDvo.getTlgmMngNo());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //��Ź�����ڵ�&#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setTrstTlgmCd(reDvo.getTrstTlgmCd());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //��Ź���������ڵ�&#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setTrstTlgmRespCd(reDvo.getTrstTlgmRespCd());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //�������μ���&#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setErrLineSeq(reDvo.getErrLineSeq());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //�������α׷�����&#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setErrPgrmCont(reDvo.getErrPgrmCont());&#xD;"
'sSrc = sSrc & vbCrLf & "&#xD;"
'sSrc = sSrc & vbCrLf & "                  /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                  * DATA ����&#xD;"
'sSrc = sSrc & vbCrLf & "                  *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                  String sSendData = reDvo.getSendDataContEnc();&#xD;"
'sSrc = sSrc & vbCrLf & "                  byte[] btSendData = ByteUtil.toBytes(sSendData, ""utf-8"");&#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  logger.info(""  �١ڡ١ڡ١ڡ١� sSendData = \n[""+ sSendData +""]\n �١ڡ١ڡ١ڡ١�"" );&#xD;"
'sSrc = sSrc & vbCrLf & "                  logger.info(""  �١ڡ١ڡ١ڡ١� btSendData = \n[""+ btSendData +""]\n �١ڡ١ڡ١ڡ١�"" );&#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  int iStrtCnt = 0;&#xD;"
'sSrc = sSrc & vbCrLf & "                  String sData = null;&#xD;"
'sSrc = sSrc & vbCrLf & "&#xD;"
'sSrc = sSrc & vbCrLf & "                  if(sSendData != null){&#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                      for (int i1602 = 1; i1602 &lt; HgIgCommConst.TRST_TLGM_CD_1602.length; i1602++) {&#xD;"
'sSrc = sSrc & vbCrLf & "                      &#xD;"
'sSrc = sSrc & vbCrLf & "                          sData = new String(ByteUtil.substring(btSendData, iStrtCnt, iStrtCnt + HgIgCommConst.TRST_TLGM_CD_1602[i1602]));&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                          logger.info(""i1602 =&gt; ]"" + iStrtCnt + "" / "" + (iStrtCnt + HgIgCommConst.TRST_TLGM_CD_1602[i1602]) + ""]"");&#xD;"
'sSrc = sSrc & vbCrLf & "                          logger.info(""size =&gt; "" + HgIgCommConst.TRST_TLGM_CD_1602[i1602]);&#xD;"
'sSrc = sSrc & vbCrLf & "                          logger.info(""sData =&gt; ["" + sData + ""]"");&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                          iStrtCnt = iStrtCnt + HgIgCommConst.TRST_TLGM_CD_1602[i1602];&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                          //TODO: byteUtil ���� ���� �� �ּ�����&#xD;"
'sSrc = sSrc & vbCrLf & "                          //TODO: MMS���� �� �������� ������ �Ľ��ϵ��� ����&#xD;"
'sSrc = sSrc & vbCrLf & "                          if(i1602 == 1){&#xD;"
'sSrc = sSrc & vbCrLf & "                              dataVo.setCsno(sData);                 //����ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                          } else if(i1602 == 2){&#xD;"
'sSrc = sSrc & vbCrLf & "                              dataVo.setJno(sData);                  //�ֹι�ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "                          } else if(i1602 == 3){&#xD;"
'sSrc = sSrc & vbCrLf & "                              dataVo.setCustNm(sData);               //����&#xD;"
'sSrc = sSrc & vbCrLf & "                          } &#xD;"
'sSrc = sSrc & vbCrLf & "//                         else if(i1602 == 4){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setZpcd(sData);                 //�����ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 5){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setAddr(sData);                 //�ּ�&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 6){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setHomeTlno(sData);             //������ȭ��ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 7){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setMbno(sData);                 //�޴�����ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 8){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setCtrlBrNm(sData);             //���������&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 9){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setBrTlno(sData);               //������ȭ��ȣ&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 10){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setDrctAcqrAmt(sData);           //�������ݾ�&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 11){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setDrctHlfwAmt(sData);           //�����ߵ��ݾ�&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 12){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setDrctRentAmt(sData);           //���������ݾ�&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 13){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setTrstAcqrAmt(sData);           //��Ź���ݾ�&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 14){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setTrstHlfwAmt(sData);           //��Ź�ߵ��ݾ�&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 15){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setTrstRentAmt(sData);           //��Ź�����ݾ�&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 16){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setSrtpTotAmt(sData);            //�Ժ��ѱݾ�&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 17){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setGrntCnt(sData);               //�����Ǽ�&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 18){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setReptCnt(sData);               //�ݺ��Ǽ�&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 19){&#xD;"
'sSrc = sSrc & vbCrLf & "                         for(int iGrntCnt = 0; iGrntCnt &lt; HgIgCommConst.TRST_TLGM_CD_1602[i1602]; iGrntCnt++){&#xD;"
'sSrc = sSrc & vbCrLf & "                             &#xD;"
'sSrc = sSrc & vbCrLf & "                             &#xD;"
'sSrc = sSrc & vbCrLf & "                             &#xD;"
'sSrc = sSrc & vbCrLf & "                             &#xD;"
'sSrc = sSrc & vbCrLf & "                         }&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 20){&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 21){&#xD;"
'sSrc = sSrc & vbCrLf & "//                        }&#xD;"
'sSrc = sSrc & vbCrLf & "                     }&#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  }</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""5"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>return</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                   outBvo.setSendHdrVo(hdrVo);&#xD;"
'sSrc = sSrc & vbCrLf & "                  outBvo.setSendDataVo(dataVo);&#xD;"
'sSrc = sSrc & vbCrLf & "                  logger.info(""  �١ڡ١ڡ١ڡ١�outBvo  selSendCont1602 = \n [""+outBvo.toString()+""] �١ڡ١ڡ١ڡ١�"" );&#xD;"
'sSrc = sSrc & vbCrLf & "                  return outBvo;</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"
sSrc = sSrc & vbCrLf & "            </flow:virtualModules>"
sSrc = sSrc & vbCrLf & "            <flow:loopModules/>"
sSrc = sSrc & vbCrLf & "            <flow:bizMethodCalls/>"
sSrc = sSrc & vbCrLf & "            <flow:serviceCalls/>"
sSrc = sSrc & vbCrLf & "            <flow:assignModules/>"
sSrc = sSrc & vbCrLf & "            <flow:dataObjectCalls/>"
sSrc = sSrc & vbCrLf & "            <flow:replyCalls/>"
sSrc = sSrc & vbCrLf & "        </flow:flowModule>"
sSrc = sSrc & vbCrLf & "        <bo:arguments>"
sSrc = sSrc & vbCrLf & "          <flow:argument>"

iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO������ & """ comments=""0'th argument"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO�� & """ resourcePath=""hf/cor/hg/ig/vo/bvo/sbvo""/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"
sSrc = sSrc & vbCrLf & "                <flow:sequence>0</flow:sequence>"
sSrc = sSrc & vbCrLf & "            </flow:argument>"
sSrc = sSrc & vbCrLf & "        </bo:arguments>"


iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "        <bo:returnType kind=""DTO"">"
sSrc = sSrc & vbCrLf & "            <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO�� & """ resourcePath=""" & sVO��� & """/>"
sSrc = sSrc & vbCrLf & "        </bo:returnType>"
sSrc = sSrc & vbCrLf & "        <bo:classInfo classPackageName=""" & sPackageSlashName & """ className=""" & sClassPname & """/>"
sSrc = sSrc & vbCrLf & "        <bo:annotations/>"
sSrc = sSrc & vbCrLf & "    </bo:bizMethod>"

'vofile ����
Call subMakeVO

sSrc = sSrc & vbCrLf & "    <bo:member/>"
sSrc = sSrc & vbCrLf & "    <bo:interfaces/>"
sSrc = sSrc & vbCrLf & "    <bo:wholeImport>"
sSrc = sSrc & vbCrLf & "        <bo:importType>java.math.BigDecimal</bo:importType>"
sSrc = sSrc & vbCrLf & "        <bo:importType>java.util.ArrayList</bo:importType>"
sSrc = sSrc & vbCrLf & "        <bo:importType>java.util.List</bo:importType>"
sSrc = sSrc & vbCrLf & "        <bo:importType>com.tmax.proobject.core.BizObject</bo:importType>"
sSrc = sSrc & vbCrLf & "        <bo:importType>com.tmax.proobject.dataobject.mapper.DataObjectMapper</bo:importType>"
sSrc = sSrc & vbCrLf & "        <bo:importType>com.tmax.proobject.logger.ProObjectLogger</bo:importType>"
sSrc = sSrc & vbCrLf & "        <bo:importType>com.tmax.proobject.logger.application.ServiceLogger</bo:importType>"
sSrc = sSrc & vbCrLf & "        <bo:importType>com.tmax.proobject.model.business.BusinessObject</bo:importType>"
sSrc = sSrc & vbCrLf & "        <bo:importType>hf.cor.co.common.util.ByteUtil</bo:importType>"
sSrc = sSrc & vbCrLf & "        <bo:importType>hf.cor.co.common.util.DateUtil</bo:importType>"
sSrc = sSrc & vbCrLf & "        <bo:importType>hf.cor.co.common.util.StringUtil</bo:importType>"
sSrc = sSrc & vbCrLf & "        <bo:importType>hf.cor.hg.ig.biz.ibiz.HgigTrstTlgmMngBiz</bo:importType>"
sSrc = sSrc & vbCrLf & "        <bo:importType>hf.cor.hg.ig.common.HgIgCommConst</bo:importType>"
sSrc = sSrc & vbCrLf & "        <bo:importType>hf.cor.hg.ig.dao.sdao.HgigGrntContInq001Dao</bo:importType>"
sSrc = sSrc & vbCrLf & "        <bo:importType>hf.cor.hg.ig.dao.tdao.TbHgc601lHstrstrecv001Dao</bo:importType>"
sSrc = sSrc & vbCrLf & "        <bo:importType>hf.cor.hg.ig.dao.tdao.TbHgc602lHstrstsend001Dao</bo:importType>"
sSrc = sSrc & vbCrLf & "        <bo:importType>hf.cor.hg.ig.vo.bvo.ibvo.HgigTrstTlgmMng01InBvo</bo:importType>"
sSrc = sSrc & vbCrLf & "        <bo:importType>hf.cor.hg.ig.vo.bvo.ibvo.HgigTrstTlgmMng02InBvo</bo:importType>"
sSrc = sSrc & vbCrLf & "        <bo:importType>hf.cor.hg.ig.vo.bvo.ibvo.HgigTrstTlgmMng03InBvo</bo:importType>"
sSrc = sSrc & vbCrLf & "        <bo:importType>hf.cor.hg.ig.vo.bvo.sbvo.HgigGrntContInq04InBvo</bo:importType>"
sSrc = sSrc & vbCrLf & "        <bo:importType>hf.cor.hg.ig.vo.dvo.sdvo.HgigGrntContInq001Dvo</bo:importType>"
sSrc = sSrc & vbCrLf & "        <bo:importType>hf.cor.hg.ig.vo.dvo.sdvo.HgigGrntContInq002Dvo</bo:importType>"
sSrc = sSrc & vbCrLf & "        <bo:importType>hf.cor.hg.ig.vo.dvo.sdvo.HgigGrntContInq003Dvo</bo:importType>"
sSrc = sSrc & vbCrLf & "        <bo:importType>hf.cor.hg.ig.vo.dvo.sdvo.HgigGrntContInq004Dvo</bo:importType>"
sSrc = sSrc & vbCrLf & "        <bo:importType>hf.cor.hg.ig.vo.dvo.tdvo.TbHgc601lHstrstrecv001Dvo</bo:importType>"
sSrc = sSrc & vbCrLf & "        <bo:importType>hf.cor.hg.ig.vo.dvo.tdvo.TbHgc602lHstrstsend001Dvo</bo:importType>"
sSrc = sSrc & vbCrLf & "        <bo:importType>hf2.cor.fw.sys.util.UserException</bo:importType>"
sSrc = sSrc & vbCrLf & "    </bo:wholeImport>"
sSrc = sSrc & vbCrLf & "    <bo:typeParameters/>"
sSrc = sSrc & vbCrLf & "</bo:bizObject>"
sSrc = sSrc & vbCrLf & ""
                                                                                                                                                                                                                                                               
End With
sMakeSrcString = sSrc
End Function

Private Sub cmdMakeSrcSVC_Click()
On Error GoTo Err_Handler
With fpSrcSVC
    Dim stmFileControl As Stream
    Dim sPrintString As String
    
    
        
    iROW_CNT = 1
    
    
    Dim sFilename As String
    iROW_CNT = iROW_CNT + 1
        
LOOP_START:
    
    
    
    Set stmFileControl = New Stream
    
    
    stmFileControl.Open
    stmFileControl.Position = 0
    stmFileControl.Charset = "UTF-8"
    
    
    'file name ����
    .Row = iROW_CNT: .Col = 2
    sFilename = .Value & ".so"
    
    .Col = 1: sClassLname = .Value
    .Col = 2: sClassPname = .Value
    .Col = 3: sPackageDotName = .Value
    .Col = 4: sUserId = .Value
    .Col = 5: sUserName = .Value
    .Col = 6: sRevision = .Value
    .Col = 7: sPackageSlashName = .Value
    
    
    
    '����� ���ڿ� ����
    sPrintString = sMakeSrcStringSVC
    
    
    'stmFileControl.WriteText (sPrintString)
    'stmFileControl.SaveToFile "c:\src\" & sFileName, adSaveCreateOverWrite
    'stmFileControl.Close
            
    'Set stmFileControl = Nothing
                    
                    
    Dim objStreamUTF8 As Stream
    Set objStreamUTF8 = New Stream
    
    
    Dim objStreamUTF8NoBOM As Stream
    Set objStreamUTF8NoBOM = New Stream
    
    
    With objStreamUTF8
        .Charset = "UTF-8"
        .Open
        .WriteText sPrintString
        .Position = 0
        .Type = adTypeText
        .Position = 3
    End With
    
    With objStreamUTF8NoBOM
        .Type = adTypeBinary
        .Open
        objStreamUTF8.CopyTo objStreamUTF8NoBOM
        .SaveToFile "c:\src\" & sFilename, adSaveCreateOverWrite
        
    End With
    
    objStreamUTF8.Close
    objStreamUTF8NoBOM.Close
    
    Set objStreamUTF8 = Nothing
    Set objStreamUTF8NoBOM = Nothing
                    
                    
                    
                    
                    
                    
    .Row = iROW_CNT
    .Col = 1
    
    iROW_CNT = iROW_CNT + 1
    .Row = iROW_CNT
    '���� �ƴϸ� �ݺ��Ѵ�.
    If UCase(Trim(.Value)) <> "END" Then
        iROW_CNT = iROW_CNT + 1
        GoTo LOOP_START
    End If
End With

    gfMsgInf "�Ϸ�Ǿ����ϴ�."
    Exit Sub
Err_Handler:
    gfMsgWrn (Err.Description)
End Sub
'����� BIZ ����
Private Sub Command1_Click()
On Error GoTo Err_Handler
With fpBG
    Dim stmFileControl As Stream
    Dim sPrintString As String
    
    
        
    iROW_CNT = 0
    
    
    Dim sFilename As String
    iROW_CNT = iROW_CNT + 1
        
LOOP_START:
    
    
    
    Set stmFileControl = New Stream
    
    
    stmFileControl.Open
    stmFileControl.Position = 0
    stmFileControl.Charset = "UTF-8"
    
    
    
    .Row = iROW_CNT
    .Col = 1: s����ڼ���ID = .Value
    .Col = 2: s�����Biz��� = .Value
    s�����BizID = .Value & "Biz"
    .Col = 3: s�����SVC�� = .Value
    

    
    'file name ����
    sFilename = "Hgbg" & s�����BizID & ".bo"
    
    
    '����� ���ڿ� ����
    sPrintString = sMake�����biz
    
    
    'stmFileControl.WriteText (sPrintString)
    'stmFileControl.SaveToFile "c:\src\" & sFileName, adSaveCreateOverWrite
    'stmFileControl.Close
            
    'Set stmFileControl = Nothing
                    
                    
    Dim objStreamUTF8 As Stream
    Set objStreamUTF8 = New Stream
    
    
    Dim objStreamUTF8NoBOM As Stream
    Set objStreamUTF8NoBOM = New Stream
    
    
    With objStreamUTF8
        .Charset = "UTF-8"
        .Open
        .WriteText sPrintString
        .Position = 0
        .Type = adTypeText
        .Position = 3
    End With
    
    With objStreamUTF8NoBOM
        .Type = adTypeBinary
        .Open
        objStreamUTF8.CopyTo objStreamUTF8NoBOM
        .SaveToFile "c:\src\" & sFilename, adSaveCreateOverWrite
        
    End With
    
    objStreamUTF8.Close
    objStreamUTF8NoBOM.Close
    
    Set objStreamUTF8 = Nothing
    Set objStreamUTF8NoBOM = Nothing
                    
                    
                    
                    
                    
                    
    .Row = iROW_CNT
    .Col = 1
    
    iROW_CNT = iROW_CNT + 1
    .Row = iROW_CNT
    '���� �ƴϸ� �ݺ��Ѵ�.
    If UCase(Trim(.Value)) <> "END" Then
        'iROW_CNT = iROW_CNT + 1
        GoTo LOOP_START
    End If
End With

    gfMsgInf "�Ϸ�Ǿ����ϴ�."
    Exit Sub
Err_Handler:
    gfMsgWrn (Err.Description)
End Sub

Private Sub Command2_Click()
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
    
        sSQL = ""
        sSQL = sSQL & vbCrLf & "SELECT  nvl(ȭ��ID, 'X') as ȭ��ID,�䱸����ID                   "
        sSQL = sSQL & vbCrLf & "  FROM �䱸�������� "
        'sSQL = sSQL & vbCrLf & " WHERE NO =  25 "


    rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

    fpRequest.MaxRows = 0
    
    Dim iCnt As Integer
    
    iCnt = 0
    
    
    
    
    
    Dim sVar() As String
    Dim i, j, k As Integer
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
        
        sVar = Split(rsTemp.Fields("ȭ��ID"), Chr(10))
        
        For j = LBound(sVar) To UBound(sVar)
            fpRequest.MaxRows = fpRequest.MaxRows + 1
            fpRequest.Row = fpRequest.MaxRows
            
            fpRequest.Col = 1
            
            fpRequest.Value = rsTemp.Fields("�䱸����ID")
            fpRequest.Col = 2
            fpRequest.Value = sVar(j)
        
        Next j
        
        
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

Private Sub Form_Load()
    Me.Width = 23415
    Me.Height = 12840
End Sub

Private Function setArray(iRowCnt As Integer) As Variant
    Dim i As Integer
    Dim sArr(10) As String
    For i = 1 To 10
        With fpSrc
            .Row = iRowCnt
            .Col = i
            sArr(i) = .Value
        End With
    Next i
    
    setArray = sArr
End Function

'�޼ҵ� ���� ���� ���� �ϱ�
Private Sub subSetMehod()
    With fpSrc
        .Row = iROW_CNT
        .Col = 2: s�޼ҵ�id = .Value
        .Col = 3: s�޼ҵ�� = .Value
        .Col = 4: sINBVO = .Value
        .Col = 5: sOUTBVO = .Value
    End With
End Sub

Private Sub subSetVO()
    With fpSrc
        .Row = iROW_CNT
        .Col = 3: sVO������ = .Value
        .Col = 4: sVOid = .Value
        .Col = 5: sVO�� = .Value
        .Col = 6: sVO��� = .Value
    End With
End Sub

Private Sub subSetSVCVO()
    With fpSrcSVC
        .Row = iROW_CNT
        .Col = 2: sSVC�޼ҵ�id = .Value
        .Col = 3: sSVCInVOid = .Value
        .Col = 4: sSVCOutVOid = .Value

    End With
End Sub



Private Sub subMakeVO()


    Dim sSrc As String
    
    
    Dim sPname, sLname, sPath As String
    sPname = sVOid
    sLname = sVO��
    sPath = sVO���

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
sSrc = sSrc & "creator=""" & sUserId & """ "
sSrc = sSrc & "owner=""" & sUserId & """ "
sSrc = sSrc & "revision=""1"" "
sSrc = sSrc & "description="""" "
sSrc = sSrc & "resourcePath=""" & sPath & """ "
sSrc = sSrc & "modifier=""" & sUserId & """ "
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
        .SaveToFile "c:\src\" & sPname & ".dto", adSaveCreateOverWrite
        
    End With
    
    objStreamUTF8.Close
    objStreamUTF8NoBOM.Close
    
    Set objStreamUTF8 = Nothing
    Set objStreamUTF8NoBOM = Nothing
                    



End Sub





Private Function sMakeSrcStringSVC() As String
Dim sSrc As String
Dim sMethodNm As String



With fpSrcSVC
sSrc = ""
sSrc = sSrc & "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
sSrc = sSrc & vbCrLf & "<ns13:serviceObject xmlns:ns2=""http://www.tmax.co.kr/proobject/message"" xmlns:ns3=""http://www.tmax.co.kr/proobject/resource"" "
sSrc = sSrc & "xmlns:ns4=""http://www.tmax.co.kr/proobject/dto"" xmlns:ns5=""http://www.tmax.co.kr/proobject/dto/validator"" "
sSrc = sSrc & "xmlns:ns6=""http://www.tmax.co.kr/proobject/testcase"" xmlns:ns7=""http://www.tmax.co.kr/proobject/serviceobject-automatic"" "
sSrc = sSrc & "xmlns:ns8=""http://www.tmax.co.kr/proobject/sourcecode"" xmlns:ns9=""http://www.tmax.co.kr/proobject/flow"" "
sSrc = sSrc & "xmlns:ns10=""http://www.tmax.co.kr/proobject/dataobjectfactory"" xmlns:ns11=""http://www.tmax.co.kr/proobject/queryobject"" "
sSrc = sSrc & "xmlns:ns12=""http://www.tmax.co.kr/proobject/bizobject"" xmlns:ns13=""http://www.tmax.co.kr/proobject/serviceobject"" "
sSrc = sSrc & "xmlns:ns14=""http://www.tmax.co.kr/proobject/dataobject"" xmlns:ns15=""http://www.tmax.co.kr/proobject/taskobject"" "
sSrc = sSrc & "xmlns:ns16=""http://www.tmax.co.kr/proobject/jobobject"" xmlns:ns17=""http://www.tmax.co.kr/proobject/servicegroup"" "
sSrc = sSrc & "xmlns:ns18=""http://www.tmax.co.kr/proobject/aspectj"" xmlns:ns19=""http://www.tmax.co.kr/proobject/mapping"" "
sSrc = sSrc & "xmlns:ns20=""http://www.tmax.co.kr/proobject/application/runtime"" xmlns:ns21=""http://www.tmaxsoft.co.kr/proobject/testsuite"" "
sSrc = sSrc & "xmlns:ns22=""http://www.tmaxsoft.co.kr/proobject/testOperation"" xmlns:ns23=""http://www.tmax.co.kr/proobject/restriction_codes"" "
sSrc = sSrc & "xmlns:ns24=""http://www.example.org/externalObjectConfig"" xmlns:ns25=""http://www.tmax.co.kr/proobject/application"" "
sSrc = sSrc & "xmlns:ns26=""http://www.tmax.co.kr/proobject/contents"" xmlns:ns27=""http://www.example.org/testData"" "
sSrc = sSrc & "xmlns:ns28=""http://www.tmaxsoft.co.kr/proobject/testresult"" xmlns:ns29=""http://www.tmax.co.kr/proobject/siteConfig"" "
sSrc = sSrc & "xmlns:ns30=""http://www.tmax.co.kr/proobject/serverConfig"" xmlns:ns31=""http://www.tmax.co.kr/proobject/probuilder_config"" "
sSrc = sSrc & "xmlns:ns32=""http://www.tmax.co.kr/proobject/property"" xmlns:ns33=""http://www.tmax.co.kr/proobject/serviceGroup"" "
sSrc = sSrc & "identifier=""general"" "
sSrc = sSrc & "name=""" & sClassPname & """ "
sSrc = sSrc & "resourceId=""" & sClassPname & """ "
sSrc = sSrc & "logicalName=""" & sClassLname & """ "
sSrc = sSrc & "physicalName=""" & sClassPname & """ "
sSrc = sSrc & "resourceType=""SERVICE_OBJECT"" "
sSrc = sSrc & "resourcePackage=""" & sPackageDotName & """ "
sSrc = sSrc & "creator=""" & sUserId & """ "
sSrc = sSrc & "revision=""1"" "
sSrc = sSrc & "description=""@���ϸ� : " & sClassPname & "&#xD;&#xA;"
sSrc = sSrc & "@���� : " & sClassLname & "&#xD;&#xA;"
sSrc = sSrc & "@�ۼ��� : " & sUserName & "(" & sUserId & ")&#xD;&#xA;--------------------------------------------------------------------&#xD;&#xA;"
sSrc = sSrc & "�������� : ������ : ��û��(SR ��ȣ) : ������ �޼ҵ� : ��������&#xD;&#xA;--------------------------------------------------------------------&#xD;&#xA;"
sSrc = sSrc & "�ۼ����� : 2020.03.19 : �ۼ��� : " & sUserName & "(" & sUserId & "),  "
sSrc = sSrc & "�ۼ����� : �����ۼ�"" "
sSrc = sSrc & "resourcePath=""" & sPackageSlashName & """ "
sSrc = sSrc & "modifier=""ns043"" "
sSrc = sSrc & "created=""2020-03-20 11:00:34.14"" "
sSrc = sSrc & "lastUpdated=""2020-03-30 14:06:29"">"

sSrc = sSrc & vbCrLf & "    <ns13:inputDo xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:type=""ns14:dataObjectType"" "
sSrc = sSrc & "resourceId=""urn:proobject:" & sPackageDotName & ":" & sClassPname & """ "
sSrc = sSrc & "physicalName=""DummyVo"" "
sSrc = sSrc & "resourceType=""DATA_OBJECT"" "
sSrc = sSrc & "serviceGroupName=""hf2_cor_service"" "
sSrc = sSrc & "resourcePackage=""hf.cor._templet.vo"" "
sSrc = sSrc & "revision=""0"" "
sSrc = sSrc & "resourcePath=""hf/cor/_templet/vo""/>"


sSrc = sSrc & vbCrLf & "    <ns13:outputDo xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:type=""ns14:dataObjectType"" "
sSrc = sSrc & "resourceId=""urn:proobject:" & sPackageDotName & ":" & sClassPname & """ "
sSrc = sSrc & "physicalName=""DummyVo"" "
sSrc = sSrc & "resourceType=""DATA_OBJECT"" "
sSrc = sSrc & "serviceGroupName=""hf2_cor_service"" "
sSrc = sSrc & "resourcePackage=""hf.cor._templet.vo"" "
sSrc = sSrc & "revision=""0"" "
sSrc = sSrc & "resourcePath=""hf/cor/_templet/vo""/>"


sSrc = sSrc & vbCrLf & "    <ns13:opCodeEnable>true</ns13:opCodeEnable>"
sSrc = sSrc & vbCrLf & "    <ns13:opCode>"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "        <ns13:opCode opCode=""" & sSVC�޼ҵ�id & """ "
sSrc = sSrc & "svcName=""" & sSVC�޼ҵ�id & """ "
sSrc = sSrc & "inputDO=""" & sSVCInVOid & """ "
sSrc = sSrc & "outputDO=""" & sSVCOutVOid & """/>"


iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "        <ns13:opCode opCode=""" & sSVC�޼ҵ�id & """ "
sSrc = sSrc & "svcName=""" & sSVC�޼ҵ�id & """ "
sSrc = sSrc & "inputDO=""" & sSVCInVOid & """ "
sSrc = sSrc & "outputDO=""" & sSVCOutVOid & """/>"

sSrc = sSrc & vbCrLf & "    </ns13:opCode>"

sSrc = sSrc & vbCrLf & "    <ns13:member>"
sSrc = sSrc & vbCrLf & "        <ns13:variable kind=""DTO"" name=""input"" arrayKind=""none"" size=""0"" scope=""member"" initObject="""" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"" inject=""NONE"">"
sSrc = sSrc & vbCrLf & "            <ns9:do physicalName=""DummyVo"" resourcePath=""hf/cor/_templet/vo"" inOutType=""INPUT""/>"
sSrc = sSrc & vbCrLf & "        </ns13:variable>"
sSrc = sSrc & vbCrLf & "        <ns13:variable kind=""DTO"" name=""output"" arrayKind=""none"" size=""0"" scope=""member"" initObject="""" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"" inject=""NONE"">"
sSrc = sSrc & vbCrLf & "            <ns9:do physicalName=""DummyVo"" resourcePath=""hf/cor/_templet/vo"" inOutType=""OUTPUT""/>"
sSrc = sSrc & vbCrLf & "        </ns13:variable>"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "        <ns13:variable kind=""OBJECT"" name=""" & sSVC�޼ҵ�id & """ scope=""member"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"" inject=""NONE"">"
sSrc = sSrc & vbCrLf & "            <ns9:userDefined>" & sSVCOutVOid & "</ns9:userDefined>"
sSrc = sSrc & vbCrLf & "        </ns13:variable>"
sSrc = sSrc & vbCrLf & "    </ns13:member>"
sSrc = sSrc & vbCrLf & "    <ns13:wholeImport>"
sSrc = sSrc & vbCrLf & "        <ns13:importType>com.tmax.proobject.core.ServiceObject</ns13:importType>"
sSrc = sSrc & vbCrLf & "        <ns13:importType>com.tmax.proobject.logger.ProObjectLogger</ns13:importType>"
sSrc = sSrc & vbCrLf & "        <ns13:importType>com.tmax.proobject.logger.application.ServiceLogger</ns13:importType>"
sSrc = sSrc & vbCrLf & "    </ns13:wholeImport>"
sSrc = sSrc & vbCrLf & "    <ns13:serviceMethod methodName=""service"" displayName=""service"" modifier=""public"" comments="""" overrided=""true"">"
sSrc = sSrc & vbCrLf & "        <ns9:flowModule>"
sSrc = sSrc & vbCrLf & "            <ns9:entryModule>"
sSrc = sSrc & vbCrLf & "                <ns9:name>service</ns9:name>"
sSrc = sSrc & vbCrLf & "                <ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns9:id>0</ns9:id>"
sSrc = sSrc & vbCrLf & "                </ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "            </ns9:entryModule>"
sSrc = sSrc & vbCrLf & "            <ns9:innerModules/>"
sSrc = sSrc & vbCrLf & "            <ns9:xorModules/>"
sSrc = sSrc & vbCrLf & "            <ns9:virtualModules>"
sSrc = sSrc & vbCrLf & "                <ns9:virtualModule nodeId=""0"">"
sSrc = sSrc & vbCrLf & "                    <ns9:name>return</ns9:name>"
sSrc = sSrc & vbCrLf & "                    <ns9:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns9:code>      &#xD;"
sSrc = sSrc & vbCrLf & "         return null;</ns9:code>"
sSrc = sSrc & vbCrLf & "                </ns9:virtualModule>"
sSrc = sSrc & vbCrLf & "            </ns9:virtualModules>"
sSrc = sSrc & vbCrLf & "            <ns9:loopModules/>"
sSrc = sSrc & vbCrLf & "            <ns9:bizMethodCalls/>"
sSrc = sSrc & vbCrLf & "            <ns9:serviceCalls/>"
sSrc = sSrc & vbCrLf & "            <ns9:assignModules/>"
sSrc = sSrc & vbCrLf & "            <ns9:deferredModules/>"
sSrc = sSrc & vbCrLf & "            <ns9:dataObjectCalls/>"
sSrc = sSrc & vbCrLf & "            <ns9:replyCalls/>"
sSrc = sSrc & vbCrLf & "        </ns9:flowModule>"
sSrc = sSrc & vbCrLf & "        <ns13:arguments>"
sSrc = sSrc & vbCrLf & "            <ns9:argument>"
sSrc = sSrc & vbCrLf & "                <ns9:variable kind=""DTO"" name=""arg0"">"
sSrc = sSrc & vbCrLf & "                    <ns9:do physicalName=""DummyVo"" resourcePath=""hf/cor/_templet/vo""/>"
sSrc = sSrc & vbCrLf & "                </ns9:variable>"
sSrc = sSrc & vbCrLf & "                <ns9:sequence>0</ns9:sequence>"
sSrc = sSrc & vbCrLf & "            </ns9:argument>"
sSrc = sSrc & vbCrLf & "        </ns13:arguments>"
sSrc = sSrc & vbCrLf & "        <ns13:returnType kind=""DTO"">"
sSrc = sSrc & vbCrLf & "            <ns9:do physicalName=""DummyVo"" resourcePath=""hf/cor/_templet/vo""/>"
sSrc = sSrc & vbCrLf & "        </ns13:returnType>"
sSrc = sSrc & vbCrLf & "        <ns13:classInfo classPackageName=""" & sPackageSlashName & """ className=""" & sClassPname & """/>"
sSrc = sSrc & vbCrLf & "        <ns13:exceptions>"
sSrc = sSrc & vbCrLf & "            <ns13:exception>java.lang.Throwable</ns13:exception>"
sSrc = sSrc & vbCrLf & "        </ns13:exceptions>"
sSrc = sSrc & vbCrLf & "        <ns13:annotations>"
sSrc = sSrc & vbCrLf & "            <ns13:annotation>Override</ns13:annotation>"
sSrc = sSrc & vbCrLf & "        </ns13:annotations>"
sSrc = sSrc & vbCrLf & "    </ns13:serviceMethod>"

'biz method
iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "    <ns13:serviceMethod methodName=""" & sSVC�޼ҵ�id & """ displayName=""" & sSVCInVOid & """ modifier=""public"" comments="""">"
sSrc = sSrc & vbCrLf & "        <ns9:flowModule>"
sSrc = sSrc & vbCrLf & "            <ns9:entryModule>"
sSrc = sSrc & vbCrLf & "                <ns9:name>" & sSVCInVOid & "</ns9:name>"
sSrc = sSrc & vbCrLf & "                <ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns9:id>0</ns9:id>"
sSrc = sSrc & vbCrLf & "                    <ns9:id>2</ns9:id>"
sSrc = sSrc & vbCrLf & "                    <ns9:id>4</ns9:id>"
sSrc = sSrc & vbCrLf & "                </ns9:nextEntries>"

'biz vo
iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "                <ns9:variable kind=""DTO"" name=""outBvo"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <ns9:do physicalName=""" & sSVC�޼ҵ�id & """ logicalName=""" & sSVCInVOid & """ resourcePath=""" & sSVCOutVOid & """/>"
sSrc = sSrc & vbCrLf & "                </ns9:variable>"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "                <ns9:variable kind=""DTO"" name=""inBvo"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <ns9:do physicalName=""" & sSVC�޼ҵ�id & """ logicalName=""" & sSVCInVOid & """ resourcePath=""" & sSVCOutVOid & """/>"
sSrc = sSrc & vbCrLf & "                </ns9:variable>"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "                <ns9:variable kind=""DTO"" name=""outSvo"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <ns9:do physicalName=""" & sSVC�޼ҵ�id & """ logicalName=""" & sSVCInVOid & """ resourcePath=""" & sSVCOutVOid & """/>"
sSrc = sSrc & vbCrLf & "                </ns9:variable>"
sSrc = sSrc & vbCrLf & "            </ns9:entryModule>"

sSrc = sSrc & vbCrLf & "            <ns9:innerModules>"
sSrc = sSrc & vbCrLf & "                <ns9:innerModule kind=""block"" nodeId=""0"">"
sSrc = sSrc & vbCrLf & "                    <ns9:name>��������ó��</ns9:name>"
sSrc = sSrc & vbCrLf & "                    <ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns9:id>1</ns9:id>"
sSrc = sSrc & vbCrLf & "                    </ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns9:condition/>"
sSrc = sSrc & vbCrLf & "                </ns9:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns9:innerModule kind=""block"" nodeId=""2"">"
sSrc = sSrc & vbCrLf & "                    <ns9:name>���� �� ó��</ns9:name>"
sSrc = sSrc & vbCrLf & "                    <ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns9:id>6</ns9:id>"
sSrc = sSrc & vbCrLf & "                    </ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns9:condition/>"
sSrc = sSrc & vbCrLf & "                </ns9:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns9:innerModule kind=""block"" nodeId=""4"">"
sSrc = sSrc & vbCrLf & "                    <ns9:name>��������ó��</ns9:name>"
sSrc = sSrc & vbCrLf & "                    <ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns9:id>5</ns9:id>"
sSrc = sSrc & vbCrLf & "                    </ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns9:condition/>"
sSrc = sSrc & vbCrLf & "                </ns9:innerModule>"
sSrc = sSrc & vbCrLf & "            </ns9:innerModules>"
sSrc = sSrc & vbCrLf & "            <ns9:xorModules/>"
sSrc = sSrc & vbCrLf & "            <ns9:virtualModules>"
sSrc = sSrc & vbCrLf & "                <ns9:virtualModule nodeId=""1"">"
sSrc = sSrc & vbCrLf & "                    <ns9:name>�Է°� Ȯ��</ns9:name>"
sSrc = sSrc & vbCrLf & "                    <ns9:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns9:code>          logger.info(""  �١ڡ١ڡ١ڡ١�inSvo   = \n[""+inSvo.toString()+""] �١ڡ١ڡ١ڡ١�"" );&#xD;"
sSrc = sSrc & vbCrLf & "            inBvo.setInqComnVo(inSvo.getInqComnVo());</ns9:code>"
sSrc = sSrc & vbCrLf & "                </ns9:virtualModule>"
sSrc = sSrc & vbCrLf & "                <ns9:virtualModule nodeId=""5"">"
sSrc = sSrc & vbCrLf & "                    <ns9:name>return</ns9:name>"
sSrc = sSrc & vbCrLf & "                    <ns9:nextEntries/>"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "                    <ns9:code>       " & sSVC�޼ҵ�id & ";&#xD;"
sSrc = sSrc & vbCrLf & "         " & sSVCInVOid & "; &#xD;"
sSrc = sSrc & vbCrLf & "         logger.info(""  �١ڡ١ڡ١ڡ١�outSvo   = \n[""+outSvo.toString()+""] �١ڡ١ڡ١ڡ١�"" );&#xD;"
sSrc = sSrc & vbCrLf & "         return outSvo;</ns9:code>"
sSrc = sSrc & vbCrLf & "                </ns9:virtualModule>"
sSrc = sSrc & vbCrLf & "            </ns9:virtualModules>"
sSrc = sSrc & vbCrLf & "            <ns9:loopModules/>"
sSrc = sSrc & vbCrLf & "            <ns9:bizMethodCalls>"
sSrc = sSrc & vbCrLf & "                <ns9:bizMethodCall nodeId=""6"">"
sSrc = sSrc & vbCrLf & "                    <ns9:nextEntries/>"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "                    <ns9:bizInstanceInfo scope=""member"" variableName=""" & sSVC�޼ҵ�id & """ lazy=""false"">"
sSrc = sSrc & vbCrLf & "                        <ns9:classInfo classPackageName=""" & sSVCInVOid & """ className=""" & sSVC�޼ҵ�id & """/>"
sSrc = sSrc & vbCrLf & "                    </ns9:bizInstanceInfo>"


iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "                    <ns9:method methodName=""" & sSVC�޼ҵ�id & """ displayName=""" & sSVCInVOid & """>"
sSrc = sSrc & vbCrLf & "                        <ns12:arguments>"
sSrc = sSrc & vbCrLf & "                            <ns9:argument>"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "<ns9:variable kind=""DTO"" name=""inBvo"">"
sSrc = sSrc & vbCrLf & "    <ns9:do physicalName=""" & sSVC�޼ҵ�id & """ resourcePath=""" & sSVCOutVOid & """/>"
sSrc = sSrc & vbCrLf & "</ns9:variable>"
sSrc = sSrc & vbCrLf & "<ns9:sequence>0</ns9:sequence>"


sSrc = sSrc & vbCrLf & "                            </ns9:argument>"
sSrc = sSrc & vbCrLf & "                        </ns12:arguments>"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "                        <ns12:returnType kind=""DTO"">"
sSrc = sSrc & vbCrLf & "                            <ns9:do physicalName=""" & sSVC�޼ҵ�id & """ resourcePath=""" & sSVCOutVOid & """/>"
sSrc = sSrc & vbCrLf & "                        </ns12:returnType>"


sSrc = sSrc & vbCrLf & "                    </ns9:method>"
sSrc = sSrc & vbCrLf & "                    <ns9:in>"
sSrc = sSrc & vbCrLf & "                        <ns9:transformation/>"
sSrc = sSrc & vbCrLf & "                        <ns9:byPass>"
sSrc = sSrc & vbCrLf & "                            <ns9:memberVariable name=""inBvo"" order=""0""/>"
sSrc = sSrc & vbCrLf & "                        </ns9:byPass>"
sSrc = sSrc & vbCrLf & "                    </ns9:in>"
sSrc = sSrc & vbCrLf & "                    <ns9:out>"
sSrc = sSrc & vbCrLf & "                        <ns9:transformation/>"
sSrc = sSrc & vbCrLf & "                        <ns9:byPass>"
sSrc = sSrc & vbCrLf & "                            <ns9:memberVariable name=""outBvo"" order=""0""/>"
sSrc = sSrc & vbCrLf & "                        </ns9:byPass>"
sSrc = sSrc & vbCrLf & "                    </ns9:out>"
sSrc = sSrc & vbCrLf & "                    <ns9:preAssign/>"
sSrc = sSrc & vbCrLf & "                    <ns9:postAssign/>"
sSrc = sSrc & vbCrLf & "                </ns9:bizMethodCall>"
sSrc = sSrc & vbCrLf & "            </ns9:bizMethodCalls>"
sSrc = sSrc & vbCrLf & "            <ns9:serviceCalls/>"
sSrc = sSrc & vbCrLf & "            <ns9:assignModules/>"
sSrc = sSrc & vbCrLf & "            <ns9:deferredModules/>"
sSrc = sSrc & vbCrLf & "            <ns9:dataObjectCalls/>"
sSrc = sSrc & vbCrLf & "            <ns9:replyCalls/>"
sSrc = sSrc & vbCrLf & "        </ns9:flowModule>"
sSrc = sSrc & vbCrLf & "        <ns13:arguments>"
sSrc = sSrc & vbCrLf & "            <ns9:argument>"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "                <ns9:variable kind=""DTO"" name=""inSvo"">"
sSrc = sSrc & vbCrLf & "                    <ns9:do physicalName=""" & sSVC�޼ҵ�id & """ resourcePath=""" & sSVCOutVOid & """/>"
sSrc = sSrc & vbCrLf & "                </ns9:variable>"


sSrc = sSrc & vbCrLf & "                <ns9:sequence>0</ns9:sequence>"
sSrc = sSrc & vbCrLf & "            </ns9:argument>"
sSrc = sSrc & vbCrLf & "        </ns13:arguments>"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "        <ns13:returnType kind=""DTO"" name=""arg0"">"
sSrc = sSrc & vbCrLf & "            <ns9:do physicalName=""" & sSVC�޼ҵ�id & """ resourcePath=""" & sSVCOutVOid & """/>"
sSrc = sSrc & vbCrLf & "        </ns13:returnType>"


sSrc = sSrc & vbCrLf & "        <ns13:classInfo classPackageName=""" & sPackageSlashName & """ className=""" & sClassPname & """/>"
sSrc = sSrc & vbCrLf & "        <ns13:annotations/>"
sSrc = sSrc & vbCrLf & "    </ns13:serviceMethod>"



'biz method
iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "    <ns13:serviceMethod methodName=""" & sSVC�޼ҵ�id & """ displayName=""" & sSVCInVOid & """ modifier=""public"" comments="""">"


sSrc = sSrc & vbCrLf & "        <ns9:flowModule>"
sSrc = sSrc & vbCrLf & "            <ns9:entryModule>"
sSrc = sSrc & vbCrLf & "                <ns9:name>" & sSVCInVOid & "</ns9:name>"
sSrc = sSrc & vbCrLf & "                <ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns9:id>0</ns9:id>"
sSrc = sSrc & vbCrLf & "                    <ns9:id>2</ns9:id>"
sSrc = sSrc & vbCrLf & "                    <ns9:id>4</ns9:id>"
sSrc = sSrc & vbCrLf & "                </ns9:nextEntries>"


'biz vo
iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "                <ns9:variable kind=""DTO"" name=""outBvo"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <ns9:do physicalName=""" & sSVC�޼ҵ�id & """ logicalName=""" & sSVCInVOid & """ resourcePath=""" & sSVCOutVOid & """/>"
sSrc = sSrc & vbCrLf & "                </ns9:variable>"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "                <ns9:variable kind=""DTO"" name=""inBvo"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <ns9:do physicalName=""" & sSVC�޼ҵ�id & """ logicalName=""" & sSVCInVOid & """ resourcePath=""" & sSVCOutVOid & """/>"
sSrc = sSrc & vbCrLf & "                </ns9:variable>"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "                <ns9:variable kind=""DTO"" name=""outSvo"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <ns9:do physicalName=""" & sSVC�޼ҵ�id & """ logicalName=""" & sSVCInVOid & """ resourcePath=""" & sSVCOutVOid & """/>"
sSrc = sSrc & vbCrLf & "                </ns9:variable>"
sSrc = sSrc & vbCrLf & "            </ns9:entryModule>"





sSrc = sSrc & vbCrLf & "            <ns9:innerModules>"
sSrc = sSrc & vbCrLf & "                <ns9:innerModule kind=""block"" nodeId=""0"">"
sSrc = sSrc & vbCrLf & "                    <ns9:name>��������ó��</ns9:name>"
sSrc = sSrc & vbCrLf & "                    <ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns9:id>1</ns9:id>"
sSrc = sSrc & vbCrLf & "                    </ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns9:condition/>"
sSrc = sSrc & vbCrLf & "                </ns9:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns9:innerModule kind=""block"" nodeId=""2"">"
sSrc = sSrc & vbCrLf & "                    <ns9:name>���� �� ó��</ns9:name>"
sSrc = sSrc & vbCrLf & "                    <ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns9:id>7</ns9:id>"
sSrc = sSrc & vbCrLf & "                    </ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns9:condition/>"
sSrc = sSrc & vbCrLf & "                </ns9:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns9:innerModule kind=""block"" nodeId=""4"">"
sSrc = sSrc & vbCrLf & "                    <ns9:name>��������ó��</ns9:name>"
sSrc = sSrc & vbCrLf & "                    <ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns9:id>5</ns9:id>"
sSrc = sSrc & vbCrLf & "                    </ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns9:condition/>"
sSrc = sSrc & vbCrLf & "                </ns9:innerModule>"
sSrc = sSrc & vbCrLf & "            </ns9:innerModules>"
sSrc = sSrc & vbCrLf & "            <ns9:xorModules/>"
sSrc = sSrc & vbCrLf & "            <ns9:virtualModules>"
sSrc = sSrc & vbCrLf & "                <ns9:virtualModule nodeId=""1"">"
sSrc = sSrc & vbCrLf & "                    <ns9:name>�Է°� Ȯ��</ns9:name>"
sSrc = sSrc & vbCrLf & "                    <ns9:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns9:code>          logger.info(""  �١ڡ١ڡ١ڡ١�inSvo   = \n[""+inSvo.toString()+""] �١ڡ١ڡ١ڡ١�"" );&#xD;"
sSrc = sSrc & vbCrLf & "            inBvo.setRecvHdrVo(inSvo.getRecvHdrVo());&#xD;"
sSrc = sSrc & vbCrLf & "            inBvo.setRecvDataVo(inSvo.getRecvDataVo());</ns9:code>"
sSrc = sSrc & vbCrLf & "                </ns9:virtualModule>"
sSrc = sSrc & vbCrLf & "                <ns9:virtualModule nodeId=""5"">"
sSrc = sSrc & vbCrLf & "                    <ns9:name>return</ns9:name>"
sSrc = sSrc & vbCrLf & "                    <ns9:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns9:code>       //outSvo.setListDvo(outBvo.getListDvo()); &#xD;"
sSrc = sSrc & vbCrLf & "         &#xD;"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "          " & sSVC�޼ҵ�id & ";&#xD;"
sSrc = sSrc & vbCrLf & "         logger.info(""  �١ڡ١ڡ١ڡ١�outSvo   = \n[""+outSvo.toString()+""] �١ڡ١ڡ١ڡ١�"" );&#xD;"
sSrc = sSrc & vbCrLf & "         return outSvo;</ns9:code>"
sSrc = sSrc & vbCrLf & "                </ns9:virtualModule>"
sSrc = sSrc & vbCrLf & "            </ns9:virtualModules>"
sSrc = sSrc & vbCrLf & "            <ns9:loopModules/>"
sSrc = sSrc & vbCrLf & "            <ns9:bizMethodCalls>"
sSrc = sSrc & vbCrLf & "                <ns9:bizMethodCall nodeId=""7"">"
sSrc = sSrc & vbCrLf & "                    <ns9:nextEntries/>"


iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "                    <ns9:bizInstanceInfo scope=""member"" variableName=""" & sSVC�޼ҵ�id & """ lazy=""false"">"
sSrc = sSrc & vbCrLf & "                        <ns9:classInfo classPackageName=""" & sSVCInVOid & """ className=""" & sSVC�޼ҵ�id & """/>"
sSrc = sSrc & vbCrLf & "                    </ns9:bizInstanceInfo>"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "                    <ns9:method methodName=""" & sSVC�޼ҵ�id & """ displayName=""" & sSVCInVOid & """>"
sSrc = sSrc & vbCrLf & "                        <ns12:arguments>"
sSrc = sSrc & vbCrLf & "                            <ns9:argument>"


iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "<ns9:variable kind=""DTO"" name=""inBvo"">"
sSrc = sSrc & vbCrLf & "    <ns9:do physicalName=""" & sSVC�޼ҵ�id & """ resourcePath=""" & sSVCOutVOid & """/>"
sSrc = sSrc & vbCrLf & "</ns9:variable>"
sSrc = sSrc & vbCrLf & "<ns9:sequence>0</ns9:sequence>"


sSrc = sSrc & vbCrLf & "                            </ns9:argument>"
sSrc = sSrc & vbCrLf & "                        </ns12:arguments>"


iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "                        <ns12:returnType kind=""DTO"">"
sSrc = sSrc & vbCrLf & "                            <ns9:do physicalName=""" & sSVC�޼ҵ�id & """ resourcePath=""" & sSVCOutVOid & """/>"
sSrc = sSrc & vbCrLf & "                        </ns12:returnType>"



sSrc = sSrc & vbCrLf & "                    </ns9:method>"
sSrc = sSrc & vbCrLf & "                    <ns9:in>"
sSrc = sSrc & vbCrLf & "                        <ns9:transformation/>"
sSrc = sSrc & vbCrLf & "                        <ns9:byPass>"
sSrc = sSrc & vbCrLf & "                            <ns9:memberVariable name=""inBvo"" order=""0""/>"
sSrc = sSrc & vbCrLf & "                        </ns9:byPass>"
sSrc = sSrc & vbCrLf & "                    </ns9:in>"
sSrc = sSrc & vbCrLf & "                    <ns9:out>"
sSrc = sSrc & vbCrLf & "                        <ns9:transformation/>"
sSrc = sSrc & vbCrLf & "                        <ns9:byPass>"
sSrc = sSrc & vbCrLf & "                            <ns9:memberVariable name=""outBvo"" order=""0""/>"
sSrc = sSrc & vbCrLf & "                        </ns9:byPass>"
sSrc = sSrc & vbCrLf & "                    </ns9:out>"
sSrc = sSrc & vbCrLf & "                    <ns9:preAssign/>"
sSrc = sSrc & vbCrLf & "                    <ns9:postAssign/>"
sSrc = sSrc & vbCrLf & "                </ns9:bizMethodCall>"
sSrc = sSrc & vbCrLf & "            </ns9:bizMethodCalls>"
sSrc = sSrc & vbCrLf & "            <ns9:serviceCalls/>"
sSrc = sSrc & vbCrLf & "            <ns9:assignModules/>"
sSrc = sSrc & vbCrLf & "            <ns9:deferredModules/>"
sSrc = sSrc & vbCrLf & "            <ns9:dataObjectCalls/>"
sSrc = sSrc & vbCrLf & "            <ns9:replyCalls/>"
sSrc = sSrc & vbCrLf & "        </ns9:flowModule>"
sSrc = sSrc & vbCrLf & "        <ns13:arguments>"
sSrc = sSrc & vbCrLf & "            <ns9:argument>"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "                <ns9:variable kind=""DTO"" name=""inSvo"">"
sSrc = sSrc & vbCrLf & "                    <ns9:do physicalName=""" & sSVC�޼ҵ�id & """ resourcePath=""" & sSVCOutVOid & """/>"
sSrc = sSrc & vbCrLf & "                </ns9:variable>"


sSrc = sSrc & vbCrLf & "                <ns9:sequence>0</ns9:sequence>"
sSrc = sSrc & vbCrLf & "            </ns9:argument>"
sSrc = sSrc & vbCrLf & "        </ns13:arguments>"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "        <ns13:returnType kind=""DTO"" name=""arg0"">"
sSrc = sSrc & vbCrLf & "            <ns9:do physicalName=""" & sSVC�޼ҵ�id & """ resourcePath=""" & sSVCOutVOid & """/>"
sSrc = sSrc & vbCrLf & "        </ns13:returnType>"
sSrc = sSrc & vbCrLf & "        <ns13:classInfo classPackageName=""" & sPackageSlashName & """ className=""" & sClassPname & """/>"
sSrc = sSrc & vbCrLf & "        <ns13:annotations/>"
sSrc = sSrc & vbCrLf & "    </ns13:serviceMethod>"
sSrc = sSrc & vbCrLf & "    <ns13:interfaces/>"
sSrc = sSrc & vbCrLf & "</ns13:serviceObject>"

                                                                                                                                                                                                                              
End With
sMakeSrcStringSVC = sSrc
End Function


'����� VO ����
Private Sub subMake�����VO()


On Error GoTo Err_Handler


    Dim sSrc As String
    
    
    'Dim sPname, sLname, sPath As String
    
    
    Dim iCnt As Integer
    
    iCnt = 1


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
                .SaveToFile "c:\src" & "\" & sPname & ".dto", adSaveCreateOverWrite
                
            End With
            
            objStreamUTF8.Close
            objStreamUTF8NoBOM.Close
            
            Set objStreamUTF8 = Nothing
            Set objStreamUTF8NoBOM = Nothing


    'gfMsgInf ("�����Ǿ����ϴ�.")
    Exit Sub
Err_Handler:
    gfMsgWrn (Err.Description)
End Sub






'����� DAO ����
Private Sub subMake�����DAO()


On Error GoTo Err_Handler


    Dim sSrc As String
    
    
    'Dim sPname, sLname, sPath As String
    
    
    Dim iCnt As Integer
    
    iCnt = 1


        sSrc = ""
        sSrc = sSrc & "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf
        sSrc = sSrc & "<ns12:queryObject xmlns:ns2=""http://www.tmax.co.kr/proobject/resource"" xmlns:ns3=""http://www.tmax.co.kr/proobject/testcase"" "
        sSrc = sSrc & "xmlns:ns4=""http://www.tmax.co.kr/proobject/message"" xmlns:ns5=""http://www.tmax.co.kr/proobject/dto"" xmlns:ns6=""http://www.tmaxsoft.co.kr/proobject/testsuite"" "
        sSrc = sSrc & "xmlns:ns7=""http://www.tmax.co.kr/proobject/dto/validator"" xmlns:ns8=""http://www.tmax.co.kr/proobject/serviceobject-automatic"" "
        sSrc = sSrc & "xmlns:ns9=""http://www.tmax.co.kr/proobject/sourcecode"" xmlns:ns10=""http://www.tmax.co.kr/proobject/flow"" xmlns:ns11=""http://www.tmax.co.kr/proobject/dataobjectfactory"" "
        sSrc = sSrc & "xmlns:ns12=""http://www.tmax.co.kr/proobject/queryobject"" xmlns:ns13=""http://www.tmax.co.kr/proobject/bizobject"" xmlns:ns14=""http://www.tmax.co.kr/proobject/serviceobject"" "
        sSrc = sSrc & "xmlns:ns15=""http://www.tmax.co.kr/proobject/dataobject"" xmlns:ns16=""http://www.tmax.co.kr/proobject/taskobject"" xmlns:ns17=""http://www.tmax.co.kr/proobject/jobobject"" "
        sSrc = sSrc & "xmlns:ns18=""http://www.tmax.co.kr/proobject/servicegroup"" "
        sSrc = sSrc & "xmlns:ns19=""http://www.tmax.co.kr/proobject/aspectj"" xmlns:ns20=""http://www.tmax.co.kr/proobject/mapping"" xmlns:ns21=""http://www.tmax.co.kr/proobject/application/runtime"" "
        sSrc = sSrc & "xmlns:ns22=""http://www.tmaxsoft.co.kr/proobject/testOperation"" xmlns:ns23=""http://www.tmax.co.kr/proobject/restriction_codes"" "
        sSrc = sSrc & "xmlns:ns24=""http://www.example.org/externalObjectConfig"" xmlns:ns25=""http://www.tmax.co.kr/proobject/application"" xmlns:ns26=""http://www.tmax.co.kr/proobject/contents"" "
        sSrc = sSrc & "xmlns:ns27=""http://www.example.org/testData"" xmlns:ns28=""http://www.tmaxsoft.co.kr/proobject/testresult"" xmlns:ns29=""http://www.tmax.co.kr/proobject/siteConfig"" sSrc = sSrc & "
        sSrc = sSrc & "xmlns:ns30=""http://www.tmax.co.kr/proobject/serverConfig"" xmlns:ns31=""http://www.tmax.co.kr/proobject/probuilder_config"" xmlns:ns32=""http://www.tmax.co.kr/proobject/property"" "
        sSrc = sSrc & "xmlns:ns33=""http://www.tmax.co.kr/proobject/serviceGroup"" queryType=""DBIO"" "
        sSrc = sSrc & "resourceId=""" & sPname & """ "
        sSrc = sSrc & "logicalName=""" & sLname & """ "
        sSrc = sSrc & "physicalName=""" & sPname & """ "

        sSrc = sSrc & "resourceType=""QUERY"" "
        sSrc = sSrc & "resourcePackage=""hf.cor.hg.bg.dao.sdao." & fService3() & """ "
        
        sSrc = sSrc & "creator=""" & gUserId & """ "
        sSrc = sSrc & "owner=""" & gUserId & """ "
        sSrc = sSrc & "revision=""1"" "
        sSrc = sSrc & "resourcePath=""" & sPath & """ "
        sSrc = sSrc & "modifier=""" & gUserId & """ "
        sSrc = sSrc & "created=""2020-04-06 19:00:54"" "
        sSrc = sSrc & "lastUpdated=""2020-04-06 19:00:54"">"
        sSrc = sSrc & vbCrLf & "    <ns12:dbioAttribute>"
        'sSrc = sSrc & "        <ns12:statements sqlType=""SELECT"" alias=""sel001"" resultType=""HgbgBzepBrfHsCnstCrs001Dvo"" resultTypeResourcePath=""hf.cor.hg.bg.vo.dvo.sdvo.s523"">"
        'sSrc = sSrc & "            <ns12:statement>/* hf.cor.hg.bg.dao.sdao.s523.HgbgBzepBrfHsCnstCrs001Dao.sel001 ��ü����-�ð������ȣ ��ȸ */"
        'sSrc = sSrc & "SELECT T1.JINT_BIZ_DVCD             AS JINT_BIZ_DVCD /* ��������ڱ����ڵ� */"
        'sSrc = sSrc & "     , T2.CNCO_CSNO                 AS CNCO_CSNO     /* �ð������ȣ */"
        'sSrc = sSrc & "  FROM TB_HGA701D_BIZGRNT T1        /* ���ýſ뺸��_����ں����� */"
        'sSrc = sSrc & "     , TB_HGA208M_BIZCNCO T2        /* ���ýſ뺸��_����ڽð���⺻ */"
        'sSrc = sSrc & " WHERE T1.GRNT_NO = T2.GRNT_NO"
        'sSrc = sSrc & "   AND T1.GRNT_NO = :grntNo"
        'sSrc = sSrc & "   AND T2.CNCO_SEQ = 1              /* �ð������(AS-IS���� �÷�.) */</ns12:statement>"
        'sSrc = sSrc & "            <ns12:parameter physicalName=""grntNo"" javaType=""String"" parameter=""true""/>"
        'sSrc = sSrc & "            <ns12:table name=""TB_HGA701D_BIZGRNT"" alias=""TB_HGA701D_BIZGRNT""/>"
        'sSrc = sSrc & "        </ns12:statements>"
        'sSrc = sSrc & "        <ns12:statements sqlType=""SELECT"" alias=""sel002"" resultType=""HgbgBzepBrfHsCnstCrs001Dvo"" resultTypeResourcePath=""hf.cor.hg.bg.vo.dvo.sdvo.s523"">"
        'sSrc = sSrc & "            <ns12:statement>/* hf.cor.hg.bg.dao.sdao.s523.HgbgBzepBrfHsCnstCrs001Dao.sel002 ��ü����-������� ��ȸ */"
        'sSrc = sSrc & "SELECT T1.RSCH_BASIS_DY                 AS RSCH_BASIS_DY                /* ����������� */"
        'sSrc = sSrc & "     , T1.LSTR_ACST_BASIS_TRED_DY       AS LSTR_ACST_BASIS_TRED_DY      /* ��������ر⸻���� */"
        'sSrc = sSrc & "     , T1.CRTM_ACST_BASIS_TRED_DY       AS CRTM_ACST_BASIS_TRED_DY      /* �������ر⸻���� */"
        'sSrc = sSrc & "     , T2.ETNO                          AS ETNO                                              /* ���ü��ȣ */"
        'sSrc = sSrc & "  FROM TB_HGJ003M_CUSTRVEW T1               /* ���ýſ뺸��_���ɻ�⺻ */"
        'sSrc = sSrc & "     , VW_CU_CUSTMST T2"
        'sSrc = sSrc & " WHERE T1.CSNO = T2.CSNO"
        'sSrc = sSrc & "   AND T1.CSNO = :csno"
        'sSrc = sSrc & "   AND T1.RVEW_SEQ = :rvewSeq</ns12:statement>"
        'sSrc = sSrc & "            <ns12:parameter physicalName=""csno"" javaType=""String"" parameter=""true""/>"
        'sSrc = sSrc & "            <ns12:parameter physicalName=""rvewSeq"" javaType=""String"" parameter=""true""/>"
        'sSrc = sSrc & "            <ns12:table name=""TB_HGJ003M_CUSTRVEW"" alias=""TB_HGJ003M_CUSTRVEW""/>"
        'sSrc = sSrc & "        </ns12:statements>"
        sSrc = sSrc & vbCrLf & "    </ns12:dbioAttribute>"
        sSrc = sSrc & vbCrLf & "</ns12:queryObject>"
                
        
        
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
                .SaveToFile "c:\src" & "\" & sPname & ".qo", adSaveCreateOverWrite
                
            End With
            
            
            
            objStreamUTF8.Close
            objStreamUTF8NoBOM.Close
            
            Set objStreamUTF8 = Nothing
            Set objStreamUTF8NoBOM = Nothing


    'gfMsgInf ("�����Ǿ����ϴ�.")
    Exit Sub
Err_Handler:
    gfMsgWrn (Err.Description)
End Sub





