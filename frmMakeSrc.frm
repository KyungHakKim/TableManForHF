VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMakeSrc 
   Caption         =   "개인보증 소스 일괄 생성"
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
      TabCaption(3)   =   "사업자"
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
         Caption         =   "작업"
         Height          =   555
         Left            =   600
         TabIndex        =   12
         Top             =   480
         Width           =   3315
      End
      Begin VB.CommandButton Command1 
         Caption         =   "사업자 biz 소스 생성"
         Height          =   555
         Left            =   -67920
         TabIndex        =   10
         Top             =   11160
         Width           =   3315
      End
      Begin VB.CommandButton cmdBGExcel 
         Caption         =   "사업자 Excel 읽기"
         Height          =   555
         Left            =   -74520
         TabIndex        =   9
         Top             =   11160
         Width           =   3315
      End
      Begin VB.CommandButton cmbBGMake 
         Caption         =   "사업자 SVC 소스 생성"
         Height          =   555
         Left            =   -71280
         TabIndex        =   8
         Top             =   11160
         Width           =   3315
      End
      Begin VB.CommandButton cmdMakeSrcSVC 
         Caption         =   "SVC 소스 생성"
         Height          =   795
         Left            =   -71220
         TabIndex        =   6
         Top             =   11040
         Width           =   3315
      End
      Begin VB.CommandButton cmdExcelSVC 
         Caption         =   "SVC Excel 읽기"
         Height          =   795
         Left            =   -74640
         TabIndex        =   5
         Top             =   11040
         Width           =   3315
      End
      Begin VB.CommandButton cmdMakeSrc 
         Caption         =   "소스 생성"
         Height          =   795
         Left            =   -71340
         TabIndex        =   3
         Top             =   10920
         Width           =   3315
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "Excel 읽기"
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

Dim s메소드id, s메소드명, sINBVO, sOUTBVO As String
Dim sVO변수명, sVOid, sVO명, sVO경로 As String
Dim sSVC메소드id, sSVCInVOid, sSVCOutVOid As String

'사업자 관련 변수 선언
Dim s사업자서비스ID, s사업자BizID, s사업자SVC명, s사업자Biz축소 As String


Dim sPname, sLname, sPath As String

'" & s사업자서비스ID & "
'" & s사업자SVC명 & "
'" & s사업자Biz축소 & "
'" & fService3() & "

Private Function fService3() As String
    
    fService3 = LCase(Left(s사업자서비스ID, 1)) & Mid(s사업자서비스ID, 2, 3)


End Function

'사업자 소스 생성 버튼 클릭
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
    .Col = 1: s사업자서비스ID = .Value
    .Col = 2: s사업자Biz축소 = .Value
    s사업자BizID = .Value & "Biz"
    .Col = 3: s사업자SVC명 = .Value
    

    
    'file name 저장
    sFilename = s사업자서비스ID & ".so"
    
    
    '출력할 문자열 산출
    sPrintString = sMake사업자SVC
    
    
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
    '끝이 아니면 반복한다.
    If UCase(Trim(.Value)) <> "END" Then
        'iROW_CNT = iROW_CNT + 1
        GoTo LOOP_START
    End If
End With

    gfMsgInf "완료되었습니다."
    Exit Sub
Err_Handler:
    gfMsgWrn (Err.Description)
End Sub




Private Function sMake사업자SVC() As String
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
sSrc = sSrc & "name=""" & s사업자서비스ID & """ "
sSrc = sSrc & "resourceId=""" & s사업자서비스ID & """ "
sSrc = sSrc & "logicalName=""" & s사업자SVC명 & """ "
sSrc = sSrc & "physicalName=""" & s사업자서비스ID & """ "
sSrc = sSrc & "resourceType=""SERVICE_OBJECT"" "
sSrc = sSrc & "resourcePackage=""hf.cor.hg.bg.svc." & fService3() & """ "
sSrc = sSrc & "creator=""ns042"" "  '여기 변경
sSrc = sSrc & "revision=""1"" "
sSrc = sSrc & "description=""@파일명 : " & s사업자서비스ID & "&#xD;&#xA;"
sSrc = sSrc & "@논리명 : " & s사업자SVC명 & " wframe 서비스&#xD;&#xA;"
sSrc = sSrc & "@작성자 : 김근수(ns042)&#xD;&#xA;--------------------------------------------------------------------&#xD;&#xA; "  '여기변경
sSrc = sSrc & "* 수정일자 : "
sSrc = sSrc & "수정자 : 요청자(SR 번호) : "
sSrc = sSrc & "수정된 메소드 : 수정내용&#xD;&#xA;--------------------------------------------------------------------&#xD;&#xA;* "
sSrc = sSrc & "작성일자 : 2020.04.03 : "
sSrc = sSrc & "작성자 : 김근수(ns042),  "   '여기변경
sSrc = sSrc & "작성내용 : "
sSrc = sSrc & "최초작성"" "
sSrc = sSrc & "resourcePath=""hf/cor/hg/bg/svc/" & fService3() & """ "
sSrc = sSrc & "modifier=""ns042"" "         '여기변경
sSrc = sSrc & "created=""2020-04-06 19:04:28"" "
sSrc = sSrc & "lastUpdated=""2020-04-06 19:04:28""> "
sSrc = sSrc & vbCrLf & "    <ns14:inputDo xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:type=""ns15:dataObjectType"" resourceId=""urn:proobject:hf.cor.hg.bg.svc." & fService3() & ":" & s사업자서비스ID & """ physicalName=""DummyVo"" resourceType=""DATA_OBJECT"" serviceGroupName=""hf2_cor_service"" resourcePackage=""hf.cor._templet.vo"" revision=""0"" resourcePath=""hf/cor/_templet/vo""/>"
sSrc = sSrc & vbCrLf & "    <ns14:outputDo xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:type=""ns15:dataObjectType"" resourceId=""urn:proobject:hf.cor.hg.bg.svc." & fService3() & ":" & s사업자서비스ID & """ physicalName=""DummyVo"" resourceType=""DATA_OBJECT"" serviceGroupName=""hf2_cor_service"" resourcePackage=""hf.cor._templet.vo"" revision=""0"" resourcePath=""hf/cor/_templet/vo""/>"
sSrc = sSrc & vbCrLf & "    <ns14:opCodeEnable>true</ns14:opCodeEnable>"
sSrc = sSrc & vbCrLf & "    <ns14:opCode>"

                                                                                         
sSrc = sSrc & vbCrLf & "        <ns14:opCode opCode=""sel" & s사업자Biz축소 & """ svcName=""sel" & s사업자Biz축소 & """ inputDO=""hf.cor.hg.bg.vo.svo." & fService3() & ".HgbgSel" & s사업자Biz축소 & "InSvo"" outputDO=""hf.cor.hg.bg.vo.svo." & fService3() & ".HgbgSel" & s사업자Biz축소 & "OutSvo""/>"

sPname = "HgbgSel" & s사업자Biz축소 & "InSvo"
sLname = s사업자SVC명 & " 조회 InSvo"
sPath = Replace("hf.cor.hg.bg.vo.svo." & fService3(), ".", "/")
Call subMake사업자VO


sPname = "HgbgSel" & s사업자Biz축소 & "OutSvo"
sLname = s사업자SVC명 & " 조회 OutSvo"
sPath = Replace("hf.cor.hg.bg.vo.svo." & fService3(), ".", "/")
Call subMake사업자VO

sSrc = sSrc & vbCrLf & "        <ns14:opCode opCode=""prcs" & s사업자Biz축소 & """ svcName=""prcs" & s사업자Biz축소 & """ inputDO=""hf.cor.hg.bg.vo.svo." & fService3() & ".HgbgPrcs" & s사업자Biz축소 & "InSvo"" outputDO=""hf.cor.hg.bg.vo.svo." & fService3() & ".HgbgPrcs" & s사업자Biz축소 & "OutSvo""/>"

sPname = "HgbgPrcs" & s사업자Biz축소 & "InSvo"
sLname = s사업자SVC명 & " 처리 InSvo"
sPath = Replace("hf.cor.hg.bg.vo.svo." & fService3(), ".", "/")
Call subMake사업자VO



sPname = "HgbgPrcs" & s사업자Biz축소 & "OutSvo"
sLname = s사업자SVC명 & " 처리 OutSvo"
sPath = Replace("hf.cor.hg.bg.vo.svo." & fService3(), ".", "/")
Call subMake사업자VO



sSrc = sSrc & vbCrLf & "    </ns14:opCode>"
sSrc = sSrc & vbCrLf & "    <ns14:member>"
sSrc = sSrc & vbCrLf & "        <ns14:variable kind=""DTO"" name=""input"" arrayKind=""none"" size=""0"" scope=""member"" initObject="""" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"" inject=""NONE"">"
sSrc = sSrc & vbCrLf & "            <ns10:do physicalName=""DummyVo"" resourcePath=""hf/cor/_templet/vo"" inOutType=""INPUT""/>"
sSrc = sSrc & vbCrLf & "        </ns14:variable>"
sSrc = sSrc & vbCrLf & "        <ns14:variable kind=""DTO"" name=""output"" arrayKind=""none"" size=""0"" scope=""member"" initObject="""" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"" inject=""NONE"">"
sSrc = sSrc & vbCrLf & "            <ns10:do physicalName=""DummyVo"" resourcePath=""hf/cor/_templet/vo"" inOutType=""OUTPUT""/>"
sSrc = sSrc & vbCrLf & "        </ns14:variable>"
sSrc = sSrc & vbCrLf & "        <ns14:variable kind=""OBJECT"" name=""bizHgbg" & s사업자Biz축소 & """ scope=""member"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"" inject=""NONE"">"
sSrc = sSrc & vbCrLf & "            <ns10:userDefined>hf.cor.hg.bg.biz.sbiz." & fService3() & ".Hgbg" & s사업자Biz축소 & "Biz</ns10:userDefined>"
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
sSrc = sSrc & vbCrLf & "        <ns14:importType>hf.cor.hg.bg.vo.svo." & fService3() & ".HgbgSel" & s사업자Biz축소 & "OutSvo</ns14:importType>"
sSrc = sSrc & vbCrLf & "        <ns14:importType>hf.cor.hg.bg.vo.svo." & fService3() & ".HgbgPrcs" & s사업자Biz축소 & "OutSvo</ns14:importType>"
sSrc = sSrc & vbCrLf & "    </ns14:wholeImport>"
sSrc = sSrc & vbCrLf & "    <ns14:serviceMethod methodName=""sel" & s사업자Biz축소 & """ displayName=""" & s사업자SVC명 & """ modifier=""public"" comments="""
sSrc = sSrc & " @메서드명 : sel" & s사업자Biz축소 & "&#xD;&#xA;"
sSrc = sSrc & " @논리명 : " & s사업자SVC명 & "&#xD;&#xA;"
sSrc = sSrc & " @입력 : HgbgSel" & s사업자Biz축소 & "InSvo&#xD;&#xA;"
sSrc = sSrc & " @출력 : HgbgSel" & s사업자Biz축소 & "OutSvo"">"
sSrc = sSrc & vbCrLf & "        <ns10:flowModule>"
sSrc = sSrc & vbCrLf & "            <ns10:entryModule>"
sSrc = sSrc & vbCrLf & "                <ns10:name>" & s사업자SVC명 & "</ns10:name>"
sSrc = sSrc & vbCrLf & "                <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:id>0</ns10:id>"
sSrc = sSrc & vbCrLf & "                    <ns10:id>2</ns10:id>"
sSrc = sSrc & vbCrLf & "                    <ns10:id>4</ns10:id>"
sSrc = sSrc & vbCrLf & "                </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                <ns10:variable kind=""DTO"" name=""outBvo"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <ns10:do physicalName=""Hgbg" & s사업자Biz축소 & "01OutBvo"" logicalName=""" & s사업자SVC명 & " OutBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"
sSrc = sSrc & vbCrLf & "                </ns10:variable>"
sSrc = sSrc & vbCrLf & "                <ns10:variable kind=""DTO"" name=""inBvo"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <ns10:do physicalName=""Hgbg" & s사업자Biz축소 & "01InBvo"" logicalName=""" & s사업자SVC명 & " InBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"
sSrc = sSrc & vbCrLf & "                </ns10:variable>"
sSrc = sSrc & vbCrLf & "                <ns10:variable kind=""DTO"" name=""outSvo"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <ns10:do physicalName=""HgbgSel" & s사업자Biz축소 & "OutSvo"" logicalName=""" & s사업자SVC명 & " OutSvo"" resourcePath=""hf/cor/hg/bg/vo/svo/" & fService3() & """/>"
sSrc = sSrc & vbCrLf & "                </ns10:variable>"
sSrc = sSrc & vbCrLf & "            </ns10:entryModule>"
sSrc = sSrc & vbCrLf & "            <ns10:innerModules>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""block"" nodeId=""0"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>업무 선행 처리</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>업무 선행 처리</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>1</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition/>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""block"" nodeId=""2"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>업무 본 처리</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>업무 본 처리</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>6</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition/>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""block"" nodeId=""4"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>업무 후행 처리</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>업무 후행 처리</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>5</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition/>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "            </ns10:innerModules>"
sSrc = sSrc & vbCrLf & "            <ns10:xorModules/>"
sSrc = sSrc & vbCrLf & "            <ns10:virtualModules>"
sSrc = sSrc & vbCrLf & "                <ns10:virtualModule nodeId=""1"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>입력값 확인</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>입력값 확인</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:code>            logger.info(""#########################"" );&#xD;"
sSrc = sSrc & vbCrLf & "            logger.info(""Svc InSvo \n"" + inSvo.toString());&#xD;"
sSrc = sSrc & vbCrLf & "            logger.info(""#########################"" );&#xD;"
sSrc = sSrc & vbCrLf & "            inBvo.setRvewInqVo(inSvo.getRvewInqVo());</ns10:code>"
sSrc = sSrc & vbCrLf & "                </ns10:virtualModule>"
sSrc = sSrc & vbCrLf & "                <ns10:virtualModule nodeId=""5"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>결과리턴</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>결과리턴</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:code>         &#xD;"
sSrc = sSrc & vbCrLf & "            return outSvo;</ns10:code>"
sSrc = sSrc & vbCrLf & "                </ns10:virtualModule>"
sSrc = sSrc & vbCrLf & "            </ns10:virtualModules>"
sSrc = sSrc & vbCrLf & "            <ns10:loopModules/>"
sSrc = sSrc & vbCrLf & "            <ns10:bizMethodCalls>"
sSrc = sSrc & vbCrLf & "                <ns10:bizMethodCall nodeId=""6"">"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:bizInstanceInfo scope=""member"" variableName=""hgbg" & s사업자Biz축소 & "Biz"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                        <ns10:classInfo classPackageName=""hf.cor.hg.bg.biz.sbiz." & fService3() & """ className=""Hgbg" & s사업자Biz축소 & "Biz""/>"
sSrc = sSrc & vbCrLf & "                    </ns10:bizInstanceInfo>"
sSrc = sSrc & vbCrLf & "                    <ns10:method methodName=""sel" & s사업자Biz축소 & """ displayName=""" & s사업자SVC명 & """>"
sSrc = sSrc & vbCrLf & "                        <ns13:arguments>"
sSrc = sSrc & vbCrLf & "                            <ns10:argument>"
sSrc = sSrc & vbCrLf & "<ns10:variable kind=""DTO"" name=""inBvo"">"
sSrc = sSrc & vbCrLf & "    <ns10:do physicalName=""Hgbg" & s사업자Biz축소 & "01InBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"
sSrc = sSrc & vbCrLf & "</ns10:variable>"
sSrc = sSrc & vbCrLf & "<ns10:sequence>0</ns10:sequence>"
sSrc = sSrc & vbCrLf & "                            </ns10:argument>"
sSrc = sSrc & vbCrLf & "                        </ns13:arguments>"
sSrc = sSrc & vbCrLf & "                        <ns13:returnType kind=""DTO"">"
sSrc = sSrc & vbCrLf & "                            <ns10:do physicalName=""Hgbg" & s사업자Biz축소 & "01OutBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"
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
sSrc = sSrc & vbCrLf & "                    <ns10:do physicalName=""HgbgSel" & s사업자Biz축소 & "InSvo"" resourcePath=""hf/cor/hg/bg/vo/svo/" & fService3() & """/>"


sPname = "HgbgSel" & s사업자Biz축소 & "InSvo"
sLname = s사업자Biz축소 & " InSvo"
sPath = Replace("hf.cor.hg.bg.vo.svo." & fService3(), ".", "/")
Call subMake사업자VO


sSrc = sSrc & vbCrLf & "                </ns10:variable>"
sSrc = sSrc & vbCrLf & "                <ns10:sequence>0</ns10:sequence>"
sSrc = sSrc & vbCrLf & "            </ns10:argument>"
sSrc = sSrc & vbCrLf & "        </ns14:arguments>"
sSrc = sSrc & vbCrLf & "        <ns14:returnType kind=""DTO"" name=""arg0"">"
sSrc = sSrc & vbCrLf & "            <ns10:do physicalName=""HgbgSel" & s사업자Biz축소 & "OutSvo"" resourcePath=""hf/cor/hg/bg/vo/svo/" & fService3() & """/>"
sSrc = sSrc & vbCrLf & "        </ns14:returnType>"
sSrc = sSrc & vbCrLf & "        <ns14:classInfo classPackageName=""hf/cor/hg/bg/svc/" & fService3() & """ className=""" & s사업자서비스ID & """/>"
sSrc = sSrc & vbCrLf & "        <ns14:annotations/>"
sSrc = sSrc & vbCrLf & "    </ns14:serviceMethod>"
sSrc = sSrc & vbCrLf & "    <ns14:serviceMethod methodName=""prcs" & s사업자Biz축소 & """ displayName=""" & s사업자SVC명 & """ modifier=""public"" comments="""
sSrc = sSrc & " @메서드명 : sel" & s사업자Biz축소 & "&#xD;&#xA;"
sSrc = sSrc & " @논리명 : " & s사업자SVC명 & "&#xD;&#xA;"
sSrc = sSrc & " @입력 : HgbgSel" & s사업자Biz축소 & "InSvo&#xD;&#xA;"
sSrc = sSrc & " @출력 : HgbgSel" & s사업자Biz축소 & "OutSvo"">"
sSrc = sSrc & vbCrLf & "        <ns10:flowModule>"
sSrc = sSrc & vbCrLf & "            <ns10:entryModule>"
sSrc = sSrc & vbCrLf & "                <ns10:name>" & s사업자SVC명 & "</ns10:name>"
sSrc = sSrc & vbCrLf & "                <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:id>0</ns10:id>"
sSrc = sSrc & vbCrLf & "                    <ns10:id>2</ns10:id>"
sSrc = sSrc & vbCrLf & "                    <ns10:id>3</ns10:id>"
sSrc = sSrc & vbCrLf & "                </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                <ns10:variable kind=""DTO"" name=""outBvo"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <ns10:do physicalName=""Hgbg" & s사업자Biz축소 & "02OutBvo"" logicalName=""" & s사업자SVC명 & " OutBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"
sSrc = sSrc & vbCrLf & "                </ns10:variable>"
sSrc = sSrc & vbCrLf & "                <ns10:variable kind=""DTO"" name=""inBvo"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <ns10:do physicalName=""Hgbg" & s사업자Biz축소 & "02InBvo"" logicalName=""" & s사업자SVC명 & " InBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"
sSrc = sSrc & vbCrLf & "                </ns10:variable>"
sSrc = sSrc & vbCrLf & "                <ns10:variable kind=""DTO"" name=""outSvo"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <ns10:do physicalName=""HgbgPrcs" & s사업자Biz축소 & "OutSvo"" logicalName=""" & s사업자SVC명 & "리 OutSvo"" resourcePath=""hf/cor/hg/bg/vo/svo/" & fService3() & """/>"
sSrc = sSrc & vbCrLf & "                </ns10:variable>"
sSrc = sSrc & vbCrLf & "            </ns10:entryModule>"
sSrc = sSrc & vbCrLf & "            <ns10:innerModules>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""block"" nodeId=""0"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>업무 선행 처리</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>업무 선행 처리</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>1</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition/>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""block"" nodeId=""2"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>업무 본 처리</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>업무 본 처리</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>5</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition/>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""block"" nodeId=""3"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>업무 후행 처리</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>업무 후행 처리</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>4</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition/>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "            </ns10:innerModules>"
sSrc = sSrc & vbCrLf & "            <ns10:xorModules/>"
sSrc = sSrc & vbCrLf & "            <ns10:virtualModules>"
sSrc = sSrc & vbCrLf & "                <ns10:virtualModule nodeId=""1"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>입력값 확인</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>입력값 확인</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:code>            logger.info(""#########################"" );&#xD;"
sSrc = sSrc & vbCrLf & "            logger.info(""Svc InSvo \n"" + inSvo.toString());&#xD;"
sSrc = sSrc & vbCrLf & "            logger.info(""#########################"" );&#xD;"
sSrc = sSrc & vbCrLf & "            </ns10:code>"
sSrc = sSrc & vbCrLf & "                </ns10:virtualModule>"
sSrc = sSrc & vbCrLf & "                <ns10:virtualModule nodeId=""4"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>결과리턴</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>결과리턴</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:code>         &#xD;"
sSrc = sSrc & vbCrLf & "            return outSvo;</ns10:code>"
sSrc = sSrc & vbCrLf & "                </ns10:virtualModule>"
sSrc = sSrc & vbCrLf & "            </ns10:virtualModules>"
sSrc = sSrc & vbCrLf & "            <ns10:loopModules/>"
sSrc = sSrc & vbCrLf & "            <ns10:bizMethodCalls>"
sSrc = sSrc & vbCrLf & "                <ns10:bizMethodCall nodeId=""5"">"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:bizInstanceInfo scope=""member"" variableName=""hgbg" & s사업자Biz축소 & "Biz"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                        <ns10:classInfo classPackageName=""hf.cor.hg.bg.biz.sbiz." & fService3() & """ className=""Hgbg" & s사업자Biz축소 & "Biz""/>"
sSrc = sSrc & vbCrLf & "                    </ns10:bizInstanceInfo>"
sSrc = sSrc & vbCrLf & "                    <ns10:method methodName=""prcs" & s사업자Biz축소 & """ displayName=""" & s사업자SVC명 & """>"
sSrc = sSrc & vbCrLf & "                        <ns13:arguments>"
sSrc = sSrc & vbCrLf & "                            <ns10:argument>"
sSrc = sSrc & vbCrLf & "<ns10:variable kind=""DTO"" name=""inBvo"">"
sSrc = sSrc & vbCrLf & "    <ns10:do physicalName=""HgbgPrcs" & s사업자Biz축소 & "02InBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"
sSrc = sSrc & vbCrLf & "</ns10:variable>"
sSrc = sSrc & vbCrLf & "<ns10:sequence>0</ns10:sequence>"
sSrc = sSrc & vbCrLf & "                            </ns10:argument>"
sSrc = sSrc & vbCrLf & "                        </ns13:arguments>"
sSrc = sSrc & vbCrLf & "                        <ns13:returnType kind=""DTO"">"
sSrc = sSrc & vbCrLf & "                            <ns10:do physicalName=""Hgbg" & s사업자Biz축소 & "OutBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"
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
sSrc = sSrc & vbCrLf & "                    <ns10:do physicalName=""Hgbg" & s사업자Biz축소 & "InSvo"" resourcePath=""hf/cor/hg/bg/vo/svo/" & fService3() & """/>"
sSrc = sSrc & vbCrLf & "                </ns10:variable>"
sSrc = sSrc & vbCrLf & "                <ns10:sequence>0</ns10:sequence>"
sSrc = sSrc & vbCrLf & "            </ns10:argument>"
sSrc = sSrc & vbCrLf & "        </ns14:arguments>"
sSrc = sSrc & vbCrLf & "        <ns14:returnType kind=""DTO"" name=""arg0"">"
sSrc = sSrc & vbCrLf & "            <ns10:do physicalName=""HgbgPrcs" & s사업자Biz축소 & "OutSvo"" resourcePath=""hf/cor/hg/bg/vo/svo/" & fService3() & """/>"
sSrc = sSrc & vbCrLf & "        </ns14:returnType>"
sSrc = sSrc & vbCrLf & "        <ns14:classInfo classPackageName=""hf/cor/hg/bg/svc/" & fService3() & """ className=""" & s사업자서비스ID & """/>"
sSrc = sSrc & vbCrLf & "        <ns14:annotations/>"
sSrc = sSrc & vbCrLf & "    </ns14:serviceMethod>"
sSrc = sSrc & vbCrLf & "    <ns14:interfaces/>"
sSrc = sSrc & vbCrLf & "</ns14:serviceObject>"


                                                                                                                                                                                                                              
End With
sMake사업자SVC = sSrc
End Function


Private Function sMake사업자biz() As String
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
sSrc = sSrc & " logicalName=""" & s사업자SVC명 & """ "
sSrc = sSrc & " physicalName=""Hgbg" & s사업자Biz축소 & "Biz"" "
sSrc = sSrc & " resourceType=""BIZ_OBJECT"" "
sSrc = sSrc & " resourcePackage=""hf.cor.hg.bg.biz.sbiz." & fService3() & """ "
sSrc = sSrc & " status=""DEV"" "
sSrc = sSrc & " creator=""ns042"" "
sSrc = sSrc & " owner=""ns042"" "
sSrc = sSrc & " revision=""1"" "
sSrc = sSrc & " description=""" & s사업자SVC명 & " "
sSrc = sSrc & " wFrame"" resourcePath=""hf/cor/hg/bg/biz/sbiz/" & fService3() & """ "
sSrc = sSrc & " modifier=""ns042"" created=""2020-04-06 19:02:53"" lastUpdated=""2020-04-06 19:02:53"">"


sSrc = sSrc & vbCrLf & "    <ns13:bizMethod methodName=""sel" & s사업자Biz축소 & """ displayName=""" & s사업자SVC명 & " 조회"" modifier=""public"" comments=""@메서드명 : sel" & s사업자Biz축소 & "&#xD;&#xA;@논리명 : " & s사업자SVC명 & " 조회&#xD;&#xA;@입력 : .Hgbg" & s사업자Biz축소 & "01InBvo&#xD;&#xA;@출력 : Hgbg" & s사업자Biz축소 & "01OutBvo"">"
sSrc = sSrc & vbCrLf & "        <ns10:flowModule>"
sSrc = sSrc & vbCrLf & "            <ns10:entryModule>"
sSrc = sSrc & vbCrLf & "                <ns10:name>" & s사업자SVC명 & " 조회</ns10:name>"
sSrc = sSrc & vbCrLf & "                <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:id>0</ns10:id>"
sSrc = sSrc & vbCrLf & "                    <ns10:id>2</ns10:id>"
sSrc = sSrc & vbCrLf & "                    <ns10:id>4</ns10:id>"
sSrc = sSrc & vbCrLf & "                </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                <ns10:variable kind=""DTO"" name=""outBvo"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <ns10:do physicalName=""Hgbg" & s사업자Biz축소 & "01OutBvo"" logicalName=""" & s사업자SVC명 & " 조회 OutBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"

sPname = "Hgbg" & s사업자Biz축소 & "01OutBvo"
sLname = s사업자SVC명 & " 조회 OutBvo"
sPath = "hf/cor/hg/bg/vo/bvo/sbvo/" & fService3()
Call subMake사업자VO


sSrc = sSrc & vbCrLf & "                </ns10:variable>"
sSrc = sSrc & vbCrLf & "            </ns10:entryModule>"
sSrc = sSrc & vbCrLf & "            <ns10:innerModules>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""block"" nodeId=""0"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>업무 선행 처리</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>업무 선행 처리</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>1</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition/>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""block"" nodeId=""2"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>업무 본 처리</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>업무 본 처리</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>9</ns10:id>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>8</ns10:id>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>3</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition/>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""block"" nodeId=""4"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>업무 후행 처리</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>업무 후행 처리</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>5</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition/>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""if"" nodeId=""6"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>조회구분 = 조회</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>조회구분이 조회일경우</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>10</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition>"
sSrc = sSrc & vbCrLf & "                        <ns10:code>""1"".equals(inBvo.getRvewInqVo().getInqDivNo())</ns10:code>"
sSrc = sSrc & vbCrLf & "                    </ns10:condition>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""if"" nodeId=""7"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>조회구분 = 변경수신</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>조회구분이 변경수신일 경우 조회</ns10:comments>"
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
sSrc = sSrc & vbCrLf & "                    <ns10:name>입력값 확인</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>입력값 확인</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:code>                    logger.info(""#########################"" );&#xD;"
sSrc = sSrc & vbCrLf & "                   logger.info(""inBvo \n"" + inBvo.toString());&#xD;"
'sSrc = sSrc & vbCrLf & "                   logger.info(""#########################"" );    &#xD;"
'sSrc = sSrc & vbCrLf & "                    if(StringUtil.isEmpty(inBvo.getRvewInqVo().getCsno())){&#xD;"
'sSrc = sSrc & vbCrLf & "                        throw new Exception(""고객번호"" + ""은(는) 필수입력입니다."");                     &#xD;"
'sSrc = sSrc & vbCrLf & "                        //throw new UserException(""E00001"", e.getMessage());&#xD;"
'sSrc = sSrc & vbCrLf & "                    }&#xD;"
'sSrc = sSrc & vbCrLf & "                    if(StringUtil.isEmpty(inBvo.getRvewInqVo().getGrntNo())){&#xD;"
'sSrc = sSrc & vbCrLf & "                        throw new Exception(""보증번호"" + ""은(는) 필수입력입니다."");                     &#xD;"
'sSrc = sSrc & vbCrLf & "                        //throw new UserException(""E00001"", e.getMessage());&#xD;"
'sSrc = sSrc & vbCrLf & "                    }&#xD;"
'sSrc = sSrc & vbCrLf & "                    if(inBvo.getRvewInqVo().getRvewSeq() == 0){&#xD;"
'sSrc = sSrc & vbCrLf & "                        throw new Exception(""심사순번"" + ""은(는) 필수입력입니다."");                     &#xD;"
'sSrc = sSrc & vbCrLf & "                        //throw new UserException(""E00001"", e.getMessage());&#xD;"
'sSrc = sSrc & vbCrLf & "                    } "
sSrc = sSrc & vbCrLf & "                </ns10:code>"
sSrc = sSrc & vbCrLf & "                </ns10:virtualModule>"
sSrc = sSrc & vbCrLf & "                <ns10:virtualModule nodeId=""3"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>" & s사업자SVC명 & " 조회</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>" & s사업자SVC명 & " 조회</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:code></ns10:code>"
sSrc = sSrc & vbCrLf & "                </ns10:virtualModule>"
sSrc = sSrc & vbCrLf & "                <ns10:virtualModule nodeId=""5"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>결과리턴</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>결과리턴</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:code>                     return outBvo;</ns10:code>"
sSrc = sSrc & vbCrLf & "                </ns10:virtualModule>"
sSrc = sSrc & vbCrLf & "                <ns10:virtualModule nodeId=""9"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>변수세팅</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:code></ns10:code>"
sSrc = sSrc & vbCrLf & "                </ns10:virtualModule>"
sSrc = sSrc & vbCrLf & "                <ns10:virtualModule nodeId=""10"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>조회</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>조회</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:code>                         "
'sSrc = sSrc & vbCrLf & "                        TbHgj052mRvewetprbrf001Dvo hgj052mRvewetprbrf001Dvo = new TbHgj052mRvewetprbrf001Dvo();        // 주택신용보증_심사기업체개요기본&#xD;"
'sSrc = sSrc & vbCrLf & "                        TbHgj053mGrntrequnsell001Dvo hgj053mGrntrequnsell001Dvo = new TbHgj053mGrntrequnsell001Dvo();  // 주택신용보증_보증신청기업미분양기본&#xD;"
'sSrc = sSrc & vbCrLf & "                        TbHgj054mCnstprgsbzpl001Dvo hgj054mCnstprgsbzpl001Dvo = new TbHgj054mCnstprgsbzpl001Dvo();     // 주택신용보증_시공기업진행중사업장기본&#xD;"
'sSrc = sSrc & vbCrLf & "                                           &#xD;"
'sSrc = sSrc & vbCrLf & "                        // 주택신용보증_심사기업체개요기본 조회&#xD;"
'sSrc = sSrc & vbCrLf & "                        DataObjectMapper&lt;TbHgj052mRvewetprbrf001Dvo&gt; mapper = new DataObjectMapper&lt;TbHgj052mRvewetprbrf001Dvo&gt;( ""po7_app"", TbHgj052mRvewetprbrf001Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper.setQuery(TbHgj052mRvewetprbrf001Dao.SEL001);&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper.setParameter(""csno"", inBvo.getRvewInqVo().getCsno());         // 고객번호                  &#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper.setParameter(""rvewSeq"", inBvo.getRvewInqVo().getRvewSeq());       // 심사순번&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper.setParameter(""rqerDvcd"", ""01"");         // 신청인구분코드(01:신청인)&#xD;"
'sSrc = sSrc & vbCrLf & "                        &#xD;"
'sSrc = sSrc & vbCrLf & "                        hgj052mRvewetprbrf001Dvo = mapper.get();       // 보증신청기업 업체개요 조회&#xD;"
'sSrc = sSrc & vbCrLf & "                        outBvo.setBzepBrfVo(hgj052mRvewetprbrf001Dvo);&#xD;"
'sSrc = sSrc & vbCrLf & "                        &#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper.setParameter(""rqerDvcd"", ""02"");         // 신청인구분코드(02:시공사)&#xD;"
'sSrc = sSrc & vbCrLf & "                        hgj052mRvewetprbrf001Dvo = mapper.get();       // 시공사 업체개요 조회&#xD;"
'sSrc = sSrc & vbCrLf & "                        outBvo.setCncoBzepBrfVo(hgj052mRvewetprbrf001Dvo);&#xD;"
'sSrc = sSrc & vbCrLf & "                        &#xD;"
'sSrc = sSrc & vbCrLf & "                         &#xD;"
'sSrc = sSrc & vbCrLf & "                        // 주택신용보증_보증신청기업미분양기본 조회 &#xD;"
'sSrc = sSrc & vbCrLf & "                        // 보증신청기업 미분양명세 조회&#xD;"
'sSrc = sSrc & vbCrLf & "                        DataObjectMapper&lt;TbHgj053mGrntrequnsell001Dvo&gt; mapper2 = new DataObjectMapper&lt;TbHgj053mGrntrequnsell001Dvo&gt;( ""po7_app"", TbHgj053mGrntrequnsell001Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper2.setQuery(TbHgj053mGrntrequnsell001Dao.SELLIST001);&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper2.setParameter(""csno"", inBvo.getRvewInqVo().getCsno());            // 고객번호                  &#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper2.setParameter(""rvewSeq"", inBvo.getRvewInqVo().getRvewSeq());          // 심사순번                  &#xD;"
'sSrc = sSrc & vbCrLf & "                        outBvo.setUnSellVoList(mapper2.getForwardList());&#xD;"
'sSrc = sSrc & vbCrLf & "                            &#xD;"
'sSrc = sSrc & vbCrLf & "                        // 주택신용보증_시공기업진행중사업장기본 조회&#xD;"
'sSrc = sSrc & vbCrLf & "                        // 시공기업 진행중인 사업장 조회&#xD;"
'sSrc = sSrc & vbCrLf & "                        DataObjectMapper&lt;TbHgj054mCnstprgsbzpl001Dvo&gt; mapper3 = new DataObjectMapper&lt;TbHgj054mCnstprgsbzpl001Dvo&gt;( ""po7_app"", TbHgj054mCnstprgsbzpl001Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper3.setQuery(TbHgj054mCnstprgsbzpl001Dao.SELLIST001);&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper3.setParameter(""csno"", inBvo.getRvewInqVo().getCsno());            // 고객번호                  &#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper3.setParameter(""rvewSeq"", inBvo.getRvewInqVo().getRvewSeq());          // 심사순번                  &#xD;"
'sSrc = sSrc & vbCrLf & "                        outBvo.setCnstPrgsBzplVoList(mapper3.getForwardList());    &#xD;"
'sSrc = sSrc & vbCrLf & "                        &#xD;"
'sSrc = sSrc & vbCrLf & "                        // 보증번호에  해당하는 시공사 고객번호 조회&#xD;"
'sSrc = sSrc & vbCrLf & "                        Hgbg" & s사업자Biz축소 & "001Dvo " & s사업자Biz축소 & "001Dvo = new Hgbg" & s사업자Biz축소 & "001Dvo();&#xD;"
'sSrc = sSrc & vbCrLf & "                        DataObjectMapper&lt;Hgbg" & s사업자Biz축소 & "001Dvo&gt; mapper4 = new DataObjectMapper&lt;Hgbg" & s사업자Biz축소 & "001Dvo&gt;( ""po7_app"", Hgbg" & s사업자Biz축소 & "001Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper4.setQuery(Hgbg" & s사업자Biz축소 & "001Dao.SEL001);&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper4.setParameter(""grntNo"", inBvo.getRvewInqVo().getGrntNo());            // 보증번호                                  &#xD;"
'sSrc = sSrc & vbCrLf & "                        " & s사업자Biz축소 & "001Dvo = mapper4.get(); &#xD;"
'sSrc = sSrc & vbCrLf & "                        outBvo.setCncoVo(" & s사업자Biz축소 & "001Dvo);  &#xD;"
'sSrc = sSrc & vbCrLf & "                             &#xD;"
'sSrc = sSrc & vbCrLf & "                        // 보증신청기업 업체개요 정보가 없을 경우&#xD;"
'sSrc = sSrc & vbCrLf & "                        if(hgj052mRvewetprbrf001Dvo == null){&#xD;"
'sSrc = sSrc & vbCrLf & "                            // 변경수신 조회&#xD;"
'sSrc = sSrc & vbCrLf & "                            " & s사업자Biz축소 & "001Dvo = new Hgbg" & s사업자Biz축소 & "001Dvo();&#xD;"
'sSrc = sSrc & vbCrLf & "                            DataObjectMapper&lt;Hgbg" & s사업자Biz축소 & "001Dvo&gt; mapper5 = new DataObjectMapper&lt;Hgbg" & s사업자Biz축소 & "001Dvo&gt;( ""po7_app"", Hgbg" & s사업자Biz축소 & "001Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                            mapper5.setQuery(Hgbg" & s사업자Biz축소 & "001Dao.SEL002);&#xD;"
'sSrc = sSrc & vbCrLf & "                            mapper5.setParameter(""grntNo"", inBvo.getRvewInqVo().getGrntNo());            // 보증번호                                  &#xD;"
'sSrc = sSrc & vbCrLf & "                            " & s사업자Biz축소 & "001Dvo = mapper5.get();&#xD;"
'sSrc = sSrc & vbCrLf & "                            &#xD;"
'sSrc = sSrc & vbCrLf & "                            // 변경수신 조회결과가 있을 경우&#xD;"
'sSrc = sSrc & vbCrLf & "                            if(" & s사업자Biz축소 & "001Dvo != null){&#xD;"
'sSrc = sSrc & vbCrLf & "                                String strEtno = """";         // 기업체번호&#xD;"
'sSrc = sSrc & vbCrLf & "                                String strFinKncd = """";      // 재무종류코드&#xD;"
'sSrc = sSrc & vbCrLf & "                                String strFinFormCd = """";        // 재무형태코드&#xD;"
'sSrc = sSrc & vbCrLf & "                                String strRschSbjtCd = """";       // 조사과목코드&#xD;"
'sSrc = sSrc & vbCrLf & "                                &#xD;"
'sSrc = sSrc & vbCrLf & "                                strEtno = " & s사업자Biz축소 & "001Dvo.getEtno();    // 기업체번호&#xD;"
'sSrc = sSrc & vbCrLf & "                                strFinKncd = ""01"";                               // 재무종류코드(01: 대차대조표)&#xD;"
'sSrc = sSrc & vbCrLf & "                                strRschSbjtCd = ""100000"";                        // 재무과목코드(100000: 자산총계)&#xD;"
'sSrc = sSrc & vbCrLf & "                                &#xD;"
'sSrc = sSrc & vbCrLf & "                            }&#xD;"
'sSrc = sSrc & vbCrLf & "                                                        &#xD;"
'sSrc = sSrc & vbCrLf & "                        }          "
sSrc = sSrc & vbCrLf & "                </ns10:code>"
sSrc = sSrc & vbCrLf & "                </ns10:virtualModule>"
sSrc = sSrc & vbCrLf & "                <ns10:virtualModule nodeId=""11"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>변경수신 조회</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>변경수신 조회</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:code>                         "
'sSrc = sSrc & vbCrLf & "                        TbHgj052mRvewetprbrf001Dvo hgj052mRvewetprbrf001Dvo = new TbHgj052mRvewetprbrf001Dvo();        // 주택신용보증_심사기업체개요기본&#xD;"
'sSrc = sSrc & vbCrLf & "                        TbHgj053mGrntrequnsell001Dvo hgj053mGrntrequnsell001Dvo = new TbHgj053mGrntrequnsell001Dvo();  // 주택신용보증_보증신청기업미분양기본&#xD;"
'sSrc = sSrc & vbCrLf & "                        TbHgj054mCnstprgsbzpl001Dvo hgj054mCnstprgsbzpl001Dvo = new TbHgj054mCnstprgsbzpl001Dvo();     // 주택신용보증_시공기업진행중사업장기본&#xD;"
'sSrc = sSrc & vbCrLf & "                                           &#xD;"
'sSrc = sSrc & vbCrLf & "                        // 주택신용보증_심사기업체개요기본 조회&#xD;"
'sSrc = sSrc & vbCrLf & "                        DataObjectMapper&lt;TbHgj052mRvewetprbrf001Dvo&gt; mapper = new DataObjectMapper&lt;TbHgj052mRvewetprbrf001Dvo&gt;( ""po7_app"", TbHgj052mRvewetprbrf001Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper.setQuery(TbHgj052mRvewetprbrf001Dao.SEL001);&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper.setParameter(""csno"", inBvo.getRvewInqVo().getCsno());         // 고객번호                  &#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper.setParameter(""rvewSeq"", inBvo.getRvewInqVo().getRvewSeq());       // 심사순번&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper.setParameter(""rqerDvcd"", ""01"");         // 신청인구분코드(01:신청인)&#xD;"
'sSrc = sSrc & vbCrLf & "                        &#xD;"
'sSrc = sSrc & vbCrLf & "                        hgj052mRvewetprbrf001Dvo = mapper.get();       // 보증신청기업 업체개요 조회&#xD;"
'sSrc = sSrc & vbCrLf & "                        outBvo.setBzepBrfVo(hgj052mRvewetprbrf001Dvo);&#xD;"
'sSrc = sSrc & vbCrLf & "                        &#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper.setParameter(""rqerDvcd"", ""02"");         // 신청인구분코드(02:시공사)&#xD;"
'sSrc = sSrc & vbCrLf & "                        hgj052mRvewetprbrf001Dvo = mapper.get();       // 시공사 업체개요 조회&#xD;"
'sSrc = sSrc & vbCrLf & "                        outBvo.setCncoBzepBrfVo(hgj052mRvewetprbrf001Dvo);&#xD;"
'sSrc = sSrc & vbCrLf & "                        &#xD;"
'sSrc = sSrc & vbCrLf & "                         &#xD;"
'sSrc = sSrc & vbCrLf & "                        // 주택신용보증_보증신청기업미분양기본 조회 &#xD;"
'sSrc = sSrc & vbCrLf & "                        // 보증신청기업 미분양명세 조회&#xD;"
'sSrc = sSrc & vbCrLf & "                        DataObjectMapper&lt;TbHgj053mGrntrequnsell001Dvo&gt; mapper2 = new DataObjectMapper&lt;TbHgj053mGrntrequnsell001Dvo&gt;( ""po7_app"", TbHgj053mGrntrequnsell001Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper2.setQuery(TbHgj053mGrntrequnsell001Dao.SELLIST001);&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper2.setParameter(""csno"", inBvo.getRvewInqVo().getCsno());            // 고객번호                  &#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper2.setParameter(""rvewSeq"", inBvo.getRvewInqVo().getRvewSeq());          // 심사순번                  &#xD;"
'sSrc = sSrc & vbCrLf & "                        outBvo.setUnSellVoList(mapper2.getForwardList());&#xD;"
'sSrc = sSrc & vbCrLf & "                            &#xD;"
'sSrc = sSrc & vbCrLf & "                        // 주택신용보증_시공기업진행중사업장기본 조회&#xD;"
'sSrc = sSrc & vbCrLf & "                        // 시공기업 진행중인 사업장 조회&#xD;"
'sSrc = sSrc & vbCrLf & "                        DataObjectMapper&lt;TbHgj054mCnstprgsbzpl001Dvo&gt; mapper3 = new DataObjectMapper&lt;TbHgj054mCnstprgsbzpl001Dvo&gt;( ""po7_app"", TbHgj054mCnstprgsbzpl001Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper3.setQuery(TbHgj054mCnstprgsbzpl001Dao.SELLIST001);&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper3.setParameter(""csno"", inBvo.getRvewInqVo().getCsno());            // 고객번호                  &#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper3.setParameter(""rvewSeq"", inBvo.getRvewInqVo().getRvewSeq());          // 심사순번                  &#xD;"
'sSrc = sSrc & vbCrLf & "                        outBvo.setCnstPrgsBzplVoList(mapper3.getForwardList());    &#xD;"
'sSrc = sSrc & vbCrLf & "                        &#xD;"
'sSrc = sSrc & vbCrLf & "                        // 보증번호에  해당하는 시공사 고객번호 조회&#xD;"
'sSrc = sSrc & vbCrLf & "                        Hgbg" & s사업자Biz축소 & "001Dvo " & s사업자Biz축소 & "001Dvo = new Hgbg" & s사업자Biz축소 & "001Dvo();&#xD;"
'sSrc = sSrc & vbCrLf & "                        DataObjectMapper&lt;Hgbg" & s사업자Biz축소 & "001Dvo&gt; mapper4 = new DataObjectMapper&lt;Hgbg" & s사업자Biz축소 & "001Dvo&gt;( ""po7_app"", Hgbg" & s사업자Biz축소 & "001Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper4.setQuery(Hgbg" & s사업자Biz축소 & "001Dao.SEL001);&#xD;"
'sSrc = sSrc & vbCrLf & "                        mapper4.setParameter(""grntNo"", inBvo.getRvewInqVo().getGrntNo());            // 보증번호                                  &#xD;"
'sSrc = sSrc & vbCrLf & "                        " & s사업자Biz축소 & "001Dvo = mapper4.get(); &#xD;"
'sSrc = sSrc & vbCrLf & "                        outBvo.setCncoVo(" & s사업자Biz축소 & "001Dvo);  &#xD;"
'sSrc = sSrc & vbCrLf & "                             &#xD;"
'sSrc = sSrc & vbCrLf & "                        // 보증신청기업 업체개요 정보가 없을 경우&#xD;"
'sSrc = sSrc & vbCrLf & "                        if(hgj052mRvewetprbrf001Dvo == null){&#xD;"
'sSrc = sSrc & vbCrLf & "                            // 변경수신 조회&#xD;"
'sSrc = sSrc & vbCrLf & "                            " & s사업자Biz축소 & "001Dvo = new Hgbg" & s사업자Biz축소 & "001Dvo();&#xD;"
'sSrc = sSrc & vbCrLf & "                            DataObjectMapper&lt;Hgbg" & s사업자Biz축소 & "001Dvo&gt; mapper5 = new DataObjectMapper&lt;Hgbg" & s사업자Biz축소 & "001Dvo&gt;( ""po7_app"", Hgbg" & s사업자Biz축소 & "001Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                            mapper5.setQuery(Hgbg" & s사업자Biz축소 & "001Dao.SEL002);&#xD;"
'sSrc = sSrc & vbCrLf & "                            mapper5.setParameter(""grntNo"", inBvo.getRvewInqVo().getGrntNo());            // 보증번호                                  &#xD;"
'sSrc = sSrc & vbCrLf & "                            " & s사업자Biz축소 & "001Dvo = mapper5.get();&#xD;"
'sSrc = sSrc & vbCrLf & "                            &#xD;"
'sSrc = sSrc & vbCrLf & "                            // 변경수신 조회결과가 있을 경우&#xD;"
'sSrc = sSrc & vbCrLf & "                            if(" & s사업자Biz축소 & "001Dvo != null){&#xD;"
'sSrc = sSrc & vbCrLf & "                                String strEtno = " & s사업자Biz축소 & "001Dvo.getEtno();                 // 기업체번호&#xD;"
'sSrc = sSrc & vbCrLf & "                                String strRschBasisDy = " & s사업자Biz축소 & "001Dvo.getRschBasisDy();   // 조사기준일자&#xD;"
'sSrc = sSrc & vbCrLf & "                                &#xD;"
'sSrc = sSrc & vbCrLf & "                                String strFinKncd = """";      // 재무종류코드&#xD;"
'sSrc = sSrc & vbCrLf & "                                String strFinFormCd = """";        // 재무형태코드&#xD;"
'sSrc = sSrc & vbCrLf & "                                String strRschSbjtCd = """";       // 조사과목코드&#xD;"
'sSrc = sSrc & vbCrLf & "                                &#xD;"
'sSrc = sSrc & vbCrLf & "                                strFinKncd = ""01"";                               // 재무종류코드(01: 대차대조표)&#xD;"
'sSrc = sSrc & vbCrLf & "                                strRschSbjtCd = ""100000"";                        // 재무과목코드(100000: 자산총계)&#xD;"
'sSrc = sSrc & vbCrLf & "                                &#xD;"
'sSrc = sSrc & vbCrLf & "                                // 조사기준일에 따라 재무형태코드 세팅&#xD;"
'sSrc = sSrc & vbCrLf & "                                if(DateUtil.getDaysDiff(strRschBasisDy, ""20080503"") &lt; 0){&#xD;"
'sSrc = sSrc & vbCrLf & "                                    strFinFormCd = ""01""; // 재무형태코드(01: 2008년 이전 양식)&#xD;"
'sSrc = sSrc & vbCrLf & "                                }&#xD;"
'sSrc = sSrc & vbCrLf & "                                else if(DateUtil.getDaysDiff(strRschBasisDy, ""20170109"") &lt; 0){&#xD;"
'sSrc = sSrc & vbCrLf & "                                    strFinFormCd = ""02""; // 재무형태코드(01: 2008년 개정 양식)&#xD;"
'sSrc = sSrc & vbCrLf & "                                }&#xD;"
'sSrc = sSrc & vbCrLf & "                                else if(DateUtil.getDaysDiff(strRschBasisDy, ""20181026"") &lt; 0){&#xD;"
'sSrc = sSrc & vbCrLf & "                                    strFinFormCd = ""03""; // 재무형태코드(01: 2017년 개정 양식)&#xD;"
'sSrc = sSrc & vbCrLf & "                                }&#xD;"
'sSrc = sSrc & vbCrLf & "                                else{&#xD;"
'sSrc = sSrc & vbCrLf & "                                    strFinFormCd = ""04""; // 재무형태코드(04: 2018년 개정 양식)&#xD;"
'sSrc = sSrc & vbCrLf & "                                }                                                               &#xD;"
'sSrc = sSrc & vbCrLf & "                                &#xD;"
'sSrc = sSrc & vbCrLf & "                                // 전기 총자산&#xD;"
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
sSrc = sSrc & vbCrLf & "                    <ns10:do physicalName=""Hgbg" & s사업자Biz축소 & "01InBvo"" logicalName=""" & s사업자SVC명 & " 조회 InBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"

sPname = "Hgbg" & s사업자Biz축소 & "01InBvo"
sLname = s사업자SVC명 & " 조회 InBvo"
sPath = "hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & ""
Call subMake사업자VO

sSrc = sSrc & vbCrLf & "                </ns10:variable>"
sSrc = sSrc & vbCrLf & "                <ns10:sequence>0</ns10:sequence>"
sSrc = sSrc & vbCrLf & "            </ns10:argument>"
sSrc = sSrc & vbCrLf & "        </ns13:arguments>"
sSrc = sSrc & vbCrLf & "        <ns13:returnType kind=""DTO"">"
sSrc = sSrc & vbCrLf & "            <ns10:do physicalName=""Hgbg" & s사업자Biz축소 & "01OutBvo"" logicalName=""" & s사업자SVC명 & " 조회 OutBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"
sSrc = sSrc & vbCrLf & "        </ns13:returnType>"
sSrc = sSrc & vbCrLf & "        <ns13:classInfo classPackageName=""hf/cor/hg/bg/biz/sbiz/" & fService3() & """ className=""Hgbg" & s사업자Biz축소 & "Biz""/>"
sSrc = sSrc & vbCrLf & "        <ns13:annotations/>"
sSrc = sSrc & vbCrLf & "    </ns13:bizMethod>"
sSrc = sSrc & vbCrLf & "    <ns13:bizMethod methodName=""prcs" & s사업자Biz축소 & """ displayName=""" & s사업자SVC명 & " 처리"" modifier=""public"" comments=""@메서드명 : prcs" & s사업자Biz축소 & "&#xD;&#xA;@논리명 : " & s사업자SVC명 & " 처리&#xD;&#xA;@입력 : .Hgbg" & s사업자Biz축소 & "02InBvo&#xD;&#xA;@출력 : Hgbg" & s사업자Biz축소 & "02OutBvo"">"
sSrc = sSrc & vbCrLf & "        <ns10:flowModule>"
sSrc = sSrc & vbCrLf & "            <ns10:entryModule>"
sSrc = sSrc & vbCrLf & "                <ns10:name>" & s사업자SVC명 & " 처리</ns10:name>"
sSrc = sSrc & vbCrLf & "                <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:id>0</ns10:id>"
sSrc = sSrc & vbCrLf & "                    <ns10:id>2</ns10:id>"
sSrc = sSrc & vbCrLf & "                    <ns10:id>4</ns10:id>"
sSrc = sSrc & vbCrLf & "                </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                <ns10:variable kind=""DTO"" name=""outBvo"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <ns10:do physicalName=""Hgbg" & s사업자Biz축소 & "02OutBvo"" logicalName=""" & s사업자SVC명 & " 처리 OutBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"

sPname = "Hgbg" & s사업자Biz축소 & "02OutBvo"
sLname = s사업자SVC명 & " 처리 OutBvo"
sPath = "hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & ""
Call subMake사업자VO

sSrc = sSrc & vbCrLf & "                </ns10:variable>"
sSrc = sSrc & vbCrLf & "            </ns10:entryModule>"
sSrc = sSrc & vbCrLf & "            <ns10:innerModules>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""block"" nodeId=""0"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>업무 선행 처리</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>업무 선행 처리</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>1</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition/>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""block"" nodeId=""2"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>업무 본 처리</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>업무 본 처리</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>3</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition/>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns10:innerModule kind=""block"" nodeId=""4"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>업무 후행 처리</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>업무 후행 처리</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns10:id>5</ns10:id>"
sSrc = sSrc & vbCrLf & "                    </ns10:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns10:condition/>"
sSrc = sSrc & vbCrLf & "                </ns10:innerModule>"
sSrc = sSrc & vbCrLf & "            </ns10:innerModules>"
sSrc = sSrc & vbCrLf & "            <ns10:xorModules/>"
sSrc = sSrc & vbCrLf & "            <ns10:virtualModules>"
sSrc = sSrc & vbCrLf & "                <ns10:virtualModule nodeId=""1"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>입력값 확인</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>입력값 확인</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:code>                    logger.info(""#########################"" );&#xD;"
sSrc = sSrc & vbCrLf & "                   logger.info(""inBvo \n"" + inBvo.toString());&#xD;"
'sSrc = sSrc & vbCrLf & "                   logger.info(""#########################"" );    &#xD;"
'sSrc = sSrc & vbCrLf & "//                  if(StringUtil.isEmpty(inBvo.getRvewInqVo().getCsno())){&#xD;"
'sSrc = sSrc & vbCrLf & "//                      throw new Exception(""고객번호"" + ""은(는) 필수입력입니다."");                     &#xD;"
'sSrc = sSrc & vbCrLf & "//                      //throw new UserException(""E00001"", e.getMessage());&#xD;"
'sSrc = sSrc & vbCrLf & "//                  }&#xD;"
'sSrc = sSrc & vbCrLf & "//                  if(StringUtil.isEmpty(inBvo.getRvewInqVo().getGrntNo())){&#xD;"
'sSrc = sSrc & vbCrLf & "//                      throw new Exception(""보증번호"" + ""은(는) 필수입력입니다."");                     &#xD;"
'sSrc = sSrc & vbCrLf & "//                      //throw new UserException(""E00001"", e.getMessage());&#xD;"
'sSrc = sSrc & vbCrLf & "//                  }&#xD;"
'sSrc = sSrc & vbCrLf & "//                  if(inBvo.getRvewInqVo().getRvewSeq() == 0){&#xD;"
'sSrc = sSrc & vbCrLf & "//                      throw new Exception(""심사순번"" + ""은(는) 필수입력입니다."");                     &#xD;"
'sSrc = sSrc & vbCrLf & "//                      //throw new UserException(""E00001"", e.getMessage());&#xD;"
'sSrc = sSrc & vbCrLf & "//                  }"
sSrc = sSrc & vbCrLf & "                </ns10:code>"
sSrc = sSrc & vbCrLf & "                </ns10:virtualModule>"
sSrc = sSrc & vbCrLf & "                <ns10:virtualModule nodeId=""3"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>" & s사업자SVC명 & " 처리</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>" & s사업자SVC명 & " 처리</ns10:comments>"
sSrc = sSrc & vbCrLf & "                    <ns10:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns10:code></ns10:code>"
sSrc = sSrc & vbCrLf & "                </ns10:virtualModule>"
sSrc = sSrc & vbCrLf & "                <ns10:virtualModule nodeId=""5"">"
sSrc = sSrc & vbCrLf & "                    <ns10:name>결과리턴</ns10:name>"
sSrc = sSrc & vbCrLf & "                    <ns10:comments>결과리턴</ns10:comments>"
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
sSrc = sSrc & vbCrLf & "                    <ns10:do physicalName=""Hgbg" & s사업자Biz축소 & "02InBvo"" logicalName=""" & s사업자SVC명 & " 처리 InBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"

sPname = "Hgbg" & s사업자Biz축소 & "02InBvo"
sLname = s사업자SVC명 & " 처리 InBvo"
sPath = "hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & ""
Call subMake사업자VO

sSrc = sSrc & vbCrLf & "                </ns10:variable>"
sSrc = sSrc & vbCrLf & "                <ns10:sequence>0</ns10:sequence>"
sSrc = sSrc & vbCrLf & "            </ns10:argument>"
sSrc = sSrc & vbCrLf & "        </ns13:arguments>"
sSrc = sSrc & vbCrLf & "        <ns13:returnType kind=""DTO"">"
sSrc = sSrc & vbCrLf & "            <ns10:do physicalName=""Hgbg" & s사업자Biz축소 & "02OutBvo"" logicalName=""" & s사업자SVC명 & " 처리 OutBvo"" resourcePath=""hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & """/>"

sPname = "Hgbg" & s사업자Biz축소 & "02OutBvo"
sLname = s사업자SVC명 & " 처리 OutBvo"
sPath = "hf/cor/hg/bg/vo/bvo/sbvo/" & fService3() & ""
Call subMake사업자VO

sSrc = sSrc & vbCrLf & "        </ns13:returnType>"
sSrc = sSrc & vbCrLf & "        <ns13:classInfo classPackageName=""hf/cor/hg/bg/biz/sbiz/" & fService3() & """ className=""Hgbg" & s사업자Biz축소 & "Biz""/>"
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
sSrc = sSrc & vbCrLf & "        <ns13:importType>hf.cor.hg.bg.dao.sdao." & fService3() & ".Hgbg" & s사업자Biz축소 & "001Dao</ns13:importType>"


sPname = "Hgbg" & s사업자Biz축소 & "001Dao"
sLname = s사업자SVC명 & "001Dao"
sPath = Replace("hf.cor.hg.bg.dao.sdao." & fService3(), ".", "/")
Call subMake사업자DAO


sPname = "Hgbg" & s사업자Biz축소 & "001Dvo"
sLname = s사업자SVC명 & " 처리 InBvo"
sPath = "hg/bg/vo/dvo/sdvo/" & fService3() & ""
Call subMake사업자VO


sSrc = sSrc & vbCrLf & "        <ns13:importType>hf.cor.hg.bg.dao.tdao.j.TbHgj052mRvewetprbrf001Dao</ns13:importType>"
sSrc = sSrc & vbCrLf & "        <ns13:importType>hf.cor.hg.bg.dao.tdao.j.TbHgj053mGrntrequnsell001Dao</ns13:importType>"
sSrc = sSrc & vbCrLf & "        <ns13:importType>hf.cor.hg.bg.dao.tdao.j.TbHgj054mCnstprgsbzpl001Dao</ns13:importType>"
sSrc = sSrc & vbCrLf & "        <ns13:importType>hf.cor.hg.bg.vo.dvo.sdvo." & fService3() & ".Hgbg" & s사업자Biz축소 & "001Dvo</ns13:importType>"
sSrc = sSrc & vbCrLf & "        <ns13:importType>hf.cor.hg.bg.vo.dvo.tdvo.j.TbHgj052mRvewetprbrf001Dvo</ns13:importType>"
sSrc = sSrc & vbCrLf & "        <ns13:importType>hf.cor.hg.bg.vo.dvo.tdvo.j.TbHgj053mGrntrequnsell001Dvo</ns13:importType>"
sSrc = sSrc & vbCrLf & "        <ns13:importType>hf.cor.hg.bg.vo.dvo.tdvo.j.TbHgj054mCnstprgsbzpl001Dvo</ns13:importType>"
sSrc = sSrc & vbCrLf & "        <ns13:importType>hf2.cor.fw.sys.util.UserException</ns13:importType>"
sSrc = sSrc & vbCrLf & "    </ns13:wholeImport>"
sSrc = sSrc & vbCrLf & "    <ns13:typeParameters/>"
sSrc = sSrc & vbCrLf & "</ns13:bizObject>"


With fpBG
                                                                                                                                                                                                                
End With
sMake사업자biz = sSrc
End Function



'사업자 엑셀 읽기
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
    x = fpBG.IsExcelFile("c:\src\사업자.xls")

    ' If file is Excel file, tell user, import sheet
    ' list, and set result to y
    If x = 1 Then
        'MsgBox "File is an Excel file.", , "File Type"
        y = fpBG.GetExcelSheetList("c:\src\사업자.xls", List, listcount, "C:\ILOGFILE_SVC.TXT", handle, True)
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
'" & s메소드id & "
'" & s메소드명 & "
'" & sINBVO & "
'" & sOUTBVO & "
'
'" & sVO변수명 & "
'" & sVOid & "
'" & sVO명 & "
'" & sVO경로 & "
'
'" & right(s메소드id,4) & "
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
'소스 생성 버튼 클릭
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
    
    
    'file name 저장
    .Row = iROW_CNT: .Col = 2
    sFilename = .Value & ".bo"
    
    .Col = 1: sClassLname = .Value
    .Col = 2: sClassPname = .Value
    .Col = 3: sPackageDotName = .Value
    .Col = 4: sUserId = .Value
    .Col = 5: sUserName = .Value
    .Col = 6: sRevision = .Value
    .Col = 7: sPackageSlashName = .Value
    
    
    
    '출력할 문자열 산출
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
    '끝이 아니면 반복한다.
    If UCase(Trim(.Value)) <> "END" Then
        iROW_CNT = iROW_CNT + 1
        GoTo LOOP_START
    End If
End With

    gfMsgInf "완료되었습니다."
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
'sSrc = sSrc & "description=""@파일명 : " & sClassPName & "&#xD;&#xA;"
'sSrc = sSrc & "@논리명 : " & sClassLName & "&#xD;&#xA;"
'sSrc = sSrc & "@작성자 : " & sUserName & "(" & sUserId & ")&#xD;&#xA;--------------------------------------------------------------------&#xD;&#xA;"
'sSrc = sSrc & "수정일자 : "
'sSrc = sSrc & "수정자 : "
'sSrc = sSrc & "요청자(SR 번호) "
'sSrc = sSrc & ": 수정된 메소드 : "
'sSrc = sSrc & "수정내용&#xD;&#xA;--------------------------------------------------------------------&#xD;&#xA;"
'sSrc = sSrc & "작성일자 : 2020.03.19 : "
'sSrc = sSrc & "작성자 : " & sUserName & "(" & sUserId & "),  "
'sSrc = sSrc & "작성내용 : 최초작성"" "
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
sSrc = sSrc & "description=""@파일명 : " & sClassPname & "&#xD;&#xA;"
sSrc = sSrc & "@논리명 : " & sClassLname & "&#xD;&#xA;"
sSrc = sSrc & "@작성자 : " & sUserName & "(" & sUserId & ")&#xD;&#xA;--------------------------------------------------------------------&#xD;&#xA;"
sSrc = sSrc & "수정일자 : "
sSrc = sSrc & "수정자 : "
sSrc = sSrc & "요청자(SR 번호) "
sSrc = sSrc & ": 수정된 메소드 : "
sSrc = sSrc & "수정내용&#xD;&#xA;--------------------------------------------------------------------&#xD;&#xA;"
sSrc = sSrc & "작성일자 : 2020.03.19 : "
sSrc = sSrc & "작성자 : " & sUserName & "(" & sUserId & "),  "
sSrc = sSrc & "작성내용 : 최초작성"" "
sSrc = sSrc & vbCrLf & "resourcePath=""" & sPackageSlashName & """ "
sSrc = sSrc & vbCrLf & "modifier=""" & sUserId & """ "
sSrc = sSrc & vbCrLf & "created=""2020-03-19 15:55:23.393"" "
sSrc = sSrc & vbCrLf & "lastUpdated=""2020-03-27 17:44:02"">"



'----------------------------------------------------
'STEP1 전문수신
'----------------------------------------------------
iROW_CNT = iROW_CNT + 1: subSetMehod


sSrc = sSrc & vbCrLf & "    <bo:bizMethod methodName=""" & s메소드id & """ "
sSrc = sSrc & "displayName=""" & s메소드명 & """ modifier=""public"" "
sSrc = sSrc & "comments=""@메서드명 : " & s메소드id & "&#xD;&#xA;"
sSrc = sSrc & "@논리명 : " & s메소드명 & "&#xD;&#xA; "
sSrc = sSrc & "@입력 : " & sINBVO & "&#xD;&#xA; "
sSrc = sSrc & "@출력 : " & sOUTBVO & """> "



sSrc = sSrc & vbCrLf & "        <flow:flowModule>"
sSrc = sSrc & vbCrLf & "            <flow:entryModule>"
sSrc = sSrc & vbCrLf & "                <flow:name>" & s메소드명 & "</flow:name>"
sSrc = sSrc & vbCrLf & "                <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:id>0</flow:id>"
sSrc = sSrc & vbCrLf & "                    <flow:id>2</flow:id>"
sSrc = sSrc & vbCrLf & "                    <flow:id>4</flow:id>"
sSrc = sSrc & vbCrLf & "                </flow:nextEntries>"

iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO변수명 & """ scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO명 & """ resourcePath=""" & sVO경로 & """/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"

'vofile 생성
Call subMakeVO

iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO변수명 & """ scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO명 & """ resourcePath=""" & sVO경로 & """/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"

'vofile 생성
Call subMakeVO

'sSrc = sSrc & vbCrLf & "                <flow:variable kind=""PRIMITIVE"" name=""sTrstTlgmRespCd"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
'sSrc = sSrc & vbCrLf & "                    <flow:primitive>String</flow:primitive>"
'sSrc = sSrc & vbCrLf & "                </flow:variable>"

sSrc = sSrc & vbCrLf & "            </flow:entryModule>"
sSrc = sSrc & vbCrLf & "            <flow:innerModules>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""0"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>업무선행처리</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>1</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""2"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>업무 본 처리</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>3</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""4"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>업무후행처리</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>5</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "            </flow:innerModules>"
sSrc = sSrc & vbCrLf & "            <flow:xorModules/>"
sSrc = sSrc & vbCrLf & "            <flow:virtualModules>"
sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""1"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>입력값 확인</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                 logger.info(""  ☆★☆★☆★☆★inBvo  " & s메소드id & " = \n[""+inBvo.toString()+""] ☆★☆★☆★☆★"" );&#xD;"
'sSrc = sSrc & vbCrLf & "                    tlgmHdrBvo = inBvo.getRecvHdrVo();&#xD;"
'sSrc = sSrc & vbCrLf & "                    //수신일자&#xD;"
'sSrc = sSrc & vbCrLf & "                    tlgmHdrBvo.setRecvDy(DateUtil.getCurrentDate(2));&#xD;"
'sSrc = sSrc & vbCrLf & "                    //수신시각&#xD;"
'sSrc = sSrc & vbCrLf & "                    tlgmHdrBvo.setRecvPtm(DateUtil.getCurrentTime(""HHmmss""));&#xD;"
'sSrc = sSrc & vbCrLf & "                    &#xD;"
'sSrc = sSrc & vbCrLf & "                    if(inBvo.getRecvDataVo() == null){&#xD;"
'sSrc = sSrc & vbCrLf & "                         //세부자료유무&#xD;"
'sSrc = sSrc & vbCrLf & "                         tlgmHdrBvo.setDtalDataYn(""N"");&#xD;"
'sSrc = sSrc & vbCrLf & "                     } else {&#xD;"
'sSrc = sSrc & vbCrLf & "                         //세부자료유무&#xD;"
'sSrc = sSrc & vbCrLf & "                         tlgmHdrBvo.setDtalDataYn(""Y"");&#xD;"
'sSrc = sSrc & vbCrLf & "                         //주민법인번호&#xD;"
'sSrc = sSrc & vbCrLf & "                         tlgmHdrBvo.setJno(inBvo.getRecvDataVo().getJno()); &#xD;"
'sSrc = sSrc & vbCrLf & "                     }&#xD;"
'sSrc = sSrc & vbCrLf & "                     sTrstTlgmRespCd = HgIgCommConst.TRST_TLGM_RESP_CD_0000;</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"



sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""3"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>수신처리</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                  try{&#xD;"
'sSrc = sSrc & vbCrLf & "                         /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                         * 수신로그 생성&#xD;"
'sSrc = sSrc & vbCrLf & "                         *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                         HgigTrstTlgmMngBiz tlgmMngBiz = new HgigTrstTlgmMngBiz();&#xD;"
'sSrc = sSrc & vbCrLf & "                         HgigTrstTlgmMng01InBvo recvLogInBvo = new HgigTrstTlgmMng01InBvo();&#xD;"
'sSrc = sSrc & vbCrLf & "                         recvLogInBvo.setTrstTlgmHdrVo(tlgmHdrBvo);                      //헤더&#xD;"
'sSrc = sSrc & vbCrLf & "                         &#xD;"
'sSrc = sSrc & vbCrLf & "                         tlgmMngBiz.insRecvLog(recvLogInBvo);&#xD;"
'sSrc = sSrc & vbCrLf & "                         &#xD;"
'sSrc = sSrc & vbCrLf & "                         /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                         * 수신처리 - 보증내역 조회&#xD;"
'sSrc = sSrc & vbCrLf & "                         *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                         inBvo.setRecvHdrVo(tlgmHdrBvo);&#xD;"
'sSrc = sSrc & vbCrLf & "                         outBvo = this.selListGrntCont(inBvo);&#xD;"
'sSrc = sSrc & vbCrLf & "                         //정상응답&#xD;"
'sSrc = sSrc & vbCrLf & "                         tlgmHdrBvo.setTrstTlgmRespCd(sTrstTlgmRespCd);&#xD;"
'sSrc = sSrc & vbCrLf & "                     } catch( Exception ex){&#xD;"
'sSrc = sSrc & vbCrLf & "                         if (ex instanceof UserException){&#xD;"
'sSrc = sSrc & vbCrLf & "                             /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                             * 에러로그 정보 생성&#xD;"
'sSrc = sSrc & vbCrLf & "                             *****************************************************************/&#xD;"
'sSrc = sSrc & vbCrLf & "                             //TODO: 오류데이터 생성 부분 추가 필요함.&#xD;"
'sSrc = sSrc & vbCrLf & "                             sTrstTlgmRespCd = ((UserException) ex).getReturnCode();&#xD;"
'sSrc = sSrc & vbCrLf & "                             tlgmHdrBvo.setTrstTlgmRespCd(sTrstTlgmRespCd);&#xD;"
'sSrc = sSrc & vbCrLf & "                             //오류프로그램내용&#xD;"
'sSrc = sSrc & vbCrLf & "                             tlgmHdrBvo.setErrPgrmCont(""오류테스트"");&#xD;"
'sSrc = sSrc & vbCrLf & "                             //오류라인순번&#xD;"
'sSrc = sSrc & vbCrLf & "                             tlgmHdrBvo.setErrLineSeq(0L); &#xD;"
'sSrc = sSrc & vbCrLf & "                             //오류메시지내용&#xD;"
'sSrc = sSrc & vbCrLf & "                             tlgmHdrBvo.setErrMsgCont(((UserException) ex).getMessage()); &#xD;"
'sSrc = sSrc & vbCrLf & "                         } else {&#xD;"
'sSrc = sSrc & vbCrLf & "                             //TODO:&#xD;"
'sSrc = sSrc & vbCrLf & "                             //throw new UserException(""90000"");&#xD;"
'sSrc = sSrc & vbCrLf & "                             //throw new LiveException(""zz.error.90000"", ex);&#xD;"
'sSrc = sSrc & vbCrLf & "                             throw ex;&#xD;"
'sSrc = sSrc & vbCrLf & "                         }&#xD;"
'sSrc = sSrc & vbCrLf & "                     }&#xD;"
'sSrc = sSrc & vbCrLf & "                     /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                     * 전문송신&#xD;"
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
sSrc = sSrc & vbCrLf & "                     logger.info(""  ☆★☆★☆★☆★outBvo  " & s메소드id & " = \n [""+outBvo.toString()+""] ☆★☆★☆★☆★"" );&#xD;"
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
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO변수명 & """ comments=""0'th argument"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO명 & """ resourcePath=""" & sVO경로 & """/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"
sSrc = sSrc & vbCrLf & "                <flow:sequence>0</flow:sequence>"
sSrc = sSrc & vbCrLf & "            </flow:argument>"
sSrc = sSrc & vbCrLf & "        </bo:arguments>"

'vofile 생성
Call subMakeVO

iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "        <bo:returnType kind=""DTO"">"
sSrc = sSrc & vbCrLf & "            <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO명 & """ resourcePath=""" & sVO경로 & """/>"
sSrc = sSrc & vbCrLf & "        </bo:returnType>"
sSrc = sSrc & vbCrLf & "        <bo:classInfo classPackageName=""" & sPackageSlashName & """ className=""" & sClassPname & """/>"
sSrc = sSrc & vbCrLf & "        <bo:annotations/>"
sSrc = sSrc & vbCrLf & "    </bo:bizMethod>"
'vofile 생성
Call subMakeVO


'----------------------------------------------
'STEP2 전문 송신
'----------------------------------------------
iROW_CNT = iROW_CNT + 1: subSetMehod

sSrc = sSrc & vbCrLf & "    <bo:bizMethod methodName=""" & s메소드id & """ "
sSrc = sSrc & "displayName=""" & s메소드명 & """ modifier=""public"" "
sSrc = sSrc & "comments=""@메서드명 : " & s메소드id & "&#xD;&#xA;"
sSrc = sSrc & "@논리명 : " & s메소드명 & "&#xD;&#xA; "
sSrc = sSrc & "@입력 : " & sINBVO & "&#xD;&#xA; "
sSrc = sSrc & "@출력 : " & sOUTBVO & """> "
sSrc = sSrc & vbCrLf & "        <flow:flowModule>"
sSrc = sSrc & vbCrLf & "            <flow:entryModule>"
sSrc = sSrc & vbCrLf & "                <flow:name>" & s메소드명 & "</flow:name>"
sSrc = sSrc & vbCrLf & "                <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:id>0</flow:id>"
sSrc = sSrc & vbCrLf & "                    <flow:id>2</flow:id>"
sSrc = sSrc & vbCrLf & "                    <flow:id>4</flow:id>"
sSrc = sSrc & vbCrLf & "                </flow:nextEntries>"




'VO
iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO변수명 & """ scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO명 & """ resourcePath=""" & sVO경로 & """/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"

'vofile 생성
Call subMakeVO

'VO
iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO변수명 & """ scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO명 & """ resourcePath=""" & sVO경로 & """/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"
'vofile 생성
Call subMakeVO


sSrc = sSrc & vbCrLf & "            </flow:entryModule>"
sSrc = sSrc & vbCrLf & "            <flow:innerModules>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""0"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>업무선행처리</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>1</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""2"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>업무 본 처리</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>3</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""4"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>업무후행처리</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>5</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "            </flow:innerModules>"
sSrc = sSrc & vbCrLf & "            <flow:xorModules/>"

sSrc = sSrc & vbCrLf & "            <flow:virtualModules>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""1"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>입력값 확인</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                  logger.info(""  ☆★☆★☆★☆★inBvo  " & s메소드id & " = \n[""+inBvo.toString()+""] ☆★☆★☆★☆★"" );&#xD;"
'sSrc = sSrc & vbCrLf & "                 tlgmHdrBvo = inBvo.getSendHdrVo();&#xD;"
'sSrc = sSrc & vbCrLf & "                 //위탁전문코드&#xD;"
'sSrc = sSrc & vbCrLf & "                 tlgmHdrBvo.setTrstTlgmCd(""" & Right(s메소드id, 4) & """);</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"


sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""3"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>송신처리</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                   &#xD;"
'sSrc = sSrc & vbCrLf & "                  //오류응답인경우&#xD;"
'sSrc = sSrc & vbCrLf & "                  if(!StringUtil.isEmpty(tlgmHdrBvo.getTrstTlgmRespCd())&#xD;"
'sSrc = sSrc & vbCrLf & "                  &amp;&amp; !HgIgCommConst.TRST_TLGM_RESP_CD_0000.equals(tlgmHdrBvo.getTrstTlgmRespCd())){&#xD;"
'sSrc = sSrc & vbCrLf & "                      /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                      * 오류로그 생성&#xD;"
'sSrc = sSrc & vbCrLf & "                      *****************************************************************/&#xD;"
'sSrc = sSrc & vbCrLf & "                      HgigTrstTlgmMngBiz tlgmMngBiz = new HgigTrstTlgmMngBiz();&#xD;"
'sSrc = sSrc & vbCrLf & "                      HgigTrstTlgmMng03InBvo sendLogInBvo = new HgigTrstTlgmMng03InBvo();&#xD;"
'sSrc = sSrc & vbCrLf & "                      sendLogInBvo.setTrstTlgmHdrVo(tlgmHdrBvo);&#xD;"
'sSrc = sSrc & vbCrLf & "                      &#xD;"
'sSrc = sSrc & vbCrLf & "                      tlgmMngBiz.insErrLog(sendLogInBvo);&#xD;"
'sSrc = sSrc & vbCrLf & "                  }&#xD;"
'sSrc = sSrc & vbCrLf & "                  /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                  * 송신로그 생성&#xD;"
'sSrc = sSrc & vbCrLf & "                  *****************************************************************/&#xD;"
'sSrc = sSrc & vbCrLf & "                  HgigTrstTlgmMngBiz tlgmMngBiz = new HgigTrstTlgmMngBiz();&#xD;"
'sSrc = sSrc & vbCrLf & "                  HgigTrstTlgmMng02InBvo sendLogInBvo = new HgigTrstTlgmMng02InBvo();&#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  //헤더&#xD;"
'sSrc = sSrc & vbCrLf & "                  sendLogInBvo.setTrstTlgmHdrVo(tlgmHdrBvo);&#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  tlgmMngBiz.insSendLog(sendLogInBvo);&#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                  * 전문송신처리&#xD;"
'sSrc = sSrc & vbCrLf & "                  *****************************************************************/&#xD;"
'sSrc = sSrc & vbCrLf & "                  //TODO:</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"


sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""5"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>return</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <flow:code>&#xD;"
sSrc = sSrc & vbCrLf & "                  logger.info(""  ☆★☆★☆★☆★outBvo  " & s메소드id & " = \n [""+outBvo.toString()+""] ☆★☆★☆★☆★"" );&#xD;"
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
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO변수명 & """ comments=""0'th argument"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO명 & """ resourcePath=""hf/cor/hg/ig/vo/bvo/sbvo""/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"
sSrc = sSrc & vbCrLf & "                <flow:sequence>0</flow:sequence>"
sSrc = sSrc & vbCrLf & "            </flow:argument>"
sSrc = sSrc & vbCrLf & "        </bo:arguments>"
'vofile 생성
Call subMakeVO


iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "        <bo:returnType kind=""DTO"">"
sSrc = sSrc & vbCrLf & "            <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO명 & """ resourcePath=""" & sVO경로 & """/>"
sSrc = sSrc & vbCrLf & "        </bo:returnType>"
sSrc = sSrc & vbCrLf & "        <bo:classInfo classPackageName=""" & sPackageSlashName & """ className=""" & sClassPname & """/>"
sSrc = sSrc & vbCrLf & "        <bo:annotations/>"
sSrc = sSrc & vbCrLf & "    </bo:bizMethod>"

'vofile 생성
Call subMakeVO

'----------------------------------------------------
'STEP3 수신내용 조회
'----------------------------------------------------
iROW_CNT = iROW_CNT + 1: subSetMehod

sSrc = sSrc & vbCrLf & "    <bo:bizMethod methodName=""" & s메소드id & """ "
sSrc = sSrc & "displayName=""" & s메소드명 & """ modifier=""public"" "
sSrc = sSrc & "comments=""@메서드명 : " & s메소드id & "&#xD;&#xA;"
sSrc = sSrc & "@논리명 : " & s메소드명 & "&#xD;&#xA; "
sSrc = sSrc & "@입력 : " & sINBVO & "&#xD;&#xA; "
sSrc = sSrc & "@출력 : " & sOUTBVO & """> "
sSrc = sSrc & vbCrLf & "        <flow:flowModule>"
sSrc = sSrc & vbCrLf & "            <flow:entryModule>"
sSrc = sSrc & vbCrLf & "                <flow:name>" & s메소드명 & "</flow:name>"
sSrc = sSrc & vbCrLf & "                <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:id>0</flow:id>"
sSrc = sSrc & vbCrLf & "                    <flow:id>2</flow:id>"
sSrc = sSrc & vbCrLf & "                    <flow:id>4</flow:id>"
sSrc = sSrc & vbCrLf & "                </flow:nextEntries>"


'VO
iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO변수명 & """ scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO명 & """ resourcePath=""" & sVO경로 & """/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"

'vofile 생성
Call subMakeVO

'VO
iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO변수명 & """ scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO명 & """ resourcePath=""" & sVO경로 & """/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"

'vofile 생성
Call subMakeVO

'VO
iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO변수명 & """ scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO명 & """ resourcePath=""" & sVO경로 & """/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"

'vofile 생성
Call subMakeVO


sSrc = sSrc & vbCrLf & "            </flow:entryModule>"
sSrc = sSrc & vbCrLf & "            <flow:innerModules>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""0"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>업무선행처리</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>1</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""2"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>업무 본 처리</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>3</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""4"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>업무후행처리</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>5</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "            </flow:innerModules>"
sSrc = sSrc & vbCrLf & "            <flow:xorModules/>"


sSrc = sSrc & vbCrLf & "            <flow:virtualModules>"
sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""1"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>입력값 확인</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                  logger.info(""  ☆★☆★☆★☆★inBvo   " & s메소드id & " = \n[""+inBvo.toString()+""] ☆★☆★☆★☆★"" );&#xD;"
'sSrc = sSrc & vbCrLf & "                 &#xD;"
'sSrc = sSrc & vbCrLf & "                 if(inBvo.getInqComnVo() == null){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""조회조건"" + ""은(는) 필수입력입니다."");&#xD;"
'sSrc = sSrc & vbCrLf & "                 }&#xD;"
'sSrc = sSrc & vbCrLf & "                 //수신일자&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(StringUtil.isEmpty(inBvo.getInqComnVo().getRecvDy())){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""수신일자"" + ""은(는) 필수입력입니다."");&#xD;"
'sSrc = sSrc & vbCrLf & "                     &#xD;"
'sSrc = sSrc & vbCrLf & "                 }&#xD;"
'sSrc = sSrc & vbCrLf & "                 //수신시각&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(StringUtil.isEmpty(inBvo.getInqComnVo().getRecvPtm())){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""수신시각"" + ""은(는) 필수입력입니다."");&#xD;"
'sSrc = sSrc & vbCrLf & "                 }&#xD;"
'sSrc = sSrc & vbCrLf & "                 //기관지로코드&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(StringUtil.isEmpty(inBvo.getInqComnVo().getOrgGiroCd())){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""기관지로코드"" + ""은(는) 필수입력입니다."");&#xD;"
'sSrc = sSrc & vbCrLf & "                 }&#xD;"
'sSrc = sSrc & vbCrLf & "                 //위탁전문코드&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(StringUtil.isEmpty(inBvo.getInqComnVo().getTrstTlgmCd())){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""위탁전문코드"" + ""은(는) 필수입력입니다."");&#xD;"
'sSrc = sSrc & vbCrLf & "                 }&#xD;"
'sSrc = sSrc & vbCrLf & "                 //전문관리번호&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(StringUtil.isEmpty(inBvo.getInqComnVo().getTlgmMngNo())){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""전문관리번호"" + ""은(는) 필수입력입니다."");&#xD;"
'sSrc = sSrc & vbCrLf & "                 }&#xD;"
'sSrc = sSrc & vbCrLf & "                 //전문순번&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(inBvo.getInqComnVo().getTlgmSeq() &lt;= 0){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""전문순번"" + ""은(는) 필수입력입니다."");&#xD;"
'sSrc = sSrc & vbCrLf & "                 }</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""3"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>수신내용 조회</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                   DataObjectMapper&lt;TbHgc601lHstrstrecv001Dvo&gt; mapper = new DataObjectMapper&lt;TbHgc601lHstrstrecv001Dvo&gt;( ""po7_app"", TbHgc601lHstrstrecv001Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setQuery(TbHgc601lHstrstrecv001Dao.SEL001);&#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  //수신일자&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""recvDy"", inBvo.getInqComnVo().getRecvDy());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //수신시각&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""recvPtm"", inBvo.getInqComnVo().getRecvPtm());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //기관지로코드&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""orgGiroCd"", inBvo.getInqComnVo().getOrgGiroCd());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //위탁전문코드&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""trstTlgmCd"", inBvo.getInqComnVo().getTrstTlgmCd());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //전문관리번호&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""tlgmMngNo"", inBvo.getInqComnVo().getTlgmMngNo());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //전문순번&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""tlgmSeq"", inBvo.getInqComnVo().getTlgmSeq());&#xD;"
'sSrc = sSrc & vbCrLf & "&#xD;"
'sSrc = sSrc & vbCrLf & "                  TbHgc601lHstrstrecv001Dvo reDvo = mapper.get();&#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                  * HEAD 조립&#xD;"
'sSrc = sSrc & vbCrLf & "                  *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                  //수신일자 &#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setRecvDy(reDvo.getRecvDy());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //수신시각 &#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setRecvPtm(reDvo.getRecvPtm());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //기관지로코드 &#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setOrgGiroCd(reDvo.getOrgGiroCd());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //송신일자 &#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setSendDy(reDvo.getSendDy());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //송신시각 &#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setSendPtm(reDvo.getSendPtm());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //위탁전문코드&#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setTrstTlgmCd(reDvo.getTrstTlgmCd());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //수탁기관명   &#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setCnsgOrgNm(reDvo.getCnsgOrgNm());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //전문관리번호&#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setTlgmMngNo(reDvo.getTlgmMngNo());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //위탁전문응답코드&#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setTrstTlgmRespCd(reDvo.getTrstTlgmRespCd());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //수신데이터내용암호화&#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setRecvDataContEnc(reDvo.getRecvDataContEnc());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //전문길이&#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setTlgmLen(reDvo.getTlgmLen());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //오류라인순번&#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setErrLineSeq(reDvo.getErrLineSeq());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //오류프로그램내용&#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setErrPgrmCont(reDvo.getErrPgrmCont());&#xD;"
'sSrc = sSrc & vbCrLf & "&#xD;"
'sSrc = sSrc & vbCrLf & "                  /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                  * DATA 조립&#xD;"
'sSrc = sSrc & vbCrLf & "                  *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                  //주민번호&#xD;"
'sSrc = sSrc & vbCrLf & "                  dataVo.setJno(reDvo.getRecvDataContEnc().trim());</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""5"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>return</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                   outBvo.setRecvHdrVo(hdrVo);&#xD;"
'sSrc = sSrc & vbCrLf & "                  outBvo.setRecvDataVo(dataVo);&#xD;"
'sSrc = sSrc & vbCrLf & "                  logger.info(""  ☆★☆★☆★☆★outBvo   " & s메소드id & " = \n [""+outBvo.toString()+""] ☆★☆★☆★☆★"" );&#xD;"
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
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO변수명 & """ comments=""0'th argument"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO명 & """ resourcePath=""hf/cor/hg/ig/vo/bvo/sbvo""/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"
sSrc = sSrc & vbCrLf & "                <flow:sequence>0</flow:sequence>"
sSrc = sSrc & vbCrLf & "            </flow:argument>"
sSrc = sSrc & vbCrLf & "        </bo:arguments>"

'vofile 생성
Call subMakeVO

iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "        <bo:returnType kind=""DTO"">"
sSrc = sSrc & vbCrLf & "            <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO명 & """ resourcePath=""" & sVO경로 & """/>"
sSrc = sSrc & vbCrLf & "        </bo:returnType>"
sSrc = sSrc & vbCrLf & "        <bo:classInfo classPackageName=""" & sPackageSlashName & """ className=""" & sClassPname & """/>"
sSrc = sSrc & vbCrLf & "        <bo:annotations/>"
sSrc = sSrc & vbCrLf & "    </bo:bizMethod>"

'vofile 생성
Call subMakeVO


'----------------------------------------------------
'STEP4 보증내역 조회 biz 처리
'----------------------------------------------------


iROW_CNT = iROW_CNT + 1: subSetMehod

sSrc = sSrc & vbCrLf & "    <bo:bizMethod methodName=""" & s메소드id & """ "
sSrc = sSrc & "displayName=""" & s메소드명 & """ modifier=""public"" "
sSrc = sSrc & "comments=""@메서드명 : " & s메소드id & "&#xD;&#xA;"
sSrc = sSrc & "@논리명 : " & s메소드명 & "&#xD;&#xA; "
sSrc = sSrc & "@입력 : " & sINBVO & "&#xD;&#xA; "
sSrc = sSrc & "@출력 : " & sOUTBVO & """> "
sSrc = sSrc & vbCrLf & "        <flow:flowModule>"
sSrc = sSrc & vbCrLf & "            <flow:entryModule>"
sSrc = sSrc & vbCrLf & "                <flow:name>" & s메소드명 & "</flow:name>"
sSrc = sSrc & vbCrLf & "                <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:id>0</flow:id>"
sSrc = sSrc & vbCrLf & "                    <flow:id>2</flow:id>"
sSrc = sSrc & vbCrLf & "                    <flow:id>4</flow:id>"
sSrc = sSrc & vbCrLf & "                </flow:nextEntries>"


'VO
iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO변수명 & """ scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO명 & """ resourcePath=""" & sVO경로 & """/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"

'vofile 생성
Call subMakeVO

''VO
'sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""custInfoVo"" arrayKind=""none"" size=""0"" scope=""local"" initObject="""" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"" inject=""NONE"" static=""false"" final=""false"">"
'sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""HgigGrntContInq001Dvo"" logicalName=""보증내역조회001Dvo"" resourcePath=""hf/cor/hg/ig/vo/dvo/sdvo""/>"
'sSrc = sSrc & vbCrLf & "                </flow:variable>"
'
''VO
'sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""grntContList"" arrayKind=""variable"" size=""0"" scope=""local"" initObject="""" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"" inject=""NONE"" static=""false"" final=""false"">"
'sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""HgigGrntContInq002Dvo"" logicalName=""보증신청내역조회Dvo"" resourcePath=""hf/cor/hg/ig/vo/dvo/sdvo""/>"
'sSrc = sSrc & vbCrLf & "                </flow:variable>"
'
''VO
'sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""grnt1601List"" arrayKind=""variable"" size=""0"" scope=""local"" initObject="""" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"" inject=""NONE"" static=""false"" final=""false"">"
'sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""HgigTrstTlgmSendData1601GrntBvo"" logicalName=""위탁전문송신데이터1601보증내용VO"" resourcePath=""hf/cor/hg/ig/vo/bvo/comn""/>"
'sSrc = sSrc & vbCrLf & "                </flow:variable>"
'
''VO
'sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""weak1601List"" arrayKind=""variable"" size=""0"" scope=""local"" initObject="""" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"" inject=""NONE"" static=""false"" final=""false"">"
'sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""HgigTrstTlgmSendData1601WeakBvo"" logicalName=""위탁전문송신데이터1601부실자료Vo"" resourcePath=""hf/cor/hg/ig/vo/bvo/comn""/>"
'sSrc = sSrc & vbCrLf & "                </flow:variable>"
'
''VO
'sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""send1601Vo"" arrayKind=""none"" size=""0"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"" inject=""NONE"" static=""false"" final=""false"">"
'sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""HgigTrstTlgmSendData1601Bvo"" logicalName=""위탁전문송신데이터1601Vo"" resourcePath=""hf/cor/hg/ig/vo/bvo/comn""/>"
'sSrc = sSrc & vbCrLf & "                </flow:variable>"
'
''VO
'sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""yth1601List"" arrayKind=""variable"" size=""0"" scope=""local"" initObject="""" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"" final=""false"">"
'sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""HgigTrstTlgmSendData1601YthBvo"" logicalName=""위탁전문송신데이터1601청년월세Vo"" resourcePath=""hf/cor/hg/ig/vo/bvo/comn""/>"
'sSrc = sSrc & vbCrLf & "                </flow:variable>"


sSrc = sSrc & vbCrLf & "            </flow:entryModule>"
sSrc = sSrc & vbCrLf & "            <flow:innerModules>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""0"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>업무선행처리</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>1</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""2"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>업무 본 처리</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>3</flow:id>"
sSrc = sSrc & vbCrLf & "                        <flow:id>7</flow:id>"
sSrc = sSrc & vbCrLf & "                        <flow:id>10</flow:id>"
sSrc = sSrc & vbCrLf & "                        <flow:id>13</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""4"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>업무후행처리</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>5</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""if"" ifElseWhenLast=""false"" nodeId=""15"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>보증고객</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:comments>관할부점으로 일반고객, 보증고객 판단</flow:comments>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>11</flow:id>"
sSrc = sSrc & vbCrLf & "                        <flow:id>18</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition>"
sSrc = sSrc & vbCrLf & "                        <flow:code>!StringUtil.isEmpty(custInfoVo.getCtrlBrofCd())</flow:code>"
sSrc = sSrc & vbCrLf & "                    </flow:condition>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""7"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>보증내역조회</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>6</flow:id>"
sSrc = sSrc & vbCrLf & "                        <flow:id>8</flow:id>"
sSrc = sSrc & vbCrLf & "                        <flow:id>9</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""13"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>공통정보</flow:name>"
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
sSrc = sSrc & vbCrLf & "                    <flow:name>입력값 확인</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'Src = sSrc & vbCrLf & "                    <flow:code>                  logger.info(""  ☆★☆★☆★☆★inBvo  selListGrntCont = \n[""+inBvo.toString()+""] ☆★☆★☆★☆★"" );&#xD;"
'Src = sSrc & vbCrLf & "                 &#xD;"
'sSrc = sSrc & vbCrLf & "                 //기관지로코드&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(StringUtil.isEmpty(inBvo.getRecvHdrVo().getOrgGiroCd())){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""기관지로코드"" + ""은(는) 필수입력입니다.""); &#xD;"
'sSrc = sSrc & vbCrLf & "                     //LiveException ex = new LiveException(""hg.info.00921"", pVOs);&#xD;"
'sSrc = sSrc & vbCrLf & "                 }&#xD;"
'sSrc = sSrc & vbCrLf & "                 String sJno = inBvo.getRecvDataVo().getJno();&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(sJno.length() != 13 || ""0000000000000"".equals(sJno)){&#xD;"
'sSrc = sSrc & vbCrLf & "                     //TODO: 오류처리 변경&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new UserException(""1234"", ""주민번호"");&#xD;"
'sSrc = sSrc & vbCrLf & "                     //LiveException ex = new LiveException(""hg.info.00829"", pVOs);&#xD;"
'sSrc = sSrc & vbCrLf & "                 }</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""5"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>return</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                   //보증내역&#xD;"
'sSrc = sSrc & vbCrLf & "                  send1601Vo.setGrntContList(grnt1601List);&#xD;"
'sSrc = sSrc & vbCrLf & "                  //부실자료제출내역&#xD;"
'sSrc = sSrc & vbCrLf & "                  send1601Vo.setWeakDataList(weak1601List);&#xD;"
'sSrc = sSrc & vbCrLf & "                  //청년월세자금보증내역&#xD;"
'sSrc = sSrc & vbCrLf & "                  send1601Vo.setYthMmrtList(yth1601List);&#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  outBvo.setGrntContVo(send1601Vo);&#xD;"
'sSrc = sSrc & vbCrLf & "                  logger.info(""  ☆★☆★☆★☆★outBvo  selListGrntCont = \n [""+outBvo.toString()+""] ☆★☆★☆★☆★"" );&#xD;"
'sSrc = sSrc & vbCrLf & "                  return outBvo;</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"


sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""6"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>고객 보증신청내역 조회</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                       /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                      * 보증신청내역 조회(MAX 20건)&#xD;"
'sSrc = sSrc & vbCrLf & "                      *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                      DataObjectMapper&lt;HgigGrntContInq002Dvo&gt; mapper = new DataObjectMapper&lt;HgigGrntContInq002Dvo&gt;( ""po7_app"", HgigGrntContInq002Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                      mapper.setQuery(HgigGrntContInq001Dao.SELLIST001);&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                      //주민번호 TODO: 암호화처리&#xD;"
'sSrc = sSrc & vbCrLf & "                      mapper.setParameter(""csno"", custInfoVo.getCsno());&#xD;"
'sSrc = sSrc & vbCrLf & "                      &#xD;"
'sSrc = sSrc & vbCrLf & "                      for(HgigGrntContInq002Dvo grntContVo : mapper.getForwardList()){&#xD;"
'sSrc = sSrc & vbCrLf & "                          grntContList.add(grntContVo);&#xD;"
'sSrc = sSrc & vbCrLf & "                      }</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""8"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>공동사업자보증 신청내역 조회</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                       /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                      * 공동사업자보증보증 신청내역 조회&#xD;"
'sSrc = sSrc & vbCrLf & "                      *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                      DataObjectMapper&lt;HgigGrntContInq002Dvo&gt; mapper = new DataObjectMapper&lt;HgigGrntContInq002Dvo&gt;( ""po7_app"", HgigGrntContInq002Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                      mapper.setQuery(HgigGrntContInq001Dao.SELLIST002);&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                      //고객번호&#xD;"
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
sSrc = sSrc & vbCrLf & "                    <flow:name>보증신청내역 추가정보 조회</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                       //보증잔액&#xD;"
'sSrc = sSrc & vbCrLf & "                      BigDecimal bdGrntRamt = BigDecimal.ZERO;&#xD;"
'sSrc = sSrc & vbCrLf & "                      hf.cor.hg.ig.vo.bvo.comn.HgigTrstTlgmSendData1601GrntBvo grnt1601Vo = null;&#xD;"
'sSrc = sSrc & vbCrLf & "                      for (int iGrntCont = 0; iGrntCont &lt; grntContList.size(); iGrntCont++) {&#xD;"
'sSrc = sSrc & vbCrLf & "                          grnt1601Vo = new hf.cor.hg.ig.vo.bvo.comn.HgigTrstTlgmSendData1601GrntBvo();&#xD;"
'sSrc = sSrc & vbCrLf & "                          //보증잔액&#xD;"
'sSrc = sSrc & vbCrLf & "                          bdGrntRamt = grntContList.get(iGrntCont).getGrntRamt();&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                          /* 1.보증(해지)일자 */&#xD;"
'sSrc = sSrc & vbCrLf & "                          if(bdGrntRamt.compareTo(BigDecimal.ZERO) &gt; 0){&#xD;"
'sSrc = sSrc & vbCrLf & "                              //보증일자(&lt;-당초보증일자)&#xD;"
'sSrc = sSrc & vbCrLf & "                              grnt1601Vo.setGrntDy(grntContList.get(iGrntCont).getOgnlGrntDy());&#xD;"
'sSrc = sSrc & vbCrLf & "                          } else {&#xD;"
'sSrc = sSrc & vbCrLf & "                              /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                              * 보증해지일자 조회&#xD;"
'sSrc = sSrc & vbCrLf & "                              *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                              DataObjectMapper&lt;HgigGrntContInq003Dvo&gt; mapper = new DataObjectMapper&lt;HgigGrntContInq003Dvo&gt;( ""po7_app"", HgigGrntContInq003Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                                  &#xD;"
'sSrc = sSrc & vbCrLf & "                              mapper.setQuery(HgigGrntContInq001Dao.SEL002);&#xD;"
'sSrc = sSrc & vbCrLf & "                                  &#xD;"
'sSrc = sSrc & vbCrLf & "                              //보증번호&#xD;"
'sSrc = sSrc & vbCrLf & "                              mapper.setParameter(""grntNo"", grntContList.get(iGrntCont).getGrntNo());&#xD;"
'sSrc = sSrc & vbCrLf & "                              &#xD;"
'sSrc = sSrc & vbCrLf & "                              HgigGrntContInq003Dvo etcInfoVo = mapper.get(); &#xD;"
'sSrc = sSrc & vbCrLf & "                              //보증일자(&lt;-실행해지기표일자)&#xD;"
'sSrc = sSrc & vbCrLf & "                              grnt1601Vo.setGrntDy(etcInfoVo.getExecCnclEntDy());&#xD;"
'sSrc = sSrc & vbCrLf & "                          }&#xD;"
'sSrc = sSrc & vbCrLf & "                          /* 2.보증번호 */&#xD;"
'sSrc = sSrc & vbCrLf & "                          grnt1601Vo.setGrntNo(grntContList.get(iGrntCont).getGrntNo());&#xD;"
'sSrc = sSrc & vbCrLf & "                          /* 3.보증구분코드 */&#xD;"
'sSrc = sSrc & vbCrLf & "                         grnt1601Vo.setGrntDvcd(grntContList.get(iGrntCont).getGrntDvcd());&#xD;"
'sSrc = sSrc & vbCrLf & "                         &#xD;"
'sSrc = sSrc & vbCrLf & "                         /* 4.보증잔액 */&#xD;"
'sSrc = sSrc & vbCrLf & "                         //주신보보증신청상태코드(07:발급)&#xD;"
'sSrc = sSrc & vbCrLf & "                         if(""07"".equals(grntContList.get(iGrntCont).getHgGrntReqStcd())){&#xD;"
'sSrc = sSrc & vbCrLf & "                             //보증잔액&#xD;"
'sSrc = sSrc & vbCrLf & "                             grnt1601Vo.setGrntRamt(grntContList.get(iGrntCont).getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                         } else {&#xD;"
'sSrc = sSrc & vbCrLf & "                             //보증잔액(&lt;-보증신청금액)&#xD;"
'sSrc = sSrc & vbCrLf & "                             grnt1601Vo.setGrntRamt(grntContList.get(iGrntCont).getGrntReqAmt());&#xD;"
'sSrc = sSrc & vbCrLf & "                         }&#xD;"
'sSrc = sSrc & vbCrLf & "                         &#xD;"
'sSrc = sSrc & vbCrLf & "                         /* 5.주채무과목명(&lt;-대출과목명) */&#xD;"
'sSrc = sSrc & vbCrLf & "                         grnt1601Vo.setMdbtSbjtNm(grntContList.get(iGrntCont).getLoanSbjtNm());&#xD;"
'sSrc = sSrc & vbCrLf & "                         &#xD;"
'sSrc = sSrc & vbCrLf & "                          /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                          * 물건지우편번호 조회&#xD;"
'sSrc = sSrc & vbCrLf & "                          *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                          DataObjectMapper&lt;HgigGrntContInq003Dvo&gt; mapper = new DataObjectMapper&lt;HgigGrntContInq003Dvo&gt;( ""po7_app"", HgigGrntContInq003Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                              &#xD;"
'sSrc = sSrc & vbCrLf & "                          mapper.setQuery(HgigGrntContInq001Dao.SEL003);&#xD;"
'sSrc = sSrc & vbCrLf & "                              &#xD;"
'sSrc = sSrc & vbCrLf & "                          //보증번호&#xD;"
'sSrc = sSrc & vbCrLf & "                          mapper.setParameter(""grntNo"", grntContList.get(iGrntCont).getGrntNo());&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                          HgigGrntContInq003Dvo etcInfoVo = mapper.get(); &#xD;"
'sSrc = sSrc & vbCrLf & "                         /* 6.물건지우편번호 */&#xD;"
'sSrc = sSrc & vbCrLf & "                         grnt1601Vo.setTlocZpcd(etcInfoVo.getTlocZpcd());&#xD;"
'sSrc = sSrc & vbCrLf & "                         /* 7.은행지로코드(&lt;-보증상대처지로코드) */&#xD;"
'sSrc = sSrc & vbCrLf & "                         grnt1601Vo.setBankGiroCd(grntContList.get(iGrntCont).getGrntCtpcGiroCd());       &#xD;"
'sSrc = sSrc & vbCrLf & "                         &#xD;"
'sSrc = sSrc & vbCrLf & "                         /* 보증내역 추가 */&#xD;"
'sSrc = sSrc & vbCrLf & "                         grnt1601List.add(grnt1601Vo);&#xD;"
'sSrc = sSrc & vbCrLf & "                      }</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""11"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>고객 보증잔액 조회</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                           /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                          * 보증잔액 조회&#xD;"
'sSrc = sSrc & vbCrLf & "                          *****************************************************************/&#xD;"
'sSrc = sSrc & vbCrLf & "                          DataObjectMapper&lt;HgigGrntContInq002Dvo&gt; mapper = new DataObjectMapper&lt;HgigGrntContInq002Dvo&gt;( ""po7_app"", HgigGrntContInq002Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                              &#xD;"
'sSrc = sSrc & vbCrLf & "                          mapper.setQuery(HgigGrntContInq001Dao.SELLIST004);&#xD;"
'sSrc = sSrc & vbCrLf & "                              &#xD;"
'sSrc = sSrc & vbCrLf & "                          //고객번호&#xD;"
'sSrc = sSrc & vbCrLf & "                          mapper.setParameter(""csno"", custInfoVo.getCsno());&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                          //직접취득금액             10.직접취득(건축,구입,개량)&#xD;"
'sSrc = sSrc & vbCrLf & "                          BigDecimal bdDrctAcqrAmt = BigDecimal.ZERO;&#xD;"
'sSrc = sSrc & vbCrLf & "                          //직접중도금액         11.직접중도금&#xD;"
'sSrc = sSrc & vbCrLf & "                          BigDecimal bdDrctHlfwAmt = BigDecimal.ZERO;&#xD;"
'sSrc = sSrc & vbCrLf & "                          //직접임차금액         12.직접임차&#xD;"
'sSrc = sSrc & vbCrLf & "                          BigDecimal bdDrctRentAmt = BigDecimal.ZERO;&#xD;"
'sSrc = sSrc & vbCrLf & "                          //위탁취득금액         13.위탁취득(개량)&#xD;"
'sSrc = sSrc & vbCrLf & "                          BigDecimal bdTrstAcqrAmt = BigDecimal.ZERO;&#xD;"
'sSrc = sSrc & vbCrLf & "                          //위탁중도금액         14.위탁중도금&#xD;"
'sSrc = sSrc & vbCrLf & "                          BigDecimal bdTrstHlfwAmt = BigDecimal.ZERO;&#xD;"
'sSrc = sSrc & vbCrLf & "                          //위탁임차금액         15.위탁임차&#xD;"
'sSrc = sSrc & vbCrLf & "                          BigDecimal bdTrstRentAmt = BigDecimal.ZERO;&#xD;"
'sSrc = sSrc & vbCrLf & "                          for(HgigGrntContInq002Dvo grntRamtVo : mapper.getForwardList()){&#xD;"
'sSrc = sSrc & vbCrLf & "                              //직접위탁구분코드(01:직접,02:위탁)&#xD;"
'sSrc = sSrc & vbCrLf & "                              if(""01"".equals(grntRamtVo.getDrctTrstDvcd())){&#xD;"
'sSrc = sSrc & vbCrLf & "                                  //TODO: 집합코드 변경&#xD;"
'sSrc = sSrc & vbCrLf & "                                  //직접취득개량자금  CodeHandler.isValidTypeCode(""T41013"", ""410003"", guarntDvcd)&#xD;"
'sSrc = sSrc & vbCrLf & "                                  if (""직접취득개량자금"".equals(grntRamtVo.getGrntDvcd())) {&#xD;"
'sSrc = sSrc & vbCrLf & "                                      bdDrctAcqrAmt = bdDrctAcqrAmt.add(grntRamtVo.getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                                  } &#xD;"
'sSrc = sSrc & vbCrLf & "                                  //직접중도금  CodeHandler.isValidTypeCode(""T41013"", ""410004"", guarntDvcd)&#xD;"
'sSrc = sSrc & vbCrLf & "                                  else if (""직접중도금"".equals(grntRamtVo.getGrntDvcd())) {&#xD;"
'sSrc = sSrc & vbCrLf & "                                      bdDrctHlfwAmt = bdDrctAcqrAmt.add(grntRamtVo.getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                                  }&#xD;"
'sSrc = sSrc & vbCrLf & "                                  //직접임차자금  CodeHandler.isValidTypeCode(""T41013"", ""410005"", guarntDvcd)&#xD;"
'sSrc = sSrc & vbCrLf & "                                  else if (""직접임차"".equals(grntRamtVo.getGrntDvcd())) {&#xD;"
'sSrc = sSrc & vbCrLf & "                                      bdDrctRentAmt = bdDrctAcqrAmt.add(grntRamtVo.getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                                  }&#xD;"
'sSrc = sSrc & vbCrLf & "                              }&#xD;"
'sSrc = sSrc & vbCrLf & "                              //위탁분 &#xD;"
'sSrc = sSrc & vbCrLf & "                              else {&#xD;"
'sSrc = sSrc & vbCrLf & "                                  //TODO: 집합코드 변경&#xD;"
'sSrc = sSrc & vbCrLf & "                                  //위탁취득금액  CodeHandler.isValidTypeCode(""T41013"", ""410003"", guarntDvcd)&#xD;"
'sSrc = sSrc & vbCrLf & "                                  if (""위탁취득금액"".equals(grntRamtVo.getGrntDvcd())) {&#xD;"
'sSrc = sSrc & vbCrLf & "                                      bdTrstAcqrAmt = bdDrctAcqrAmt.add(grntRamtVo.getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                                  } &#xD;"
'sSrc = sSrc & vbCrLf & "                                  //위탁중도금액  CodeHandler.isValidTypeCode(""T41013"", ""410004"", guarntDvcd)&#xD;"
'sSrc = sSrc & vbCrLf & "                                  else if (""위탁중도금액"".equals(grntRamtVo.getGrntDvcd())) {&#xD;"
'sSrc = sSrc & vbCrLf & "                                      bdTrstHlfwAmt = bdDrctAcqrAmt.add(grntRamtVo.getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                                  }&#xD;"
'sSrc = sSrc & vbCrLf & "                                  //위탁임차금액  CodeHandler.isValidTypeCode(""T41013"", ""410005"", guarntDvcd)&#xD;"
'sSrc = sSrc & vbCrLf & "                                  else if (""위탁임차금액"".equals(grntRamtVo.getGrntDvcd())) {&#xD;"
'sSrc = sSrc & vbCrLf & "                                      bdTrstRentAmt = bdDrctAcqrAmt.add(grntRamtVo.getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                                  }&#xD;"
'sSrc = sSrc & vbCrLf & "                              }&#xD;"
'sSrc = sSrc & vbCrLf & "                          }&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                          //직접취득금액             10.직접취득(건축,구입,개량)&#xD;"
'sSrc = sSrc & vbCrLf & "                          send1601Vo.setDrctAcqrAmt(bdDrctAcqrAmt);&#xD;"
'sSrc = sSrc & vbCrLf & "                          //직접중도금액         11.직접중도금&#xD;"
'sSrc = sSrc & vbCrLf & "                          send1601Vo.setDrctHlfwAmt(bdDrctHlfwAmt);&#xD;"
'sSrc = sSrc & vbCrLf & "                          //직접임차금액         12.직접임차&#xD;"
'sSrc = sSrc & vbCrLf & "                          send1601Vo.setDrctRentAmt(bdDrctRentAmt);&#xD;"
'sSrc = sSrc & vbCrLf & "                          //위탁취득금액         13.위탁취득(개량)&#xD;"
'sSrc = sSrc & vbCrLf & "                          send1601Vo.setTrstAcqrAmt(bdTrstAcqrAmt);&#xD;"
'sSrc = sSrc & vbCrLf & "                          //위탁중도금액         14.위탁중도금&#xD;"
'sSrc = sSrc & vbCrLf & "                          send1601Vo.setTrstHlfwAmt(bdTrstHlfwAmt);&#xD;"
'sSrc = sSrc & vbCrLf & "                          //위탁임차금액         15.위탁임차&#xD;"
'sSrc = sSrc & vbCrLf & "                          send1601Vo.setTrstRentAmt(bdTrstRentAmt);</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"


sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""17"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>보증잔액 정보 초기값 SET</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                       /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                      * 보증잔액 정보 초기값 SET&#xD;"
'sSrc = sSrc & vbCrLf & "                      *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                      //직접취득금액             10.직접취득(건축,구입,개량)&#xD;"
'sSrc = sSrc & vbCrLf & "                      send1601Vo.setDrctAcqrAmt(BigDecimal.ZERO);&#xD;"
'sSrc = sSrc & vbCrLf & "                      //직접중도금액         11.직접중도금&#xD;"
'sSrc = sSrc & vbCrLf & "                      send1601Vo.setDrctHlfwAmt(BigDecimal.ZERO);&#xD;"
'sSrc = sSrc & vbCrLf & "                      //직접임차금액         12.직접임차&#xD;"
'sSrc = sSrc & vbCrLf & "                      send1601Vo.setDrctRentAmt(BigDecimal.ZERO);&#xD;"
'sSrc = sSrc & vbCrLf & "                      //위탁취득금액         13.위탁취득(개량)&#xD;"
'sSrc = sSrc & vbCrLf & "                      send1601Vo.setTrstAcqrAmt(BigDecimal.ZERO);&#xD;"
'sSrc = sSrc & vbCrLf & "                      //위탁중도금액         14.위탁중도금&#xD;"
'sSrc = sSrc & vbCrLf & "                      send1601Vo.setTrstHlfwAmt(BigDecimal.ZERO);&#xD;"
'sSrc = sSrc & vbCrLf & "                      //위탁임차금액         15.위탁임차&#xD;"
'sSrc = sSrc & vbCrLf & "                      send1601Vo.setTrstRentAmt(BigDecimal.ZERO);&#xD;"
'sSrc = sSrc & vbCrLf & "                      //입보총금액           16.입보총금액&#xD;"
'sSrc = sSrc & vbCrLf & "                      send1601Vo.setSrtpTotAmt(BigDecimal.ZERO);</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""18"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>공동사업자 보증잔액 조회</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                           /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                          * 공동사업자 보증잔액 조회&#xD;"
'sSrc = sSrc & vbCrLf & "                          *****************************************************************/&#xD;"
'sSrc = sSrc & vbCrLf & "                          DataObjectMapper&lt;HgigGrntContInq002Dvo&gt; mapper = new DataObjectMapper&lt;HgigGrntContInq002Dvo&gt;( ""po7_app"", HgigGrntContInq002Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                              &#xD;"
'sSrc = sSrc & vbCrLf & "                          mapper.setQuery(HgigGrntContInq001Dao.SELLIST005);&#xD;"
'sSrc = sSrc & vbCrLf & "                              &#xD;"
'sSrc = sSrc & vbCrLf & "                          //고객번호&#xD;"
'sSrc = sSrc & vbCrLf & "                          mapper.setParameter(""csno"", custInfoVo.getCsno());&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                          //직접취득금액             10.직접취득(건축,구입,개량)&#xD;"
'sSrc = sSrc & vbCrLf & "                          BigDecimal bdDrctAcqrAmt = send1601Vo.getDrctAcqrAmt();&#xD;"
'sSrc = sSrc & vbCrLf & "                          //직접중도금액         11.직접중도금&#xD;"
'sSrc = sSrc & vbCrLf & "                          BigDecimal bdDrctHlfwAmt = send1601Vo.getDrctHlfwAmt();&#xD;"
'sSrc = sSrc & vbCrLf & "                          //직접임차금액         12.직접임차&#xD;"
'sSrc = sSrc & vbCrLf & "                          BigDecimal bdDrctRentAmt = send1601Vo.getDrctRentAmt();&#xD;"
'sSrc = sSrc & vbCrLf & "                          //위탁취득금액         13.위탁취득(개량)&#xD;"
'sSrc = sSrc & vbCrLf & "                          BigDecimal bdTrstAcqrAmt = send1601Vo.getTrstAcqrAmt();&#xD;"
'sSrc = sSrc & vbCrLf & "                          //위탁중도금액         14.위탁중도금&#xD;"
'sSrc = sSrc & vbCrLf & "                          BigDecimal bdTrstHlfwAmt = send1601Vo.getTrstHlfwAmt();&#xD;"
'sSrc = sSrc & vbCrLf & "                          //위탁임차금액         15.위탁임차&#xD;"
'sSrc = sSrc & vbCrLf & "                          BigDecimal bdTrstRentAmt = send1601Vo.getTrstRentAmt();&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                          for(HgigGrntContInq002Dvo jintGrntRamtVo : mapper.getForwardList()){&#xD;"
'sSrc = sSrc & vbCrLf & "                              //직접위탁구분코드(01:직접,02:위탁)&#xD;"
'sSrc = sSrc & vbCrLf & "                              if(""01"".equals(jintGrntRamtVo.getDrctTrstDvcd())){&#xD;"
'sSrc = sSrc & vbCrLf & "                                  //TODO: 집합코드 변경&#xD;"
'sSrc = sSrc & vbCrLf & "                                  //직접취득개량자금  CodeHandler.isValidTypeCode(""T41013"", ""410003"", guarntDvcd)&#xD;"
'sSrc = sSrc & vbCrLf & "                                  if (""직접취득개량자금"".equals(jintGrntRamtVo.getGrntDvcd())) {&#xD;"
'sSrc = sSrc & vbCrLf & "                                      bdDrctAcqrAmt = bdDrctAcqrAmt.add(jintGrntRamtVo.getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                                  } &#xD;"
'sSrc = sSrc & vbCrLf & "                                  //직접중도금  CodeHandler.isValidTypeCode(""T41013"", ""410004"", guarntDvcd)&#xD;"
'sSrc = sSrc & vbCrLf & "                                  else if (""직접중도금"".equals(jintGrntRamtVo.getGrntDvcd())) {&#xD;"
'sSrc = sSrc & vbCrLf & "                                      bdDrctHlfwAmt = bdDrctAcqrAmt.add(jintGrntRamtVo.getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                                  }&#xD;"
'sSrc = sSrc & vbCrLf & "                                  //직접임차자금  CodeHandler.isValidTypeCode(""T41013"", ""410005"", guarntDvcd)&#xD;"
'sSrc = sSrc & vbCrLf & "                                  else if (""직접임차"".equals(jintGrntRamtVo.getGrntDvcd())) {&#xD;"
'sSrc = sSrc & vbCrLf & "                                      bdDrctRentAmt = bdDrctAcqrAmt.add(jintGrntRamtVo.getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                                  }&#xD;"
'sSrc = sSrc & vbCrLf & "                              }&#xD;"
'sSrc = sSrc & vbCrLf & "                              //위탁분 &#xD;"
'sSrc = sSrc & vbCrLf & "                              else {&#xD;"
'sSrc = sSrc & vbCrLf & "                                  //TODO: 집합코드 변경&#xD;"
'sSrc = sSrc & vbCrLf & "                                  //위탁취득금액  CodeHandler.isValidTypeCode(""T41013"", ""410003"", guarntDvcd)&#xD;"
'sSrc = sSrc & vbCrLf & "                                  if (""위탁취득금액"".equals(jintGrntRamtVo.getGrntDvcd())) {&#xD;"
'sSrc = sSrc & vbCrLf & "                                      bdTrstAcqrAmt = bdDrctAcqrAmt.add(jintGrntRamtVo.getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                                  } &#xD;"
'sSrc = sSrc & vbCrLf & "                                  //위탁중도금액  CodeHandler.isValidTypeCode(""T41013"", ""410004"", guarntDvcd)&#xD;"
'sSrc = sSrc & vbCrLf & "                                  else if (""위탁중도금액"".equals(jintGrntRamtVo.getGrntDvcd())) {&#xD;"
'sSrc = sSrc & vbCrLf & "                                      bdTrstHlfwAmt = bdDrctAcqrAmt.add(jintGrntRamtVo.getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                                  }&#xD;"
'sSrc = sSrc & vbCrLf & "                                  //위탁임차금액  CodeHandler.isValidTypeCode(""T41013"", ""410005"", guarntDvcd)&#xD;"
'sSrc = sSrc & vbCrLf & "                                  else if (""위탁임차금액"".equals(jintGrntRamtVo.getGrntDvcd())) {&#xD;"
'sSrc = sSrc & vbCrLf & "                                      bdTrstRentAmt = bdDrctAcqrAmt.add(jintGrntRamtVo.getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                                  }&#xD;"
'sSrc = sSrc & vbCrLf & "                              }&#xD;"
'sSrc = sSrc & vbCrLf & "                          }&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                          //직접취득금액             10.직접취득(건축,구입,개량)&#xD;"
'sSrc = sSrc & vbCrLf & "                          send1601Vo.setDrctAcqrAmt(bdDrctAcqrAmt);&#xD;"
'sSrc = sSrc & vbCrLf & "                          //직접중도금액         11.직접중도금&#xD;"
'sSrc = sSrc & vbCrLf & "                          send1601Vo.setDrctHlfwAmt(bdDrctHlfwAmt);&#xD;"
'sSrc = sSrc & vbCrLf & "                          //직접임차금액         12.직접임차&#xD;"
'sSrc = sSrc & vbCrLf & "                          send1601Vo.setDrctRentAmt(bdDrctRentAmt);&#xD;"
'sSrc = sSrc & vbCrLf & "                          //위탁취득금액         13.위탁취득(개량)&#xD;"
'sSrc = sSrc & vbCrLf & "                          send1601Vo.setTrstAcqrAmt(bdTrstAcqrAmt);&#xD;"
'sSrc = sSrc & vbCrLf & "                          //위탁중도금액         14.위탁중도금&#xD;"
'sSrc = sSrc & vbCrLf & "                          send1601Vo.setTrstHlfwAmt(bdTrstHlfwAmt);&#xD;"
'sSrc = sSrc & vbCrLf & "                          //위탁임차금액         15.위탁임차&#xD;"
'sSrc = sSrc & vbCrLf & "                          send1601Vo.setTrstRentAmt(bdTrstRentAmt);</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"


sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""12"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>입보총금액 조회</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                       /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                      * 입보총금액 조회&#xD;"
'sSrc = sSrc & vbCrLf & "                      *****************************************************************/&#xD;"
'sSrc = sSrc & vbCrLf & "                      DataObjectMapper&lt;HgigGrntContInq003Dvo&gt; mapper = new DataObjectMapper&lt;HgigGrntContInq003Dvo&gt;( ""po7_app"", HgigGrntContInq003Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                      mapper.setQuery(HgigGrntContInq001Dao.SEL004);&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                      //고객번호&#xD;"
'sSrc = sSrc & vbCrLf & "                      mapper.setParameter(""csno"", custInfoVo.getCsno());&#xD;"
'sSrc = sSrc & vbCrLf & "                      &#xD;"
'sSrc = sSrc & vbCrLf & "                      HgigGrntContInq003Dvo etcInfoVo = mapper.get(); &#xD;"
'sSrc = sSrc & vbCrLf & "                      &#xD;"
'sSrc = sSrc & vbCrLf & "                      //입보총금액&#xD;"
'sSrc = sSrc & vbCrLf & "                      send1601Vo.setSrtpTotAmt(etcInfoVo.getSrtpTotAmt());                             </flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""19"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>청년월세보증 신청내역조회</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                       /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                      * 청년월세보증 신청내역조회(MAX 10건)&#xD;"
'sSrc = sSrc & vbCrLf & "                      *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                         DataObjectMapper&lt;HgigGrntContInq002Dvo&gt; mapper = new DataObjectMapper&lt;HgigGrntContInq002Dvo&gt;( ""po7_app"", HgigGrntContInq002Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                              &#xD;"
'sSrc = sSrc & vbCrLf & "                          mapper.setQuery(HgigGrntContInq001Dao.SELLIST006);&#xD;"
'sSrc = sSrc & vbCrLf & "                              &#xD;"
'sSrc = sSrc & vbCrLf & "                          //고객번호&#xD;"
'sSrc = sSrc & vbCrLf & "                          mapper.setParameter(""csno"", custInfoVo.getCsno());&#xD;"
'sSrc = sSrc & vbCrLf & "                          &#xD;"
'sSrc = sSrc & vbCrLf & "                          hf.cor.hg.ig.vo.bvo.comn.HgigTrstTlgmSendData1601YthBvo yth1601Vo = null;&#xD;"
'sSrc = sSrc & vbCrLf & "                          for(HgigGrntContInq002Dvo ythMmrtGrntVo : mapper.getForwardList()){&#xD;"
'sSrc = sSrc & vbCrLf & "                              yth1601Vo = new hf.cor.hg.ig.vo.bvo.comn.HgigTrstTlgmSendData1601YthBvo();&#xD;"
'sSrc = sSrc & vbCrLf & "                              //1.보증번호&#xD;"
'sSrc = sSrc & vbCrLf & "                              yth1601Vo.setGrntNo(ythMmrtGrntVo.getGrntNo());&#xD;"
'sSrc = sSrc & vbCrLf & "                              //2.분할대출실행금액 합계&#xD;"
'sSrc = sSrc & vbCrLf & "                              yth1601Vo.setLoanTotAmt(ythMmrtGrntVo.getLoanTotAmt());&#xD;"
'sSrc = sSrc & vbCrLf & "                              //3.보증잔액&#xD;"
'sSrc = sSrc & vbCrLf & "                              yth1601Vo.setGrntRamt(ythMmrtGrntVo.getGrntRamt());&#xD;"
'sSrc = sSrc & vbCrLf & "                              //4.은행지로코드&#xD;"
'sSrc = sSrc & vbCrLf & "                              yth1601Vo.setBankGiroCd(ythMmrtGrntVo.getNvlGrntCtpcGiroCd());&#xD;"
'sSrc = sSrc & vbCrLf & "                              &#xD;"
'sSrc = sSrc & vbCrLf & "                              yth1601List.add(yth1601Vo);&#xD;"
'sSrc = sSrc & vbCrLf & "                          }</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""20"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>고객정보 SET</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                       /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                      * 고객 정보 SET&#xD;"
'sSrc = sSrc & vbCrLf & "                      *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                     //1.고객번호&#xD;"
'sSrc = sSrc & vbCrLf & "                     send1601Vo.setCsno(custInfoVo.getCsno());&#xD;"
'sSrc = sSrc & vbCrLf & "                     //2.주민번호&#xD;"
'sSrc = sSrc & vbCrLf & "                     send1601Vo.setJno(custInfoVo.getJcno());&#xD;"
'sSrc = sSrc & vbCrLf & "                     //3.성명&#xD;"
'sSrc = sSrc & vbCrLf & "                     send1601Vo.setCustNm(custInfoVo.getCustNm());&#xD;"
'sSrc = sSrc & vbCrLf & "                     //4.거주지 우편번호&#xD;"
'sSrc = sSrc & vbCrLf & "                     send1601Vo.setZpcd(custInfoVo.getHomeZpcd());&#xD;"
'sSrc = sSrc & vbCrLf & "                     //5.거주지 주소&#xD;"
'sSrc = sSrc & vbCrLf & "                     send1601Vo.setAddr(custInfoVo.getHomeAddr());&#xD;"
'sSrc = sSrc & vbCrLf & "                     //6.자택전화번호&#xD;"
'sSrc = sSrc & vbCrLf & "                     send1601Vo.setHomeTlno(custInfoVo.getHomeTlno());&#xD;"
'sSrc = sSrc & vbCrLf & "                     //7.휴대폰번호&#xD;"
'sSrc = sSrc & vbCrLf & "                     send1601Vo.setMbno(custInfoVo.getMbno());&#xD;"
'sSrc = sSrc & vbCrLf & "                     //8.보증관할지사&#xD;"
'sSrc = sSrc & vbCrLf & "                     send1601Vo.setCtrlBrNm(custInfoVo.getCtrlBrofCdNm());&#xD;"
'sSrc = sSrc & vbCrLf & "                     //9.보증관할지사전화번호&#xD;"
'sSrc = sSrc & vbCrLf & "                     send1601Vo.setBrTlno(custInfoVo.getBrTlno());&#xD;"
'sSrc = sSrc & vbCrLf & "                     //17.보증번호건수&#xD;"
'sSrc = sSrc & vbCrLf & "                     send1601Vo.setGrntCnt(grnt1601List.size());</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""3"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>고객정보 조회</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                   /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                  * 고객정보 조회&#xD;"
'sSrc = sSrc & vbCrLf & "                  *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                  DataObjectMapper&lt;HgigGrntContInq001Dvo&gt; mapper = new DataObjectMapper&lt;HgigGrntContInq001Dvo&gt;( ""po7_app"", HgigGrntContInq001Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                      &#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setQuery(HgigGrntContInq001Dao.SEL001);&#xD;"
'sSrc = sSrc & vbCrLf & "                      &#xD;"
'sSrc = sSrc & vbCrLf & "                  //주민법인번호 TODO: 암호화처리&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""jcno"", inBvo.getRecvDataVo().getJno());&#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  custInfoVo = mapper.get(); &#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  //정상응답으로 리턴&#xD;"
'sSrc = sSrc & vbCrLf & "                  if(custInfoVo == null){&#xD;"
'sSrc = sSrc & vbCrLf & "                      hf.cor.hg.ig.vo.bvo.comn.HgigTrstTlgmSendData1601Bvo reVo = new hf.cor.hg.ig.vo.bvo.comn.HgigTrstTlgmSendData1601Bvo();&#xD;"
'sSrc = sSrc & vbCrLf & "                      reVo.setJno(inBvo.getRecvDataVo().getJno());&#xD;"
'sSrc = sSrc & vbCrLf & "                      outBvo.setGrntContVo(reVo);&#xD;"
'sSrc = sSrc & vbCrLf & "                      return outBvo;&#xD;"
'sSrc = sSrc & vbCrLf & "                      // 해당 고객이 등록되어 있지 않습니다&#xD;"
'sSrc = sSrc & vbCrLf & "                      //throw new LiveException(""hg.info.00512"", pVOs);&#xD;"
'sSrc = sSrc & vbCrLf & "                  }</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""10"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>부실자료 제출자내역 조회</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                   /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                  * 부실자료 제출자내역 조회(MAX 10건)&#xD;"
'sSrc = sSrc & vbCrLf & "                  *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                  DataObjectMapper&lt;HgigGrntContInq004Dvo&gt; mapper = new DataObjectMapper&lt;HgigGrntContInq004Dvo&gt;( ""po7_app"", HgigGrntContInq004Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "                      &#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setQuery(HgigGrntContInq001Dao.SELLIST003);&#xD;"
'sSrc = sSrc & vbCrLf & "                      &#xD;"
'sSrc = sSrc & vbCrLf & "                  //고객번호&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""csno"", custInfoVo.getCsno());&#xD;"
'sSrc = sSrc & vbCrLf & "                  hf.cor.hg.ig.vo.bvo.comn.HgigTrstTlgmSendData1601WeakBvo weak1601Vo = null;&#xD;"
'sSrc = sSrc & vbCrLf & "                  for(HgigGrntContInq004Dvo weakDataVo : mapper.getForwardList()){&#xD;"
'sSrc = sSrc & vbCrLf & "                      weak1601Vo = new hf.cor.hg.ig.vo.bvo.comn.HgigTrstTlgmSendData1601WeakBvo();&#xD;"
'sSrc = sSrc & vbCrLf & "                      &#xD;"
'sSrc = sSrc & vbCrLf & "                      //등록일자&#xD;"
'sSrc = sSrc & vbCrLf & "                      weak1601Vo.setRegDy(weakDataVo.getInptDy());&#xD;"
'sSrc = sSrc & vbCrLf & "                      //부실자료종류코드&#xD;"
'sSrc = sSrc & vbCrLf & "                      weak1601Vo.setWeakDataKncd1(weakDataVo.getWeakDataKncd1());&#xD;"
'sSrc = sSrc & vbCrLf & "                      //부실자료종류코드&#xD;"
'sSrc = sSrc & vbCrLf & "                      weak1601Vo.setWeakDataKncd2(weakDataVo.getWeakDataKncd2());&#xD;"
'sSrc = sSrc & vbCrLf & "                      //부실자료종류코드&#xD;"
'sSrc = sSrc & vbCrLf & "                      weak1601Vo.setWeakDataKncd3(weakDataVo.getWeakDataKncd3());&#xD;"
'sSrc = sSrc & vbCrLf & "                      //기관명&#xD;"
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
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO변수명 & """ comments=""0'th argument"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO명 & """ resourcePath=""hf/cor/hg/ig/vo/bvo/sbvo""/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"
sSrc = sSrc & vbCrLf & "                <flow:sequence>0</flow:sequence>"
sSrc = sSrc & vbCrLf & "            </flow:argument>"
sSrc = sSrc & vbCrLf & "        </bo:arguments>"

'vofile 생성
Call subMakeVO

iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "        <bo:returnType kind=""DTO"">"
sSrc = sSrc & vbCrLf & "            <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO명 & """ resourcePath=""" & sVO경로 & """/>"
sSrc = sSrc & vbCrLf & "        </bo:returnType>"
sSrc = sSrc & vbCrLf & "        <bo:classInfo classPackageName=""" & sPackageSlashName & """ className=""" & sClassPname & """/>"
sSrc = sSrc & vbCrLf & "        <bo:annotations/>"
sSrc = sSrc & vbCrLf & "    </bo:bizMethod>"

'vofile 생성
Call subMakeVO

'----------------------------------------------------
'STEP5 송신내용조회
'----------------------------------------------------
iROW_CNT = iROW_CNT + 1: subSetMehod

sSrc = sSrc & vbCrLf & "    <bo:bizMethod methodName=""" & s메소드id & """ "
sSrc = sSrc & "displayName=""" & s메소드명 & """ modifier=""public"" "
sSrc = sSrc & "comments=""@메서드명 : " & s메소드id & "&#xD;&#xA;"
sSrc = sSrc & "@논리명 : " & s메소드명 & "&#xD;&#xA; "
sSrc = sSrc & "@입력 : " & sINBVO & "&#xD;&#xA; "
sSrc = sSrc & "@출력 : " & sOUTBVO & """> "
sSrc = sSrc & vbCrLf & "        <flow:flowModule>"
sSrc = sSrc & vbCrLf & "            <flow:entryModule>"
sSrc = sSrc & vbCrLf & "                <flow:name>" & s메소드명 & "</flow:name>"
sSrc = sSrc & vbCrLf & "                <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:id>0</flow:id>"
sSrc = sSrc & vbCrLf & "                    <flow:id>2</flow:id>"
sSrc = sSrc & vbCrLf & "                    <flow:id>4</flow:id>"
sSrc = sSrc & vbCrLf & "                </flow:nextEntries>"



'VO
iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO변수명 & """ scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO명 & """ resourcePath=""" & sVO경로 & """/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"

'vofile 생성
Call subMakeVO

'VO
iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO변수명 & """ scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO명 & """ resourcePath=""" & sVO경로 & """/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"

'vofile 생성
Call subMakeVO

'VO
iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO변수명 & """ scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO명 & """ resourcePath=""" & sVO경로 & """/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"

'vofile 생성
Call subMakeVO


sSrc = sSrc & vbCrLf & "            </flow:entryModule>"
sSrc = sSrc & vbCrLf & "            <flow:innerModules>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""0"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>업무선행처리</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>1</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""2"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>업무 본 처리</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>3</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "                <flow:innerModule kind=""block"" nodeId=""4"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>업무후행처리</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <flow:id>5</flow:id>"
sSrc = sSrc & vbCrLf & "                    </flow:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <flow:condition/>"
sSrc = sSrc & vbCrLf & "                </flow:innerModule>"
sSrc = sSrc & vbCrLf & "            </flow:innerModules>"
sSrc = sSrc & vbCrLf & "            <flow:xorModules/>"

sSrc = sSrc & vbCrLf & "            <flow:virtualModules>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""1"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>입력값 확인</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                  logger.info(""  ☆★☆★☆★☆★inBvo  " & s메소드id & " = \n[""+inBvo.toString()+""] ☆★☆★☆★☆★"" );&#xD;"
'sSrc = sSrc & vbCrLf & "                 &#xD;"
'sSrc = sSrc & vbCrLf & "                 if(inBvo.getInqComnVo() == null){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""조회조건"" + ""은(는) 필수입력입니다."");&#xD;"
'sSrc = sSrc & vbCrLf & "                 }&#xD;"
'sSrc = sSrc & vbCrLf & "                 //수신일자&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(StringUtil.isEmpty(inBvo.getInqComnVo().getRecvDy())){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""수신일자"" + ""은(는) 필수입력입니다."");&#xD;"
'sSrc = sSrc & vbCrLf & "                     &#xD;"
'sSrc = sSrc & vbCrLf & "                 }&#xD;"
'sSrc = sSrc & vbCrLf & "                 //수신시각&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(StringUtil.isEmpty(inBvo.getInqComnVo().getRecvPtm())){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""수신시각"" + ""은(는) 필수입력입니다."");&#xD;"
'sSrc = sSrc & vbCrLf & "                 }&#xD;"
'sSrc = sSrc & vbCrLf & "                 //기관지로코드&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(StringUtil.isEmpty(inBvo.getInqComnVo().getOrgGiroCd())){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""기관지로코드"" + ""은(는) 필수입력입니다."");&#xD;"
'sSrc = sSrc & vbCrLf & "                 }&#xD;"
'sSrc = sSrc & vbCrLf & "                 //위탁전문코드&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(StringUtil.isEmpty(inBvo.getInqComnVo().getTrstTlgmCd())){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""위탁전문코드"" + ""은(는) 필수입력입니다."");&#xD;"
'sSrc = sSrc & vbCrLf & "                 }&#xD;"
'sSrc = sSrc & vbCrLf & "                 //전문관리번호&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(StringUtil.isEmpty(inBvo.getInqComnVo().getTlgmMngNo())){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""전문관리번호"" + ""은(는) 필수입력입니다."");&#xD;"
'sSrc = sSrc & vbCrLf & "                 }&#xD;"
'sSrc = sSrc & vbCrLf & "                 //전문순번&#xD;"
'sSrc = sSrc & vbCrLf & "                 if(inBvo.getInqComnVo().getTlgmSeq() &lt;= 0){&#xD;"
'sSrc = sSrc & vbCrLf & "                     throw new Exception(""전문순번"" + ""은(는) 필수입력입니다."");&#xD;"
'sSrc = sSrc & vbCrLf & "                 }</flow:code>"
sSrc = sSrc & vbCrLf & "                </flow:virtualModule>"

sSrc = sSrc & vbCrLf & "                <flow:virtualModule nodeId=""3"">"
sSrc = sSrc & vbCrLf & "                    <flow:name>송신내용 조회</flow:name>"
sSrc = sSrc & vbCrLf & "                    <flow:nextEntries/>"
'sSrc = sSrc & vbCrLf & "                    <flow:code>                   /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                  * 송신내용 조회&#xD;"
'sSrc = sSrc & vbCrLf & "                  *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                  DataObjectMapper&lt;TbHgc602lHstrstsend001Dvo&gt; mapper = new DataObjectMapper&lt;TbHgc602lHstrstsend001Dvo&gt;( ""po7_app"", TbHgc602lHstrstsend001Dvo.class );&#xD;"
'sSrc = sSrc & vbCrLf & "&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setQuery(TbHgc602lHstrstsend001Dao.SEL001);&#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  //수신일자&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""recvDy"", inBvo.getInqComnVo().getRecvDy());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //수신시각&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""recvPtm"", inBvo.getInqComnVo().getRecvPtm());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //기관지로코드&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""orgGiroCd"", inBvo.getInqComnVo().getOrgGiroCd());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //위탁전문코드&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""trstTlgmCd"", inBvo.getInqComnVo().getTrstTlgmCd());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //전문관리번호&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""tlgmMngNo"", inBvo.getInqComnVo().getTlgmMngNo());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //전문순번&#xD;"
'sSrc = sSrc & vbCrLf & "                  mapper.setParameter(""tlgmSeq"", inBvo.getInqComnVo().getTlgmSeq());&#xD;"
'sSrc = sSrc & vbCrLf & "&#xD;"
'sSrc = sSrc & vbCrLf & "                  TbHgc602lHstrstsend001Dvo reDvo = mapper.get();&#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  if(reDvo == null){&#xD;"
'sSrc = sSrc & vbCrLf & "                      //TODO:&#xD;"
'sSrc = sSrc & vbCrLf & "                      throw new UserException(""송신내용 조회결과가 없습니다."");&#xD;"
'sSrc = sSrc & vbCrLf & "                  }&#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                  * HEAD 조립&#xD;"
'sSrc = sSrc & vbCrLf & "                  *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                  //기관지로코드 &#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setOrgGiroCd(reDvo.getOrgGiroCd());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //송신일자 &#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setSendDy(reDvo.getSendDy());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //송신시각 &#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setSendPtm(reDvo.getSendPtm());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //수탁기관명   &#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setCnsgOrgNm(reDvo.getCnsgOrgNm());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //전문관리번호&#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setTlgmMngNo(reDvo.getTlgmMngNo());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //위탁전문코드&#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setTrstTlgmCd(reDvo.getTrstTlgmCd());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //위탁전문응답코드&#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setTrstTlgmRespCd(reDvo.getTrstTlgmRespCd());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //오류라인순번&#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setErrLineSeq(reDvo.getErrLineSeq());&#xD;"
'sSrc = sSrc & vbCrLf & "                  //오류프로그램내용&#xD;"
'sSrc = sSrc & vbCrLf & "                  hdrVo.setErrPgrmCont(reDvo.getErrPgrmCont());&#xD;"
'sSrc = sSrc & vbCrLf & "&#xD;"
'sSrc = sSrc & vbCrLf & "                  /****************************************************************&#xD;"
'sSrc = sSrc & vbCrLf & "                  * DATA 조립&#xD;"
'sSrc = sSrc & vbCrLf & "                  *****************************************************************/ &#xD;"
'sSrc = sSrc & vbCrLf & "                  String sSendData = reDvo.getSendDataContEnc();&#xD;"
'sSrc = sSrc & vbCrLf & "                  byte[] btSendData = ByteUtil.toBytes(sSendData, ""utf-8"");&#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  &#xD;"
'sSrc = sSrc & vbCrLf & "                  logger.info(""  ☆★☆★☆★☆★ sSendData = \n[""+ sSendData +""]\n ☆★☆★☆★☆★"" );&#xD;"
'sSrc = sSrc & vbCrLf & "                  logger.info(""  ☆★☆★☆★☆★ btSendData = \n[""+ btSendData +""]\n ☆★☆★☆★☆★"" );&#xD;"
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
'sSrc = sSrc & vbCrLf & "                          //TODO: byteUtil 오류 개선 후 주석해제&#xD;"
'sSrc = sSrc & vbCrLf & "                          //TODO: MMS도입 후 전문으로 데이터 파싱하도록 변경&#xD;"
'sSrc = sSrc & vbCrLf & "                          if(i1602 == 1){&#xD;"
'sSrc = sSrc & vbCrLf & "                              dataVo.setCsno(sData);                 //고객번호&#xD;"
'sSrc = sSrc & vbCrLf & "                          } else if(i1602 == 2){&#xD;"
'sSrc = sSrc & vbCrLf & "                              dataVo.setJno(sData);                  //주민번호&#xD;"
'sSrc = sSrc & vbCrLf & "                          } else if(i1602 == 3){&#xD;"
'sSrc = sSrc & vbCrLf & "                              dataVo.setCustNm(sData);               //고객명&#xD;"
'sSrc = sSrc & vbCrLf & "                          } &#xD;"
'sSrc = sSrc & vbCrLf & "//                         else if(i1602 == 4){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setZpcd(sData);                 //우편번호&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 5){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setAddr(sData);                 //주소&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 6){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setHomeTlno(sData);             //자택전화번호&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 7){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setMbno(sData);                 //휴대폰번호&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 8){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setCtrlBrNm(sData);             //관할지사명&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 9){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setBrTlno(sData);               //지사전화번호&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 10){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setDrctAcqrAmt(sData);           //직접취득금액&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 11){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setDrctHlfwAmt(sData);           //직접중도금액&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 12){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setDrctRentAmt(sData);           //직접임차금액&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 13){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setTrstAcqrAmt(sData);           //위탁취득금액&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 14){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setTrstHlfwAmt(sData);           //위탁중도금액&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 15){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setTrstRentAmt(sData);           //위탁임차금액&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 16){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setSrtpTotAmt(sData);            //입보총금액&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 17){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setGrntCnt(sData);               //보증건수&#xD;"
'sSrc = sSrc & vbCrLf & "//                        } else if(i1602 == 18){&#xD;"
'sSrc = sSrc & vbCrLf & "//                            dataVo.setReptCnt(sData);               //반복건수&#xD;"
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
'sSrc = sSrc & vbCrLf & "                  logger.info(""  ☆★☆★☆★☆★outBvo  selSendCont1602 = \n [""+outBvo.toString()+""] ☆★☆★☆★☆★"" );&#xD;"
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
sSrc = sSrc & vbCrLf & "                <flow:variable kind=""DTO"" name=""" & sVO변수명 & """ comments=""0'th argument"">"
sSrc = sSrc & vbCrLf & "                    <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO명 & """ resourcePath=""hf/cor/hg/ig/vo/bvo/sbvo""/>"
sSrc = sSrc & vbCrLf & "                </flow:variable>"
sSrc = sSrc & vbCrLf & "                <flow:sequence>0</flow:sequence>"
sSrc = sSrc & vbCrLf & "            </flow:argument>"
sSrc = sSrc & vbCrLf & "        </bo:arguments>"


iROW_CNT = iROW_CNT + 1: subSetVO
sSrc = sSrc & vbCrLf & "        <bo:returnType kind=""DTO"">"
sSrc = sSrc & vbCrLf & "            <flow:do physicalName=""" & sVOid & """ logicalName=""" & sVO명 & """ resourcePath=""" & sVO경로 & """/>"
sSrc = sSrc & vbCrLf & "        </bo:returnType>"
sSrc = sSrc & vbCrLf & "        <bo:classInfo classPackageName=""" & sPackageSlashName & """ className=""" & sClassPname & """/>"
sSrc = sSrc & vbCrLf & "        <bo:annotations/>"
sSrc = sSrc & vbCrLf & "    </bo:bizMethod>"

'vofile 생성
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
    
    
    'file name 저장
    .Row = iROW_CNT: .Col = 2
    sFilename = .Value & ".so"
    
    .Col = 1: sClassLname = .Value
    .Col = 2: sClassPname = .Value
    .Col = 3: sPackageDotName = .Value
    .Col = 4: sUserId = .Value
    .Col = 5: sUserName = .Value
    .Col = 6: sRevision = .Value
    .Col = 7: sPackageSlashName = .Value
    
    
    
    '출력할 문자열 산출
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
    '끝이 아니면 반복한다.
    If UCase(Trim(.Value)) <> "END" Then
        iROW_CNT = iROW_CNT + 1
        GoTo LOOP_START
    End If
End With

    gfMsgInf "완료되었습니다."
    Exit Sub
Err_Handler:
    gfMsgWrn (Err.Description)
End Sub
'사업자 BIZ 생성
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
    .Col = 1: s사업자서비스ID = .Value
    .Col = 2: s사업자Biz축소 = .Value
    s사업자BizID = .Value & "Biz"
    .Col = 3: s사업자SVC명 = .Value
    

    
    'file name 저장
    sFilename = "Hgbg" & s사업자BizID & ".bo"
    
    
    '출력할 문자열 산출
    sPrintString = sMake사업자biz
    
    
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
    '끝이 아니면 반복한다.
    If UCase(Trim(.Value)) <> "END" Then
        'iROW_CNT = iROW_CNT + 1
        GoTo LOOP_START
    End If
End With

    gfMsgInf "완료되었습니다."
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
        sSQL = sSQL & vbCrLf & "SELECT  nvl(화면ID, 'X') as 화면ID,요구사항ID                   "
        sSQL = sSQL & vbCrLf & "  FROM 요구사항추적 "
        'sSQL = sSQL & vbCrLf & " WHERE NO =  25 "


    rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

    fpRequest.MaxRows = 0
    
    Dim iCnt As Integer
    
    iCnt = 0
    
    
    
    
    
    Dim sVar() As String
    Dim i, j, k As Integer
    rsTemp.MoveFirst
    For i = 1 To rsTemp.RecordCount
        
        sVar = Split(rsTemp.Fields("화면ID"), Chr(10))
        
        For j = LBound(sVar) To UBound(sVar)
            fpRequest.MaxRows = fpRequest.MaxRows + 1
            fpRequest.Row = fpRequest.MaxRows
            
            fpRequest.Col = 1
            
            fpRequest.Value = rsTemp.Fields("요구사항ID")
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

'메소드 관련 정보 설정 하기
Private Sub subSetMehod()
    With fpSrc
        .Row = iROW_CNT
        .Col = 2: s메소드id = .Value
        .Col = 3: s메소드명 = .Value
        .Col = 4: sINBVO = .Value
        .Col = 5: sOUTBVO = .Value
    End With
End Sub

Private Sub subSetVO()
    With fpSrc
        .Row = iROW_CNT
        .Col = 3: sVO변수명 = .Value
        .Col = 4: sVOid = .Value
        .Col = 5: sVO명 = .Value
        .Col = 6: sVO경로 = .Value
    End With
End Sub

Private Sub subSetSVCVO()
    With fpSrcSVC
        .Row = iROW_CNT
        .Col = 2: sSVC메소드id = .Value
        .Col = 3: sSVCInVOid = .Value
        .Col = 4: sSVCOutVOid = .Value

    End With
End Sub



Private Sub subMakeVO()


    Dim sSrc As String
    
    
    Dim sPname, sLname, sPath As String
    sPname = sVOid
    sLname = sVO명
    sPath = sVO경로

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
sSrc = sSrc & "description=""@파일명 : " & sClassPname & "&#xD;&#xA;"
sSrc = sSrc & "@논리명 : " & sClassLname & "&#xD;&#xA;"
sSrc = sSrc & "@작성자 : " & sUserName & "(" & sUserId & ")&#xD;&#xA;--------------------------------------------------------------------&#xD;&#xA;"
sSrc = sSrc & "수정일자 : 수정자 : 요청자(SR 번호) : 수정된 메소드 : 수정내용&#xD;&#xA;--------------------------------------------------------------------&#xD;&#xA;"
sSrc = sSrc & "작성일자 : 2020.03.19 : 작성자 : " & sUserName & "(" & sUserId & "),  "
sSrc = sSrc & "작성내용 : 최초작성"" "
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
sSrc = sSrc & vbCrLf & "        <ns13:opCode opCode=""" & sSVC메소드id & """ "
sSrc = sSrc & "svcName=""" & sSVC메소드id & """ "
sSrc = sSrc & "inputDO=""" & sSVCInVOid & """ "
sSrc = sSrc & "outputDO=""" & sSVCOutVOid & """/>"


iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "        <ns13:opCode opCode=""" & sSVC메소드id & """ "
sSrc = sSrc & "svcName=""" & sSVC메소드id & """ "
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
sSrc = sSrc & vbCrLf & "        <ns13:variable kind=""OBJECT"" name=""" & sSVC메소드id & """ scope=""member"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"" inject=""NONE"">"
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
sSrc = sSrc & vbCrLf & "    <ns13:serviceMethod methodName=""" & sSVC메소드id & """ displayName=""" & sSVCInVOid & """ modifier=""public"" comments="""">"
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
sSrc = sSrc & vbCrLf & "                    <ns9:do physicalName=""" & sSVC메소드id & """ logicalName=""" & sSVCInVOid & """ resourcePath=""" & sSVCOutVOid & """/>"
sSrc = sSrc & vbCrLf & "                </ns9:variable>"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "                <ns9:variable kind=""DTO"" name=""inBvo"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <ns9:do physicalName=""" & sSVC메소드id & """ logicalName=""" & sSVCInVOid & """ resourcePath=""" & sSVCOutVOid & """/>"
sSrc = sSrc & vbCrLf & "                </ns9:variable>"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "                <ns9:variable kind=""DTO"" name=""outSvo"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <ns9:do physicalName=""" & sSVC메소드id & """ logicalName=""" & sSVCInVOid & """ resourcePath=""" & sSVCOutVOid & """/>"
sSrc = sSrc & vbCrLf & "                </ns9:variable>"
sSrc = sSrc & vbCrLf & "            </ns9:entryModule>"

sSrc = sSrc & vbCrLf & "            <ns9:innerModules>"
sSrc = sSrc & vbCrLf & "                <ns9:innerModule kind=""block"" nodeId=""0"">"
sSrc = sSrc & vbCrLf & "                    <ns9:name>업무선행처리</ns9:name>"
sSrc = sSrc & vbCrLf & "                    <ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns9:id>1</ns9:id>"
sSrc = sSrc & vbCrLf & "                    </ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns9:condition/>"
sSrc = sSrc & vbCrLf & "                </ns9:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns9:innerModule kind=""block"" nodeId=""2"">"
sSrc = sSrc & vbCrLf & "                    <ns9:name>업무 본 처리</ns9:name>"
sSrc = sSrc & vbCrLf & "                    <ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns9:id>6</ns9:id>"
sSrc = sSrc & vbCrLf & "                    </ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns9:condition/>"
sSrc = sSrc & vbCrLf & "                </ns9:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns9:innerModule kind=""block"" nodeId=""4"">"
sSrc = sSrc & vbCrLf & "                    <ns9:name>업무후행처리</ns9:name>"
sSrc = sSrc & vbCrLf & "                    <ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns9:id>5</ns9:id>"
sSrc = sSrc & vbCrLf & "                    </ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns9:condition/>"
sSrc = sSrc & vbCrLf & "                </ns9:innerModule>"
sSrc = sSrc & vbCrLf & "            </ns9:innerModules>"
sSrc = sSrc & vbCrLf & "            <ns9:xorModules/>"
sSrc = sSrc & vbCrLf & "            <ns9:virtualModules>"
sSrc = sSrc & vbCrLf & "                <ns9:virtualModule nodeId=""1"">"
sSrc = sSrc & vbCrLf & "                    <ns9:name>입력값 확인</ns9:name>"
sSrc = sSrc & vbCrLf & "                    <ns9:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns9:code>          logger.info(""  ☆★☆★☆★☆★inSvo   = \n[""+inSvo.toString()+""] ☆★☆★☆★☆★"" );&#xD;"
sSrc = sSrc & vbCrLf & "            inBvo.setInqComnVo(inSvo.getInqComnVo());</ns9:code>"
sSrc = sSrc & vbCrLf & "                </ns9:virtualModule>"
sSrc = sSrc & vbCrLf & "                <ns9:virtualModule nodeId=""5"">"
sSrc = sSrc & vbCrLf & "                    <ns9:name>return</ns9:name>"
sSrc = sSrc & vbCrLf & "                    <ns9:nextEntries/>"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "                    <ns9:code>       " & sSVC메소드id & ";&#xD;"
sSrc = sSrc & vbCrLf & "         " & sSVCInVOid & "; &#xD;"
sSrc = sSrc & vbCrLf & "         logger.info(""  ☆★☆★☆★☆★outSvo   = \n[""+outSvo.toString()+""] ☆★☆★☆★☆★"" );&#xD;"
sSrc = sSrc & vbCrLf & "         return outSvo;</ns9:code>"
sSrc = sSrc & vbCrLf & "                </ns9:virtualModule>"
sSrc = sSrc & vbCrLf & "            </ns9:virtualModules>"
sSrc = sSrc & vbCrLf & "            <ns9:loopModules/>"
sSrc = sSrc & vbCrLf & "            <ns9:bizMethodCalls>"
sSrc = sSrc & vbCrLf & "                <ns9:bizMethodCall nodeId=""6"">"
sSrc = sSrc & vbCrLf & "                    <ns9:nextEntries/>"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "                    <ns9:bizInstanceInfo scope=""member"" variableName=""" & sSVC메소드id & """ lazy=""false"">"
sSrc = sSrc & vbCrLf & "                        <ns9:classInfo classPackageName=""" & sSVCInVOid & """ className=""" & sSVC메소드id & """/>"
sSrc = sSrc & vbCrLf & "                    </ns9:bizInstanceInfo>"


iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "                    <ns9:method methodName=""" & sSVC메소드id & """ displayName=""" & sSVCInVOid & """>"
sSrc = sSrc & vbCrLf & "                        <ns12:arguments>"
sSrc = sSrc & vbCrLf & "                            <ns9:argument>"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "<ns9:variable kind=""DTO"" name=""inBvo"">"
sSrc = sSrc & vbCrLf & "    <ns9:do physicalName=""" & sSVC메소드id & """ resourcePath=""" & sSVCOutVOid & """/>"
sSrc = sSrc & vbCrLf & "</ns9:variable>"
sSrc = sSrc & vbCrLf & "<ns9:sequence>0</ns9:sequence>"


sSrc = sSrc & vbCrLf & "                            </ns9:argument>"
sSrc = sSrc & vbCrLf & "                        </ns12:arguments>"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "                        <ns12:returnType kind=""DTO"">"
sSrc = sSrc & vbCrLf & "                            <ns9:do physicalName=""" & sSVC메소드id & """ resourcePath=""" & sSVCOutVOid & """/>"
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
sSrc = sSrc & vbCrLf & "                    <ns9:do physicalName=""" & sSVC메소드id & """ resourcePath=""" & sSVCOutVOid & """/>"
sSrc = sSrc & vbCrLf & "                </ns9:variable>"


sSrc = sSrc & vbCrLf & "                <ns9:sequence>0</ns9:sequence>"
sSrc = sSrc & vbCrLf & "            </ns9:argument>"
sSrc = sSrc & vbCrLf & "        </ns13:arguments>"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "        <ns13:returnType kind=""DTO"" name=""arg0"">"
sSrc = sSrc & vbCrLf & "            <ns9:do physicalName=""" & sSVC메소드id & """ resourcePath=""" & sSVCOutVOid & """/>"
sSrc = sSrc & vbCrLf & "        </ns13:returnType>"


sSrc = sSrc & vbCrLf & "        <ns13:classInfo classPackageName=""" & sPackageSlashName & """ className=""" & sClassPname & """/>"
sSrc = sSrc & vbCrLf & "        <ns13:annotations/>"
sSrc = sSrc & vbCrLf & "    </ns13:serviceMethod>"



'biz method
iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "    <ns13:serviceMethod methodName=""" & sSVC메소드id & """ displayName=""" & sSVCInVOid & """ modifier=""public"" comments="""">"


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
sSrc = sSrc & vbCrLf & "                    <ns9:do physicalName=""" & sSVC메소드id & """ logicalName=""" & sSVCInVOid & """ resourcePath=""" & sSVCOutVOid & """/>"
sSrc = sSrc & vbCrLf & "                </ns9:variable>"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "                <ns9:variable kind=""DTO"" name=""inBvo"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <ns9:do physicalName=""" & sSVC메소드id & """ logicalName=""" & sSVCInVOid & """ resourcePath=""" & sSVCOutVOid & """/>"
sSrc = sSrc & vbCrLf & "                </ns9:variable>"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "                <ns9:variable kind=""DTO"" name=""outSvo"" scope=""local"" modifier=""public"" getter=""false"" setter=""false"" lazy=""false"">"
sSrc = sSrc & vbCrLf & "                    <ns9:do physicalName=""" & sSVC메소드id & """ logicalName=""" & sSVCInVOid & """ resourcePath=""" & sSVCOutVOid & """/>"
sSrc = sSrc & vbCrLf & "                </ns9:variable>"
sSrc = sSrc & vbCrLf & "            </ns9:entryModule>"





sSrc = sSrc & vbCrLf & "            <ns9:innerModules>"
sSrc = sSrc & vbCrLf & "                <ns9:innerModule kind=""block"" nodeId=""0"">"
sSrc = sSrc & vbCrLf & "                    <ns9:name>업무선행처리</ns9:name>"
sSrc = sSrc & vbCrLf & "                    <ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns9:id>1</ns9:id>"
sSrc = sSrc & vbCrLf & "                    </ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns9:condition/>"
sSrc = sSrc & vbCrLf & "                </ns9:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns9:innerModule kind=""block"" nodeId=""2"">"
sSrc = sSrc & vbCrLf & "                    <ns9:name>업무 본 처리</ns9:name>"
sSrc = sSrc & vbCrLf & "                    <ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns9:id>7</ns9:id>"
sSrc = sSrc & vbCrLf & "                    </ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns9:condition/>"
sSrc = sSrc & vbCrLf & "                </ns9:innerModule>"
sSrc = sSrc & vbCrLf & "                <ns9:innerModule kind=""block"" nodeId=""4"">"
sSrc = sSrc & vbCrLf & "                    <ns9:name>업무후행처리</ns9:name>"
sSrc = sSrc & vbCrLf & "                    <ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                        <ns9:id>5</ns9:id>"
sSrc = sSrc & vbCrLf & "                    </ns9:nextEntries>"
sSrc = sSrc & vbCrLf & "                    <ns9:condition/>"
sSrc = sSrc & vbCrLf & "                </ns9:innerModule>"
sSrc = sSrc & vbCrLf & "            </ns9:innerModules>"
sSrc = sSrc & vbCrLf & "            <ns9:xorModules/>"
sSrc = sSrc & vbCrLf & "            <ns9:virtualModules>"
sSrc = sSrc & vbCrLf & "                <ns9:virtualModule nodeId=""1"">"
sSrc = sSrc & vbCrLf & "                    <ns9:name>입력값 확인</ns9:name>"
sSrc = sSrc & vbCrLf & "                    <ns9:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns9:code>          logger.info(""  ☆★☆★☆★☆★inSvo   = \n[""+inSvo.toString()+""] ☆★☆★☆★☆★"" );&#xD;"
sSrc = sSrc & vbCrLf & "            inBvo.setRecvHdrVo(inSvo.getRecvHdrVo());&#xD;"
sSrc = sSrc & vbCrLf & "            inBvo.setRecvDataVo(inSvo.getRecvDataVo());</ns9:code>"
sSrc = sSrc & vbCrLf & "                </ns9:virtualModule>"
sSrc = sSrc & vbCrLf & "                <ns9:virtualModule nodeId=""5"">"
sSrc = sSrc & vbCrLf & "                    <ns9:name>return</ns9:name>"
sSrc = sSrc & vbCrLf & "                    <ns9:nextEntries/>"
sSrc = sSrc & vbCrLf & "                    <ns9:code>       //outSvo.setListDvo(outBvo.getListDvo()); &#xD;"
sSrc = sSrc & vbCrLf & "         &#xD;"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "          " & sSVC메소드id & ";&#xD;"
sSrc = sSrc & vbCrLf & "         logger.info(""  ☆★☆★☆★☆★outSvo   = \n[""+outSvo.toString()+""] ☆★☆★☆★☆★"" );&#xD;"
sSrc = sSrc & vbCrLf & "         return outSvo;</ns9:code>"
sSrc = sSrc & vbCrLf & "                </ns9:virtualModule>"
sSrc = sSrc & vbCrLf & "            </ns9:virtualModules>"
sSrc = sSrc & vbCrLf & "            <ns9:loopModules/>"
sSrc = sSrc & vbCrLf & "            <ns9:bizMethodCalls>"
sSrc = sSrc & vbCrLf & "                <ns9:bizMethodCall nodeId=""7"">"
sSrc = sSrc & vbCrLf & "                    <ns9:nextEntries/>"


iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "                    <ns9:bizInstanceInfo scope=""member"" variableName=""" & sSVC메소드id & """ lazy=""false"">"
sSrc = sSrc & vbCrLf & "                        <ns9:classInfo classPackageName=""" & sSVCInVOid & """ className=""" & sSVC메소드id & """/>"
sSrc = sSrc & vbCrLf & "                    </ns9:bizInstanceInfo>"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "                    <ns9:method methodName=""" & sSVC메소드id & """ displayName=""" & sSVCInVOid & """>"
sSrc = sSrc & vbCrLf & "                        <ns12:arguments>"
sSrc = sSrc & vbCrLf & "                            <ns9:argument>"


iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "<ns9:variable kind=""DTO"" name=""inBvo"">"
sSrc = sSrc & vbCrLf & "    <ns9:do physicalName=""" & sSVC메소드id & """ resourcePath=""" & sSVCOutVOid & """/>"
sSrc = sSrc & vbCrLf & "</ns9:variable>"
sSrc = sSrc & vbCrLf & "<ns9:sequence>0</ns9:sequence>"


sSrc = sSrc & vbCrLf & "                            </ns9:argument>"
sSrc = sSrc & vbCrLf & "                        </ns12:arguments>"


iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "                        <ns12:returnType kind=""DTO"">"
sSrc = sSrc & vbCrLf & "                            <ns9:do physicalName=""" & sSVC메소드id & """ resourcePath=""" & sSVCOutVOid & """/>"
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
sSrc = sSrc & vbCrLf & "                    <ns9:do physicalName=""" & sSVC메소드id & """ resourcePath=""" & sSVCOutVOid & """/>"
sSrc = sSrc & vbCrLf & "                </ns9:variable>"


sSrc = sSrc & vbCrLf & "                <ns9:sequence>0</ns9:sequence>"
sSrc = sSrc & vbCrLf & "            </ns9:argument>"
sSrc = sSrc & vbCrLf & "        </ns13:arguments>"

iROW_CNT = iROW_CNT + 1: subSetSVCVO
sSrc = sSrc & vbCrLf & "        <ns13:returnType kind=""DTO"" name=""arg0"">"
sSrc = sSrc & vbCrLf & "            <ns9:do physicalName=""" & sSVC메소드id & """ resourcePath=""" & sSVCOutVOid & """/>"
sSrc = sSrc & vbCrLf & "        </ns13:returnType>"
sSrc = sSrc & vbCrLf & "        <ns13:classInfo classPackageName=""" & sPackageSlashName & """ className=""" & sClassPname & """/>"
sSrc = sSrc & vbCrLf & "        <ns13:annotations/>"
sSrc = sSrc & vbCrLf & "    </ns13:serviceMethod>"
sSrc = sSrc & vbCrLf & "    <ns13:interfaces/>"
sSrc = sSrc & vbCrLf & "</ns13:serviceObject>"

                                                                                                                                                                                                                              
End With
sMakeSrcStringSVC = sSrc
End Function


'사업자 VO 생성
Private Sub subMake사업자VO()


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


    'gfMsgInf ("생성되었습니다.")
    Exit Sub
Err_Handler:
    gfMsgWrn (Err.Description)
End Sub






'사업자 DAO 생성
Private Sub subMake사업자DAO()


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
        'sSrc = sSrc & "            <ns12:statement>/* hf.cor.hg.bg.dao.sdao.s523.HgbgBzepBrfHsCnstCrs001Dao.sel001 업체개요-시공사고객번호 조회 */"
        'sSrc = sSrc & "SELECT T1.JINT_BIZ_DVCD             AS JINT_BIZ_DVCD /* 공동사업자구분코드 */"
        'sSrc = sSrc & "     , T2.CNCO_CSNO                 AS CNCO_CSNO     /* 시공사고객번호 */"
        'sSrc = sSrc & "  FROM TB_HGA701D_BIZGRNT T1        /* 주택신용보증_사업자보증상세 */"
        'sSrc = sSrc & "     , TB_HGA208M_BIZCNCO T2        /* 주택신용보증_사업자시공사기본 */"
        'sSrc = sSrc & " WHERE T1.GRNT_NO = T2.GRNT_NO"
        'sSrc = sSrc & "   AND T1.GRNT_NO = :grntNo"
        'sSrc = sSrc & "   AND T2.CNCO_SEQ = 1              /* 시공사순번(AS-IS없던 컬럼.) */</ns12:statement>"
        'sSrc = sSrc & "            <ns12:parameter physicalName=""grntNo"" javaType=""String"" parameter=""true""/>"
        'sSrc = sSrc & "            <ns12:table name=""TB_HGA701D_BIZGRNT"" alias=""TB_HGA701D_BIZGRNT""/>"
        'sSrc = sSrc & "        </ns12:statements>"
        'sSrc = sSrc & "        <ns12:statements sqlType=""SELECT"" alias=""sel002"" resultType=""HgbgBzepBrfHsCnstCrs001Dvo"" resultTypeResourcePath=""hf.cor.hg.bg.vo.dvo.sdvo.s523"">"
        'sSrc = sSrc & "            <ns12:statement>/* hf.cor.hg.bg.dao.sdao.s523.HgbgBzepBrfHsCnstCrs001Dao.sel002 업체개요-변경수신 조회 */"
        'sSrc = sSrc & "SELECT T1.RSCH_BASIS_DY                 AS RSCH_BASIS_DY                /* 조사기준일자 */"
        'sSrc = sSrc & "     , T1.LSTR_ACST_BASIS_TRED_DY       AS LSTR_ACST_BASIS_TRED_DY      /* 전기결산기준기말일자 */"
        'sSrc = sSrc & "     , T1.CRTM_ACST_BASIS_TRED_DY       AS CRTM_ACST_BASIS_TRED_DY      /* 당기결산기준기말일자 */"
        'sSrc = sSrc & "     , T2.ETNO                          AS ETNO                                              /* 기업체번호 */"
        'sSrc = sSrc & "  FROM TB_HGJ003M_CUSTRVEW T1               /* 주택신용보증_고객심사기본 */"
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


    'gfMsgInf ("생성되었습니다.")
    Exit Sub
Err_Handler:
    gfMsgWrn (Err.Description)
End Sub





