Attribute VB_Name = "basCommon"
Option Explicit
Public Const C_SPECIAL = "N"
Public rscode As ADODB.Recordset
Public Const C_DADA_VIEW_BASIC_COL_SIZE = 5

Public gUserId As String
Public gUserNm As String
Public gJuSinBo As Boolean
'-------------------------------
'ORACLE
'-------------------------------
'drop table TB_DIC_TAB
';
'CREATE TABLE TB_DIC_TAB
'( OWNER      VARCHAR2(20)        NOT NULL
' ,TABLE_NAME VARCHAR2(20)        NOT NULL
' ,TABLE_COMM VARCHAR2(100)
' ,BIGO1      VARCHAR2(1000)
' ,BIGO2      VARCHAR2(1000)
')
';
'ALTER TABLE TB_DIC_TAB ADD PRIMARY KEY (OWNER, TABLE_NAME)
';
'
'drop table TB_DIC_COL
';
'CREATE TABLE TB_DIC_COL
'( OWNER       VARCHAR2(20)         NOT NULL
' ,TABLE_NAME  VARCHAR2(20)         NOT NULL
' ,COLUMN_NAME VARCHAR2(30)         NOT NULL
' ,COLUMN_COMM VARCHAR2(100)
' ,CODE_DOM    VARCHAR2(10)
' ,BIGO1       VARCHAR2(2000)
' ,BIGO2       VARCHAR2(2000)
' )
';
'ALTER TABLE TB_DIC_COL ADD PRIMARY KEY (OWNER, TABLE_NAME, COLUMN_NAME)
';
'





' CREATE TABLE DIC_TAB_DESC (
'     OWNER   CHAR(20) NOT NULL,
'     TABID   CHAR(20) NOT NULL,
'     ORGTAB  CHAR(50) ,
'     CYCLE   VARCHAR(100),
'     USE_YN  CHAR(50),
'     IMP_DESC VARCHAR(5000),
'     CR_LOG VarChar(5000)
' );
' ALTER TABLE DIC_TAB_DESC ADD CONSTRAINT PK_DIC_TAB_DESC PRIMARY KEY(OWNER, TABID);
'

' CREATE TABLE DIC_TAB_IN (
'     OWNER   CHAR(20) NOT NULL,
'     TABID   CHAR(20) NOT NULL,
'     INTAB  Char(50) NOT NULL
' );
' ALTER TABLE DIC_TAB_IN ADD CONSTRAINT PK_DIC_TAB_IN PRIMARY KEY(OWNER, TABID, INTAB);


'    CREATE TABLE DIC_TAB (
'        OWNER   CHAR(20) NOT NULL,
'        TABID   CHAR(20) NOT NULL,
'        TABNM   CHAR(100) ,
'        BIGO1   VARCHAR(1500),
'        BIGO2 VarChar(1500)
'    )
'    ALTER TABLE DIC_TAB  ADD CONSTRAINT PK_DIC_TAB  PRIMARY KEY(OWNER, TABID)
'
'    CREATE TABLE DIC_COL (
'        OWNER   CHAR(20) NOT NULL,
'        TABID   CHAR(20) NOT NULL,
'        COLID   CHAR(30) NOT NULL,
'        COLNM   CHAR(50) ,
'        CDDOM   CHAR(10),
'        BIGO1   VARCHAR(1500),
'        BIGO2 VarChar(1500)
'    )
'    ALTER TABLE DIC_COL  ADD CONSTRAINT PK_DIC_COL  PRIMARY KEY(OWNER, TABID, COLID)




Public strDSN, strUser, strPass, strProvider, strDBA, strDBAUPass, strReal As String
Public strFTPAddress, strFTP_ID, strFTP_PASS As String
Public gIPAddress As String

Public gColId() As String
Public gColName() As String
Public gType() As String
Public gPK() As String

Public Function GetConnString() As String
    'Dim strDSN, strUser, strPass, strProvider As String
    
    '������Ʈ������ DB���� ������ �����´�.
    'strProvider =
    'strDSN = GetSetting("gmcfis", "serversection", "dsn", "")
    'strUser = GetSetting("gmcfis", "serversection", "dbid", "")
    'strPass = GetSetting("gmcfis", "serversection", "dbpass", "")
    
    'Connection String ����
'    GetConnString = "Provider=MSDAORA.1;" & _
                    "Data Source=" & strDSN & ";" & _
                    "User ID=" & strUser & ";" & _
                    "Password=" & strPass & ";"
    
    'strProvider = "IBMDADB2.DB2COPY1"
    'strDSN = "RSMDB"
    'strUser = "rsmuser"
    'strPass = "rsmuser1"
    
    GetConnString = "Provider=" & strProvider & ";" & _
                    "Persist Security Info=True;" & _
                    "User ID=" & strUser & ";" & _
                    "Password=" & strPass & ";" & _
                    "Data Source=" & strDSN & ";"
    
    ' GetConnString = "Driver={Microsoft ODBC for Oracle}; Server=STG2; Uid=USHG_DML; Pwd=dlatl00!"
    
End Function


Public Function GetConnStringDBA() As String
    'Dim strDSN, strUser, strPass, strProvider As String
    
    '������Ʈ������ DB���� ������ �����´�.
    'strProvider =
    'strDSN = GetSetting("gmcfis", "serversection", "dsn", "")
    'strUser = GetSetting("gmcfis", "serversection", "dbid", "")
    'strPass = GetSetting("gmcfis", "serversection", "dbpass", "")
    
    'Connection String ����
'    GetConnString = "Provider=MSDAORA.1;" & _
                    "Data Source=" & strDSN & ";" & _
                    "User ID=" & strUser & ";" & _
                    "Password=" & strPass & ";"
    
    'strProvider = "IBMDADB2.DB2COPY1"
    'strDSN = "RSM_DB"
    'strUser = strDBA
    'strPass = strDBAUPass
    
    GetConnStringDBA = "Provider=" & strProvider & ";" & _
                    "Persist Security Info=True;" & _
                    "User ID=" & strDBA & ";" & _
                    "Password=" & strDBAUPass & ";" & _
                    "Data Source=" & strDSN & ";"
    
     
    
End Function




Public Function fConnTest() As ADODB.Recordset

    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    

On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    
    Set conn = New ADODB.Connection
    conn.ConnectionString = GetConnString

    conn.Open

    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient


    strSQL = " SELECT * FROM SYSIBM.SYSTABLES WHERE ROWNUM < 5 "

    rsTemp.Open strSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText

    'rsTemp.Open strSQL, conn, adOpenStatic, adLockOptimistic, adCmdText


    Set fConnTest = rsTemp
    
    
    Screen.MousePointer = vbDefault
    gfMsg ("Database ���� �׽�Ʈ �Ϸ�")

    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    Exit Function
ErrorHandler:
    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    Screen.MousePointer = vbDefault
    
    Err.Raise Number:=Err.Number, _
              Source:="GT_SRV.clsGT001.fQuery", _
              Description:=Err.Description
'

End Function



Public Sub gfMsg(ByVal sMsg As String, Optional ByVal sTitle As String = "")
    If sTitle = "" Then
        MsgBox sMsg, vbInformation, "�˸�"
    Else
        MsgBox sMsg, vbInformation, sTitle
    End If
End Sub





'*****************************************************************************
'���α׷��� : gfRsToFp
'��      �� : ���ڵ���� ������ �������忡 ���÷��� �Ѵ�.
'��      �� : ByVal fpTemp As Object(�ʼ�)  ��������
'             ByVal rsTemp As ADODB.Recordset(�ʼ�)       ���ڵ��
'��  ��  �� : None
'��  ��  �� : Call gfRsToFp(fpGrid, rs1)
'��  ��  �� : �����
'��  ��  �� : 2001.07.25
'����  �̷� :
'*****************************************************************************
Public Sub gfRsToFp(ByVal fpTemp As Object, ByVal rsTemp As ADODB.Recordset, _
                                    Optional ByVal fpSum As Object = Nothing, _
                                    Optional ByVal trimYn As String = "Y")
    Dim i As Integer
    Dim j As Integer
    Dim K As Integer
    Dim sum As Double
    Dim sum_col As String
    Dim chk As Boolean
    
    '���ڵ� �Ǽ��� ������ Exit Sub
    If rsTemp.RecordCount < 1 Then Exit Sub


        '***********  sum  ****************
        If Not (fpSum Is Nothing) Then

                fpSum.MaxRows = 0
                fpSum.MaxRows = 1
        End If
        '***********  sum  ****************


    With fpTemp
        'Spread �ʱ�ȭ
        .MaxRows = rsTemp.RecordCount
          
        .Col = -1
        '.UserResize = 1
        .SelectBlockOptions = SS_SELBLOCKOPT_COLS Or SS_SELBLOCKOPT_ROWS Or SS_SELBLOCKOPT_BLOCKS Or SS_SELBLOCKOPT_ALL
        'Determine if the horz and/or the vert scroll bars are to be displayed
        .ScrollBars = SS_SCROLLBAR_BOTH
        'Determine if scroll bars are to be displayed if need
        .ScrollBarExtMode = True
        'Determine if the last row/col is to scroll to the end of the screen
        .ScrollBarMaxAlign = True
        'Set the maximum range of the scroll bars displayed in the spreadsheet
        .ScrollBarShowMax = True
        'Sets Scroll Bar Tracking
        .ScrollBarTrack = SS_SCROLLBARTRACK_OFF
    
    
     For j = 1 To .MaxCols
        If .ColsFrozen = j Then
           For i = 1 To j
               .Col = i
               .Row = -1
               'Determine the color of background, foreground and border color
               .ForeColor = RGB(0, 0, 0)
               .BackColor = RGB(240, 255, 240)
           Next
           Exit For
         End If
     Next

        rsTemp.MoveFirst
        For i = 1 To .MaxRows
            .Row = i
            For j = 1 To .MaxCols
                If rsTemp.Fields(j - 1).Name = "DENY_REASON" Then
                .Col = j
                If Not rsTemp.Fields(j - 1).Value = "" Then
                    .Col = -1
                    .BackColor = vbYellow
                    .BackColorStyle = BackColorStyleUnderGrid
                End If
                End If
                
                If rsTemp.Fields(j - 1).Name = "TERMLOSS" Then
                .Col = j
                    If rsTemp.Fields(j - 1).Value = "10" Then
                        .Col = 1
                        .BackColor = &HC0C0FF
                        
                    End If
                End If
                
                If rsTemp.Fields(j - 1).Name = "DELAY_YN" Then
                .Col = j
                    If rsTemp.Fields(j - 1).Value > 0 Then
                        .Col = 1
                        .BackColor = &HC0C0FF
                        
                    End If
                End If
                
                .Col = j
                If Not IsNull(rsTemp.Fields(j - 1).Value) Then
                    'CellType�� ���� Data Set
                    Select Case .CellType

                        Case CellTypeEdit
                            If trimYn = "N" Then
                                .Value = IIf(IsNull(rsTemp.Fields(j - 1).Value), "", (rsTemp.Fields(j - 1).Value))
                            Else
                                .Value = IIf(IsNull(rsTemp.Fields(j - 1).Value), "", Trim(rsTemp.Fields(j - 1).Value))
                            End If
                        Case CellTypeComboBox
                            .Text = IIf(IsNull(rsTemp.Fields(j - 1).Value), "", Trim(rsTemp.Fields(j - 1).Value))

                        Case CellTypeCheckBox
                            .Value = IIf(IsNull(rsTemp.Fields(j - 1).Value), "", Trim(rsTemp.Fields(j - 1).Value))
                            
                        Case CellTypePic
                            .Value = IIf(IsNull(rsTemp.Fields(j - 1).Value), "", Trim(rsTemp.Fields(j - 1).Value))
                        Case CellTypeButton
                            .TypeButtonText = IIf(IsNull(rsTemp.Fields(j - 1).Value), "", Trim(rsTemp.Fields(j - 1).Value))
                            .Text = IIf(IsNull(rsTemp.Fields(j - 1).Value), "", Trim(rsTemp.Fields(j - 1).Value))
                        Case CellTypeFloat
                            .Value = IIf(IsNull(rsTemp.Fields(j - 1).Value), 0, Trim(rsTemp.Fields(j - 1).Value))
                        Case CellTypeNumber
                            .Value = IIf(IsNull(rsTemp.Fields(j - 1).Value), 0, Trim(rsTemp.Fields(j - 1).Value))
                            
                    End Select

                End If
            Next j
 
            If Not (fpSum Is Nothing) And i = .MaxRows Then
                For K = 1 To .MaxCols
                      .Row = 1: .Col = K
                    If .CellType = CellTypeFloat Or .CellType = CellTypeInteger Then
                       'Call AddNewLineForCalc3(fpTemp, k, chk)
                       chk = True
                    End If
                Next K
                'Call AddNewLineForCalc2(fpTemp, fpSum)
               .MaxRows = .MaxRows - 1
            End If
            
            rsTemp.MoveNext
        Next i
                       
    End With

End Sub






'*************************************************************************************
'fpSpread �� ������ ��ȯ�Ѵ�.
'*************************************************************************************
Public Sub gfFpToExcel(objSpread As Object)

    Dim XL As Object
    Dim WB As Object
    Dim WS As Object
    Dim i As Integer
    Dim first As Integer, last As Integer
    Dim x As Integer
    
    On Error GoTo ErrTrap
    
    Screen.MousePointer = vbHourglass
    
    Set XL = GetObject(, "Excel.Application")
    
    XL.Visible = True
    
    Set WB = XL.Workbooks.Add()
    Set WS = WB.Worksheets.Add
    
    objSpread.Row = 0
    objSpread.Row2 = objSpread.MaxRows
    first = 1
    
    For i = 1 To objSpread.MaxCols
        objSpread.Col = i
        If objSpread.ColHidden = True Then
            last = i - 1
            objSpread.Col = first
            objSpread.Col2 = last
            Clipboard.Clear
            Clipboard.SetText objSpread.Clip, 1
            x = Int(first / 26)
            If x = 0 Then
               WS.Range(Chr(64 + first) & "1").Select
            Else
               WS.Range(Chr(64 + x) & Chr(65 + (first Mod 26)) & "1").Select
            End If
            
            WS.Paste
            first = i + 1
        End If
    
    Next
    
    
    last = i - 1
    objSpread.Col = first
    objSpread.Col2 = last
    Clipboard.Clear
    Clipboard.SetText objSpread.Clip, 1
    x = Int(first / 26)
    If x = 0 Then
       WS.Range(Chr(64 + first) & "1").Select
    Else
       WS.Range(Chr(64 + x) & Chr(65 + (first Mod 26)) & "1").Select
    End If
    
    WS.Paste

    
    Set WS = Nothing
    Set WB = Nothing
    Set XL = Nothing
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
ErrTrap:
    If Err = 432 Or Err = 429 Then
        Set XL = CreateObject("Excel.Application")
    
        Resume Next
    Else
        Screen.MousePointer = vbDefault
        
        MsgBox "���� ���� ����", vbCritical, "���"
    End If
    
End Sub




'*************************************************************************
'   ���� : �׸��� ��Ʈ
'*************************************************************************

Public Sub gfGrid_Sort(objFp As Object, Col As Long, Row As Long)


    
 Dim i As Long
 Dim blnAsc As Boolean
    
    With objFp
        If Row > 0 Or Col < 1 Then Exit Sub
        .Col = Col
        .Row = 0
        
        '���� - ��������, �Ķ� - ��������
        If .ForeColor = S_BLACK Or .ForeColor = S_BLUE Then
            blnAsc = True
            .ForeColor = S_RED
        Else
            blnAsc = False
            .ForeColor = S_BLUE
        End If
        
        .Row = -1
        .Col = -1
        .Lock = True
        .BackColor = C_NONE
       
        Screen.MousePointer = vbHourglass
        .BlockMode = True
        .Col = 1
        .Row = 1
        .Col2 = .MaxCols
        .Row2 = .MaxRows
        .SortBy = SortByRow
        .SortKey(1) = Col
        If blnAsc Then
            .SortKeyOrder(1) = SortKeyOrderAscending
        Else
            .SortKeyOrder(1) = SortKeyOrderDescending
        End If
                        
        .Action = ActionSort
        .BlockMode = False
        .Col = .ActiveCol
        
'        For i = 1 To .MaxRows
'            .Row = i
'            If .BackColor = S_LIGHT_GRAY Or .BackColor = S_YELLOW Then
'                .Action = 0
'                Exit For
'            End If
'        Next i
        .Row = 1
        .Col = Col
        .Action = ActionActiveCell
    End With
    
    DoEvents
    Screen.MousePointer = vbDefault

End Sub




Public Function gfMakeString(ByVal pStr As String, ByVal plen As Integer) As String
    Dim i As Integer
    Dim sTemp As String
    sTemp = ""
    For i = 1 To plen
        sTemp = sTemp & pStr
    Next i
    gfMakeString = sTemp
End Function




'*****************************************************************************
'���α׷��� : gfMsgInf
'��      �� : ����ڿ��� �˸��� �޼��� �ڽ��� ���δ�.
'��      �� : ByVal sMsg As String (�ʼ�)  ������ �޼���
'             Optional ByVal sTitle As String(����)  Ÿ��Ʋ
'��  ��  �� : None
'��  ��  �� : Call gfMsgInf("�Ϸ�Ǿ����ϴ�.")
'��  ��  �� : �����
'��  ��  �� : 2001.07. 24.
'����  �̷� :
'*****************************************************************************
Public Sub gfMsgInf(ByVal sMsg As String, Optional ByVal sTitle As String = "")
    If sTitle = "" Then
        MsgBox sMsg, vbInformation, "�˸�"
    Else
        MsgBox sMsg, vbInformation, sTitle
    End If
End Sub

'*****************************************************************************
'���α׷��� : gfMsgQst
'��      �� : ����ڿ��� ��, �ƴϿ� ������ ���� �޼��� �ڽ��� ���̰� ����� �����Ѵ�.
'��      �� : ByVal sMsg As String (�ʼ�)  ������ �޼���
'             Optional ByVal sTitle As String (����) Ÿ��Ʋ
'��  ��  �� : Integer (vbYes, vbNo)
'��  ��  �� : If gfMsgqst("����Ͻðڽ��ϱ�?") then
'                   Call a
'             Else
'                   Exit Sub
'             End If
'��  ��  �� : �����
'��  ��  �� : 2001.07. 24.
'����  �̷� :
'*****************************************************************************
Public Function gfMsgQst(ByVal sMsg As String, Optional ByVal sTitle As String = "") As Integer
    If sTitle = "" Then
        gfMsgQst = MsgBox(sMsg, vbYesNo + vbQuestion, "Ȯ��")
    Else
        gfMsgQst = MsgBox(sMsg, vbYesNo + vbQuestion, sTitle)
    End If
End Function

'*****************************************************************************
'���α׷��� : gfMsgWrn
'��      �� : ����ڿ��� ����� �޽��� �ڽ��� ���δ�.
'��      �� : ByVal sMsg As String (�ʼ�)  ������ �޼���
'             Optional ByVal sTitle As String(����)  Ÿ��Ʋ
'��  ��  �� : None
'��  ��  �� : Call gfMsgInf("�߸��� �ֹε�Ϲ�ȣ �Դϴ�.")
'��  ��  �� : �����
'��  ��  �� : 2001.07. 24.
'����  �̷� :
'*****************************************************************************
Public Sub gfMsgWrn(ByVal sMsg As String, Optional ByVal sTitle As String = "")
    If sTitle = "" Then
        MsgBox sMsg, vbExclamation, "���"
    Else
        MsgBox sMsg, vbExclamation, sTitle
    End If
End Sub

Public Sub gfMsgCrt(ByVal sMsg As String, Optional ByVal sTitle As String = "")
    If sTitle = "" Then
        MsgBox sMsg, vbCritical, "����"
    Else
        MsgBox sMsg, vbCritical, sTitle
    End If
End Sub






'------------------------------------------------------------
'�۷ι� SQL ó�� (��ȸ�� �������� ����)
'------------------------------------------------------------
Public Function gfGlobalSQL(pSQL As String) As ADODB.Recordset
    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    
    If GCONN_USE = False Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = GetConnString
    
        conn.Open
    End If
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient


  



    sSQL = pSQL

    
    
    
    If GCONN_USE = True Then
        rsTemp.Open sSQL, gConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Else
        rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If


    
    Screen.MousePointer = vbDefault
    
    Set gfGlobalSQL = rsTemp
    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    Exit Function
ErrorHandler:
    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg Err.Description
'


End Function



'*****************************************************************************
' fpSpread �μ�
'*****************************************************************************
Public Sub subSpreadSmart( _
    ByVal iOrientation As Integer, _
    ByVal lTMargin As Long, _
    ByVal lBMargin As Long, _
    ByRef objSpread As Object, _
    ByVal sMainHeader As String, _
    ByVal sSubHeader As String, _
    Optional ByVal sSubHeader2 As String, _
    Optional ByVal sSubHeader3 As String)
    
 '�μ����� �����Ѵ�
    Dim iPrintMethod As Integer
'    iPrintMethod = MsgBox("                    �����ͷ� �ٷ� �μ⸦ �Ͻðڽ��ϱ�?" & vbCrLf & vbCrLf & _
                          "�����ͷ� �ٷ� �μ� (��)  ,  �μ� �̸����� (�ƴϿ�)  ,  �μ⸦ ��� (���)", _
                          vbYesNoCancel + vbQuestion, "Ȯ��")
    
    Screen.MousePointer = vbHourglass
    
    '��Ҹ� �����Ͽ��� ���
    If iPrintMethod = vbCancel Then
        '����
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    '������ ���
    
    '���� ���� ����
    Dim pPrinter As Printer
    Dim lPrinterW As Long
    Dim lPrinterH As Long
    Dim lColSum As Long
    Dim lColTwip As Long
    Dim lSheetWidth As Long
    Dim siColWidth As Single
    Dim iLoop1 As Integer
    Dim sFonts1 As String
    Dim sFonts2 As String
    Dim sTitles As String
    Dim sFooter As String
    Dim a_lColWidth() As Single
    
    '������ ������ ��/���� ���� (Landscape����)
    Set pPrinter = Printers(0)
    If pPrinter.Orientation = 1 Then
        lPrinterW = pPrinter.Height - 1000
        lPrinterH = pPrinter.Width - 1000
    Else
        lPrinterW = pPrinter.Width - 1000
        lPrinterH = pPrinter.Height - 1000
    End If
    
    '��� ��Ʈ ����
    sFonts1 = "/fn""����ü""/fz""15""/fb1/fi0/fu1/fk0/fs1"
    sFonts2 = "/fn""����ü""/fz""9""/fb0/fi0/fu0/fk0/fs2"
    sTitles = sFonts1 & sFonts2
    
    '��� �⺻ ����
    objSpread.PrintOrientation = iOrientation
    objSpread.PrintMarginTop = lTMargin
    objSpread.PrintMarginBottom = lBMargin
    objSpread.PrintMarginLeft = 0
    objSpread.PrintMarginRight = 0
    objSpread.PrintColor = True
    
    
    objSpread.PrintColHeaders = True
    objSpread.PrintRowHeaders = True
    objSpread.PrintType = PrintTypeAll
    objSpread.PrintGrid = True
    objSpread.PrintShadows = True
    'objSpread.PrintUseDataMax = True
    
    '���������Ʈ�� ��ü �÷� �� Ȯ��
    ReDim a_lColWidth(0 To objSpread.MaxCols)
    lColSum = 0
    For iLoop1 = 0 To objSpread.MaxCols
        '���� �߰� �������� �÷��� ����Ÿ�� ���� ��� �μⰡ �ȵǹǷ� �����̽� ���� ����
        objSpread.Row = 1
        objSpread.Col = iLoop1
        If Trim(objSpread.Text) = "" And iLoop1 <> 0 Then
'test ���⼭ ���� type���� �ǵ�����
            objSpread.CellType = CellTypeEdit
            objSpread.Text = " "
        End If
    
        a_lColWidth(iLoop1) = objSpread.ColWidth(iLoop1)
        If objSpread.ColWidth(iLoop1) < objSpread.MaxTextColWidth(iLoop1) Then
            objSpread.ColWidth(iLoop1) = objSpread.MaxTextColWidth(iLoop1)
        End If
        siColWidth = objSpread.ColWidth(iLoop1)
        objSpread.ColWidthToTwips siColWidth, lColTwip
        lColSum = lColSum + lColTwip
    Next iLoop1
    
    '���������Ʈ���� BackColor ���� (�������)
    'objSpread.BlockMode = True
    'objSpread.Row = 1
    'objSpread.Row2 = objSpread.MaxRows
    'objSpread.Col = 0
    'objSpread.Col2 = objSpread.MaxCols
    'objSpread.BackColor = RGB(255, 255, 255)
    'objSpread.BlockMode = False
        
    '���������Ʈ �� �缳��
    'If lColSum < objSpread.Width Then
    '    lSheetWidth = objSpread.Width
    '    objSpread.Width = lColSum
    'Else
    '    lSheetWidth = 0
    'End If
      
    '��� �¿츶�� ������
    If objSpread.PrintOrientation = PrintOrientationPortrait Then
        If lColSum < lPrinterH Then
            objSpread.PrintMarginLeft = (lPrinterH - lColSum) / 2
            objSpread.PrintMarginRight = (lPrinterH - lColSum) / 2
            objSpread.PrintSmartPrint = False
        Else
            objSpread.PrintOrientation = PrintOrientationLandscape
        End If
    End If
    If objSpread.PrintOrientation = PrintOrientationLandscape Then
        If lColSum < lPrinterW Then
            objSpread.PrintMarginLeft = (lPrinterW - lColSum) / 2
            objSpread.PrintMarginRight = (lPrinterW - lColSum) / 2
            objSpread.PrintSmartPrint = False
        Else
            objSpread.PrintMarginLeft = 500
            objSpread.PrintMarginRight = 500
            objSpread.PrintSmartPrint = True
        End If
    End If
    
    '��� �Ӹ���/�ٴڱ� ����
    sTitles = sTitles & "/n"
    sTitles = sTitles & "/n"
    
    sTitles = sTitles & "/f1" & "/c" & sMainHeader
    sTitles = sTitles & "/n"
    sTitles = sTitles & "/n"
    
    sTitles = sTitles & "/n" & "/f2" & "/l" & sSubHeader
    If sSubHeader2 <> "" Then
       sTitles = sTitles & "/n" & "/f2" & "/l" & sSubHeader2
    End If
    If sSubHeader3 <> "" Then
       sTitles = sTitles & "/n" & "/f2" & "/l" & sSubHeader3
    End If
    
    objSpread.PrintHeader = sTitles
    sFooter = "/n"
    'sFooter = sFooter & "/l" & Trim(gUser_nm) & "(" & Trim(gUser_no) & ")"
    sFooter = sFooter & "/c" & "/p" & "/" & objSpread.PrintPageCount
    sFooter = sFooter & "/r" & Date & " " & Time
    objSpread.PrintFooter = sFooter
    
    '�̸����⸦ �����Ͽ��� ���
    'If iPrintMethod = vbNo Then
   '
   '     Screen.MousePointer = vbDefault
   '
   '     frmSmartPreview.subSpreadSmart objSpread
   '     Screen.MousePointer = vbDefault
   '     '���������Ʈ�� ��ü �÷� �� ����ġ
   '     For iLoop1 = 0 To objSpread.MaxCols
   '         objSpread.ColWidth(iLoop1) = a_lColWidth(iLoop1)
   '     Next iLoop1
   '     '����
   '     Screen.MousePointer = vbDefault
   '     Exit Sub
   ' End If
     
    '��� ó��
    objSpread.PrintAbortMsg = "�μ����Դϴ�. �μ⸦ �����Ϸ��� Cancel��ư�� ��������."
    objSpread.Action = ActionSmartPrint
    
    '���������Ʈ�� ��ü �÷� �� ����ġ
    For iLoop1 = 0 To objSpread.MaxCols
        objSpread.ColWidth(iLoop1) = a_lColWidth(iLoop1)
    Next iLoop1
    '���������Ʈ �� ������ġ
    'If lSheetWidth <> 0 Then
    '    objSpread.Width = lSheetWidth
    'End If
    
    Screen.MousePointer = vbDefault
End Sub





Public Function nLeft(ByVal strTmp As String)
    Dim bytesize As Integer
    Dim i As Integer
    bytesize = 0
    
    ' ���ڿ� ���� ��ŭ �ݺ��ϸ鼭 ���ڱ��̸�ŭ �߶����
    For i = 1 To Len(strTmp)
    
            ' �ѱ۰� ������ �����ϱ�
            If Asc(Mid(strTmp, i, 1)) > 0 Then      '�ѱ۰��� 0���� �۴�
                bytesize = bytesize + 1 '�ѱ��� �ƴѰ�� �� ������ 1Byte
            Else
                bytesize = bytesize + 2 '�ѱ��� ��� 2Byte
            End If

    Next
    
    nLeft = bytesize
    
End Function







Public Function gfMid(sInStr As String, iStart As Integer, iCnt As Integer) As String

    gfMid = StrConv(MidB(StrConv(sInStr, vbFromUnicode), iStart, iCnt), vbUnicode)

End Function



