Attribute VB_Name = "basMeta"
Option Explicit

Public gsHangulWord() As String
Public gsEnglisWord() As String


Public gsDomainName() As String '�����θ�
Public gsDomianType() As String '����������
Public gsDomainLength() As String '�����α���
Public gsDomainPrec() As String '�����μҼ���
Public gsDomainGroup() As String    '�����α׷��
Public gsDomainGuard() As String '������������ȣ����
Public gsDomainEnc() As String '�����ξ�ȣȭ����



Public gsDongEum1() As String       '�������� �ѱ۸�
Public gsDongEum2() As String       '�������� ������
Public gsDongEum3() As String       '�������� ����

Public Sub subSetMetaWord()
    'gsHangulWord = Split("����;����;����;����;����;�ݾ�;���ñ�������", ";")
    'gsEnglisWord = Split("PRSN;HOUSE;FIN;GONG;LOAN;AMT;HOUGONG", ";")
    
    'Call FileRead("�ܾ���.txt", 1)
    Call subSetMetaDB(1)
    
    Call FileRead("���δܾ���.txt", 2)
    
    Call FileRead("������.txt", 3)
    
    Call FileRead("�������Ǿ�.txt", 4)
End Sub


Public Sub subSetMetaDB(iCallCount As Integer)
    Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    

    Set conn = New ADODB.Connection
    conn.ConnectionString = GetConnStringHgMETA

    conn.Open

    
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
        sSQL = ""
        sSQL = sSQL & vbCrLf & "SELECT DIC_LOG_NM, DIC_PHY_NM"
        sSQL = sSQL & vbCrLf & "  FROM VW_MET_WORD           "
        sSQL = sSQL & vbCrLf & " WHERE 1=1"
    rsTemp.Open sSQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    
    Dim i As Integer
    Dim iCnt As Integer
    
    rsTemp.MoveFirst
    iCnt = 0
    For i = 1 To rsTemp.RecordCount
        
        iCnt = iCnt + 1
        
        ReDim Preserve gsHangulWord(iCnt)
        ReDim Preserve gsEnglisWord(iCnt)
        
    
        gsHangulWord(iCnt) = Trim(rsTemp.Fields("DIC_LOG_NM").Value)
        gsEnglisWord(iCnt) = Trim(rsTemp.Fields("DIC_PHY_NM").Value)
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




'------------------------------------------------------------
'�ܾ��� ���� �б�
'------------------------------------------------------------
Private Sub FileRead(pFileName As String, iCallCount As Integer)
On Error GoTo ErrHandler
    Dim ffile As Integer
    Dim sTemp  As String
    Dim sTemp2() As String
    Dim iCnt As Integer
    
    '������ �б�� ù �迭
    If iCallCount = 3 Then
        iCnt = 0
    ElseIf iCallCount = 4 Then
        iCnt = 0
    ElseIf iCallCount = 2 Then
        iCnt = UBound(gsHangulWord)
    ElseIf iCallCount = 1 Then
        iCnt = 0
    End If
    
    ffile = FreeFile

    Open App.Path & "\" & pFileName For Input As #ffile   ' ������ ���ϴ�.
    
    Do
    
        Input #ffile, sTemp
        
        sTemp2 = Split(sTemp, ";")
        
        iCnt = iCnt + 1
        
        If iCallCount = 3 Then
        
            ReDim Preserve gsDomainName(iCnt)
            ReDim Preserve gsDomianType(iCnt)
        
            ReDim Preserve gsDomainLength(iCnt) '�����α���
            ReDim Preserve gsDomainPrec(iCnt)   '�����μҼ���
            ReDim Preserve gsDomainGroup(iCnt)  '�����α׷��
            ReDim Preserve gsDomainGuard(iCnt)  '������������ȣ����
            ReDim Preserve gsDomainEnc(iCnt)    '�����ξ�ȣȭ����
        
        
            gsDomainName(iCnt) = Trim(sTemp2(0))
            gsDomianType(iCnt) = Trim(sTemp2(1))
            
            gsDomainLength(iCnt) = Trim(sTemp2(2))
            gsDomainPrec(iCnt) = Trim(sTemp2(3))
            gsDomainGroup(iCnt) = Trim(sTemp2(4))
            gsDomainGuard(iCnt) = Trim(sTemp2(5))
            gsDomainEnc(iCnt) = Trim(sTemp2(6))
            
            
            
        ElseIf iCallCount = 4 Then
            ReDim Preserve gsDongEum1(iCnt)
            ReDim Preserve gsDongEum2(iCnt)
            ReDim Preserve gsDongEum3(iCnt)
            
            gsDongEum1(iCnt) = Trim(sTemp2(0))
            gsDongEum2(iCnt) = Trim(sTemp2(1))
            gsDongEum3(iCnt) = Trim(sTemp2(2))
        
        
        Else
            ReDim Preserve gsHangulWord(iCnt)
            ReDim Preserve gsEnglisWord(iCnt)
        
        
            gsHangulWord(iCnt) = Trim(sTemp2(0))
            gsEnglisWord(iCnt) = Trim(sTemp2(1))
        End If
        
    Loop Until EOF(1)
    
    Close #ffile
    Exit Sub
ErrHandler:
    gfMsgInf (Err.Description)
    
End Sub

'�����ܾ� ����
Public Function gfMappingEng(pInWord As String, pComment As String, pUpperLower As String, pUnderBar As Boolean) As String

    Dim i, j As Integer
    Dim sOrgArr() As String
    Dim sReturn As String
    
    sOrgArr = Split(pInWord, "_")
    
    Dim sEngWord As String
    
    Dim bFind As Boolean
    For i = LBound(sOrgArr) To UBound(sOrgArr)
        
        bFind = False
        
        For j = LBound(gsEnglisWord) To UBound(gsEnglisWord)
            
            If sOrgArr(i) = gsHangulWord(j) Then
                
                
                If pUpperLower = "U" Then
                    '��� �빮�ڷ�
                    
                    sEngWord = UCase(gsEnglisWord(j))
                
                ElseIf pUpperLower = "L" Then
                    '��� �ҹ��ڷ�
                    sEngWord = LCase(gsEnglisWord(j))
                ElseIf pUpperLower = "F" Then
                    'ù���ڸ� �빮�ڷ�
                    sEngWord = gfUpperFirst(gsEnglisWord(j))
                Else
                    'ù���ڸ� �ҹ��� ������ ù���ڴ빮��
                    sEngWord = gfUpperFirst(gsEnglisWord(j))
                    
                End If
                
                
                '��� �� ���̱�
                If pUnderBar = True Then
                    If sReturn <> "" Then
                        sReturn = sReturn & "_" & sEngWord
                    Else
                        sReturn = sReturn & sEngWord
                    End If
                    
                Else
                    sReturn = sReturn & sEngWord
                End If
                bFind = True
                Exit For
                
            End If
        
        Next j
        
        If bFind = False Then
            sReturn = sReturn & sOrgArr(i)
        End If
    
    Next i

        
    gfMappingEng = sReturn

End Function


'�� �ܾ� ������ �и�
Public Function gfSplitWord(pInWord As String) As String

    Dim sSplitArr(100) As String
    Dim i, j As Integer
    Dim iCnt As Integer
    Dim sTemp As String
    Dim sOrgWord As String
    
    
    sOrgWord = Trim(pInWord)
    iCnt = 0
    
    Do
        iCnt = iCnt + 1
        
        If Trim(sOrgWord) = "" Then Exit Do
        
        For i = Len(sOrgWord) To 1 Step -1
        
            sTemp = Mid(sOrgWord, i, Len(sOrgWord))
                
            '�迭���� ��ġ�ϴ� word �ֳ� serarch
            For j = LBound(gsHangulWord) To UBound(gsHangulWord)
                
                If sTemp = Trim(gsHangulWord(j)) Then
                    sSplitArr(iCnt) = sTemp
                    Exit For
                End If
            
            Next j
        
        Next i
        
        If sSplitArr(iCnt) = "" Then
            sSplitArr(iCnt) = Right(sOrgWord, 1)
        End If
        
        
        sOrgWord = Replace(sOrgWord, sSplitArr(iCnt), "")
        
        
    Loop
    
    
    '���� �и��� ��� return
    Dim sResult As String
    
    
    For i = 100 To 1 Step -1
        
        If sSplitArr(i) <> "" Then
                
            '�������Ǿ�˻��ؼ� �������Ǿ� �ΰ�� �ܾ�(�������Ǿ�)�� ǥ��
            For j = LBound(gsDongEum1) To UBound(gsDongEum1)
                
                If Trim(gsDongEum1(j)) = Trim(sSplitArr(i)) Then
                    sSplitArr(i) = sSplitArr(i) & "(�������Ǿ�)"
                    Exit For
                End If
            
            Next j
            
            
            If sResult = "" Then
                sResult = sSplitArr(i)
            Else
                sResult = sResult & "_" & sSplitArr(i)
            End If
            
                    
        End If
    Next i
    
    gfSplitWord = sResult
    
End Function

'ù���ڸ� �빮�ڷ�
Public Function gfUpperFirst(sInStr) As String

    Dim i As Integer
    Dim sReturn As String
    
    For i = 1 To Len(sInStr)
        
        If i = 1 Then
            sReturn = sReturn & UCase(Mid(sInStr, i, 1))
        Else
            sReturn = sReturn & LCase(Mid(sInStr, i, 1))
        End If
    
    Next i
    
    gfUpperFirst = sReturn

End Function

'-----------------------------------------------
'������ �о �迭�� �����Ͽ� �����Ѵ�.
'-----------------------------------------------
Public Function gfReadFile(pFileName As String, pDelimiter As String, sResult() As String) As Boolean
On Error GoTo ErrHandler
    Dim ffile As Integer
    Dim iCnt As Integer
    Dim sTemp As String
    Dim sTemp2() As String
    Dim j As Integer
    'Dim sResult() As String
    
    ffile = FreeFile



    Open pFileName For Input As #ffile   ' ������ ���ϴ�.
    
    iCnt = 0
    
    Do
    
        Input #ffile, sTemp
        iCnt = iCnt + 1
    Loop Until EOF(1)
    
    Close #ffile


    

    Open pFileName For Input As #ffile   ' ������ ���ϴ�.
    
    Dim iTotCnt As Integer
    iTotCnt = iCnt
    
    iCnt = 0
    
    
    Do
    
        Input #ffile, sTemp
        
        sTemp2 = Split(sTemp, pDelimiter)
        
        
        
        
        ReDim Preserve sResult(0 To iTotCnt, 0 To UBound(sTemp2))
        
        For j = LBound(sTemp2) To UBound(sTemp2)
        
            sResult(iCnt, j) = sTemp2(j)
        Next j
        
        
        iCnt = iCnt + 1
    Loop Until EOF(1)
    
    Close #ffile
    
    gfReadFile = True
    
    Exit Function
ErrHandler:
    If Err.Number = 53 Then
        gfMsgInf pFileName & " ������ ã�� �� �����ϴ�."
    ElseIf Err.Number = 62 Then
        gfMsgInf pFileName & " ���Ͽ� ����� ������ ������ϴ�. ���;���̵� ���� ���Ͽ� ����ϼ���. ."
    Else
        gfMsgInf (Err.Description & " " & Err.Number)
    End If

    gfReadFile = False


End Function

'-----------------------------------------------------
'XML �������� Ư���ʵ��� value�� �����Ѵ�.
'-----------------------------------------------------
Public Function gfGetXMLValue(pOrg As String, pFind As String)

    Dim sTemp As String
    Dim iCnt As Integer
    Dim sRet As String
    Dim sChar As String
    sTemp = pOrg
    
    Dim bFirstColon As Boolean
    Dim bLastColon As Boolean
    Dim bEqual As Boolean
    
    bFirstColon = False
    bLastColon = False
    bEqual = False
    
    
    iCnt = InStr(1, UCase(pOrg), UCase(pFind))
    Do
        iCnt = iCnt + 1
        
        sChar = Mid(pOrg, iCnt, 1)
        
        If sChar = "=" Then
            bEqual = True
        End If
        
        
        If sChar = """" And bFirstColon = True Then
            Exit Do
        End If
        
        If bFirstColon = True Then
            sRet = sRet & sChar
        End If
        
        If sChar = """" And bFirstColon = False Then
            bFirstColon = True
        End If
        
    Loop
    
    gfGetXMLValue = sRet


End Function



Public Function GetConnStringHgDev() As String
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
    
    GetConnStringHgDev = "Provider=" & "OraOLEDB.Oracle.1" & ";" & _
                    "Persist Security Info=True;" & _
                    "User ID=" & "hgdev" & ";" & _
                    "Password=" & "Phgdev1234!" & ";" & _
                    "Data Source=" & "ENTDBD" & ";"
    
     
    
End Function



Public Function GetConnStringCorApp() As String

    
    GetConnStringCorApp = "Provider=" & "OraOLEDB.Oracle.1" & ";" & _
                    "Persist Security Info=True;" & _
                    "User ID=" & "COR_APP" & ";" & _
                    "Password=" & "Pcorapp1234!" & ";" & _
                    "Data Source=" & "ENTDBD" & ";"
    
     
    
End Function




'Public Function GetConnStringHgDML() As String
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
    
    'GetConnStringHgDML = "Provider=" & "OraOLEDB.Oracle.1" & ";" & _
                    "Persist Security Info=True;" & _
                    "User ID=" & "USHG_DML" & ";" & _
                    "Password=" & "dlatl00!" & ";" & _
                    "Data Source=" & "STG2" & ";"
    
     
    
'End Function


Public Function GetConnStringHgDMB() As String
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
    
    GetConnStringHgDMB = "Provider=" & "OraOLEDB.Oracle.1" & ";" & _
                    "Persist Security Info=True;" & _
                    "User ID=" & "MIGADM" & ";" & _
                    "Password=" & "Pmigadm1234!" & ";" & _
                    "Data Source=" & "ENTDBD" & ";"
    
     
    
End Function




Public Function GetConnStringHgMETA() As String
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
    
    GetConnStringHgMETA = "Provider=" & "OraOLEDB.Oracle.1" & ";" & _
                    "Persist Security Info=True;" & _
                    "User ID=" & "MET_INF" & ";" & _
                    "Password=" & "met_inf_test000" & ";" & _
                    "Data Source=" & "NINFO" & ";"
    
     
    
End Function





