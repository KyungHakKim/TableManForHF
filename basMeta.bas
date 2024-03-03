Attribute VB_Name = "basMeta"
Option Explicit

Public gsHangulWord() As String
Public gsEnglisWord() As String


Public gsDomainName() As String '도메인명
Public gsDomianType() As String '도메인유형
Public gsDomainLength() As String '도메인길이
Public gsDomainPrec() As String '도메인소수점
Public gsDomainGroup() As String    '도메인그룹명
Public gsDomainGuard() As String '도메인정보보호여부
Public gsDomainEnc() As String '도메인암호화여부



Public gsDongEum1() As String       '동음이의 한글명
Public gsDongEum2() As String       '동음이의 영문명
Public gsDongEum3() As String       '동음이의 설명

Public Sub subSetMetaWord()
    'gsHangulWord = Split("개인;주택;금융;공사;대출;금액;주택금융공사", ";")
    'gsEnglisWord = Split("PRSN;HOUSE;FIN;GONG;LOAN;AMT;HOUGONG", ";")
    
    'Call FileRead("단어장.txt", 1)
    Call subSetMetaDB(1)
    
    Call FileRead("개인단어장.txt", 2)
    
    Call FileRead("도메인.txt", 3)
    
    Call FileRead("동음이의어.txt", 4)
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
'단어장 파일 읽기
'------------------------------------------------------------
Private Sub FileRead(pFileName As String, iCallCount As Integer)
On Error GoTo ErrHandler
    Dim ffile As Integer
    Dim sTemp  As String
    Dim sTemp2() As String
    Dim iCnt As Integer
    
    '도메인 읽기면 첫 배열
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

    Open App.Path & "\" & pFileName For Input As #ffile   ' 파일을 엽니다.
    
    Do
    
        Input #ffile, sTemp
        
        sTemp2 = Split(sTemp, ";")
        
        iCnt = iCnt + 1
        
        If iCallCount = 3 Then
        
            ReDim Preserve gsDomainName(iCnt)
            ReDim Preserve gsDomianType(iCnt)
        
            ReDim Preserve gsDomainLength(iCnt) '도메인길이
            ReDim Preserve gsDomainPrec(iCnt)   '도메인소수점
            ReDim Preserve gsDomainGroup(iCnt)  '도메인그룹명
            ReDim Preserve gsDomainGuard(iCnt)  '도메인정보보호여부
            ReDim Preserve gsDomainEnc(iCnt)    '도메인암호화여부
        
        
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

'영문단어 매핑
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
                    '모두 대문자로
                    
                    sEngWord = UCase(gsEnglisWord(j))
                
                ElseIf pUpperLower = "L" Then
                    '모두 소문자로
                    sEngWord = LCase(gsEnglisWord(j))
                ElseIf pUpperLower = "F" Then
                    '첫글자를 대문자로
                    sEngWord = gfUpperFirst(gsEnglisWord(j))
                Else
                    '첫글자를 소문자 나머지 첫글자대문자
                    sEngWord = gfUpperFirst(gsEnglisWord(j))
                    
                End If
                
                
                '언더 바 붙이기
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


'용어를 단어 단위로 분리
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
                
            '배열에서 일치하는 word 있나 serarch
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
    
    
    '최종 분리된 용어 return
    Dim sResult As String
    
    
    For i = 100 To 1 Step -1
        
        If sSplitArr(i) <> "" Then
                
            '동음이의어검색해서 동음이의어 인경우 단어(동음이의어)로 표기
            For j = LBound(gsDongEum1) To UBound(gsDongEum1)
                
                If Trim(gsDongEum1(j)) = Trim(sSplitArr(i)) Then
                    sSplitArr(i) = sSplitArr(i) & "(동음이의어)"
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

'첫글자만 대문자로
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
'파일을 읽어서 배열에 저장하여 리턴한다.
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



    Open pFileName For Input As #ffile   ' 파일을 엽니다.
    
    iCnt = 0
    
    Do
    
        Input #ffile, sTemp
        iCnt = iCnt + 1
    Loop Until EOF(1)
    
    Close #ffile


    

    Open pFileName For Input As #ffile   ' 파일을 엽니다.
    
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
        gfMsgInf pFileName & " 파일을 찾을 수 없습니다."
    ElseIf Err.Number = 62 Then
        gfMsgInf pFileName & " 파일에 사용자 정보가 비었습니다. 사번;아이디 값을 파일에 등록하세요. ."
    Else
        gfMsgInf (Err.Description & " " & Err.Number)
    End If

    gfReadFile = False


End Function

'-----------------------------------------------------
'XML 구문에서 특정필드의 value를 추출한다.
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
    
    '레지스트리에서 DB연결 정보를 가져온다.
    'strProvider =
    'strDSN = GetSetting("gmcfis", "serversection", "dsn", "")
    'strUser = GetSetting("gmcfis", "serversection", "dbid", "")
    'strPass = GetSetting("gmcfis", "serversection", "dbpass", "")
    
    'Connection String 생성
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
    
    '레지스트리에서 DB연결 정보를 가져온다.
    'strProvider =
    'strDSN = GetSetting("gmcfis", "serversection", "dsn", "")
    'strUser = GetSetting("gmcfis", "serversection", "dbid", "")
    'strPass = GetSetting("gmcfis", "serversection", "dbpass", "")
    
    'Connection String 생성
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
    
    '레지스트리에서 DB연결 정보를 가져온다.
    'strProvider =
    'strDSN = GetSetting("gmcfis", "serversection", "dsn", "")
    'strUser = GetSetting("gmcfis", "serversection", "dbid", "")
    'strPass = GetSetting("gmcfis", "serversection", "dbpass", "")
    
    'Connection String 생성
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
    
    '레지스트리에서 DB연결 정보를 가져온다.
    'strProvider =
    'strDSN = GetSetting("gmcfis", "serversection", "dsn", "")
    'strUser = GetSetting("gmcfis", "serversection", "dbid", "")
    'strPass = GetSetting("gmcfis", "serversection", "dbpass", "")
    
    'Connection String 생성
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





