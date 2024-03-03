Attribute VB_Name = "basFileModule"
Option Explicit

Public gBas_FileDir() As String
Public gBas_FileName() As String
Private gBas_FileCount As Integer

Public Sub GetMetaInfo(pSearchFolder As String, lblobjcet As Label)

  Dim folderPath As String
  Dim FileName As String
  'Dim fsFile As FileSearch
  Dim oldStatusBar As Boolean

 ReDim gBas_FileDir(0)
 ReDim gBas_FileName(0)
 
 gBas_FileCount = 0

  
  '검색할 대상 폴더 설정
  folderPath = pSearchFolder
  
  Call SearchFiles(folderPath, lblobjcet)

  
End Sub

Sub SearchFiles(strFolder, lblobjcet As Label)
  Dim i As Long
  Dim j As Integer
  
  Dim fs As Variant
  Dim fsFolder As Variant
  Dim fsFile As Variant
  
  
  Dim fileExt As String
  Dim fileExt1 As String
  Dim fileExt2 As String
  Dim fileType As String
  
'  i = Range("a1").CurrentRegion.Rows.Count + 1
  Set fs = CreateObject("Scripting.FileSystemObject")
  Set fsFolder = fs.GetFolder(strFolder)
  
  For Each fsFile In fsFolder.Files
    DoEvents
    lblobjcet.Caption = fsFile.ParentFolder & "\" & fsFile.Name
    DoEvents
    If Trim(fsFile.Name) <> "" Then
        gBas_FileCount = gBas_FileCount + 1
        ReDim Preserve gBas_FileDir(0 To gBas_FileCount)
        ReDim Preserve gBas_FileName(0 To gBas_FileCount)
        
        gBas_FileDir(gBas_FileCount) = fsFile.ParentFolder
        gBas_FileName(gBas_FileCount) = fsFile.Name
    End If
  Next
  
  For Each fsFolder In fsFolder.SubFolders
    Call SearchFiles(fsFolder, lblobjcet)
  Next

End Sub


Public Sub gfSplitNuclear(pSentence As String, sRet() As String)


    Dim iCnt As Integer
    Dim iArrCnt  As Integer
    Dim sWord As String
    
    sWord = ""
    iCnt = 0
    iArrCnt = -1
    Do
        iCnt = iCnt + 1
        If iCnt > Len(pSentence) Then Exit Do
        
        Select Case Mid(pSentence, iCnt, 1)
            Case ".", "(", """", ")", "//", ";", " ", Chr(9), "/*", "*/"
            
            
                iArrCnt = iArrCnt + 1
                ReDim Preserve sRet(0 To iArrCnt)
                
                sRet(iArrCnt) = sWord
                sWord = ""
                
                '분리자도 배열에 저장
                iArrCnt = iArrCnt + 1
                ReDim Preserve sRet(0 To iArrCnt)
                sRet(iArrCnt) = Mid(pSentence, iCnt, 1)

            Case Else
                sWord = sWord & Mid(pSentence, iCnt, 1)
        End Select

    Loop
    
    '마지막 문자는 저장되지 않음
    iArrCnt = iArrCnt + 1
    ReDim Preserve sRet(0 To iArrCnt)
                
    sRet(iArrCnt) = sWord
    


End Sub




