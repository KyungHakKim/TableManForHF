VERSION 5.00
Begin VB.Form frmJSP 
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   5775
   ClientTop       =   5040
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   10620
   Begin VB.CommandButton cmdParsing 
      Caption         =   "GO"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "frmJSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

End Sub

Private Sub cmdParsing_Click()
    
    Dim sFielName As String
    Dim sLine() As String
    
    Dim iCnt As Long

    sFielName = "D:\000.����������\2000.�����۾�\2200.�ҽ��м�\2210.JSP �м�\����FILE.txt"
    Open sFielName For Input As #1
    
    iCnt = 0
    
    Dim bBody As Boolean
    Dim sFileLine As String
    bBody = False
    While Not EOF(1)
        
        
        
        Line Input #1, sFileLine
        
        
        'body �������� ����
        
        If InStr(UCase(sFileLine), "<BODY") > 0 Then
            bBody = True
        End If
        
        If bBody Then
        
            iCnt = iCnt + 1
            ReDim Preserve sLine(iCnt)
            sLine(iCnt) = sFileLine
    
                
            sLine(iCnt) = Replace(sLine(iCnt), "< ", "<")
            sLine(iCnt) = Replace(sLine(iCnt), " >", ">")
            sLine(iCnt) = Replace(sLine(iCnt), "  ", " ")
        End If
    
    Wend

    Close #1
    
    Dim i As Integer
    Dim sFindStr() As String
    Dim iFindCnt As Integer
    
   
    Dim bFind As Boolean
    iFindCnt = 0
    
    'select ã��
    For i = 1 To iCnt
        
        If InStr(1, UCase(sLine(i)), "<SELECT") > 0 Then
            
            iFindCnt = iFindCnt + 1
            ReDim Preserve sFindStr(iFindCnt)
            
            bFind = True
        End If
    
        If bFind = True Then
        sFindStr(iFindCnt) = sFindStr(iFindCnt) & sLine(i)
        End If
        If InStr(1, UCase(sLine(i)), "/SELECT") > 0 Then
            bFind = False
        End If
    
    
    Next i
    
    Dim sSelect As String
    Dim sOption() As String
    Dim j As Integer
    Dim sOutput As String
    'select �м�
    For i = 1 To iFindCnt
            
        '�ΰ��̻� ���� �ϳ��� ����
        sFindStr(i) = Replace(sFindStr(i), "  ", " ")
        
        'SELECT TAG FIND
        sSelect = Left(sFindStr(i), InStr(1, UCase(sFindStr(i)), "<OPTION") - 1)
    
        'SELECT TAG ����
        sFindStr(i) = Mid(sFindStr(i), InStr(1, UCase(sFindStr(i)), "<OPTION"), Len(sFindStr(i)))
        
        'OPTION TAG�� SPLIT
        sOption = Split(sFindStr(i), "</option> <option")
    
    
        '�׸� ���
        Debug.Print Trim(sSelect)
        For j = 0 To UBound(sOption)
            sOutput = UCase(sOption(j))
            
            
            sOutput = Replace(sOutput, "<OPTION", "")
            sOutput = Replace(sOutput, "</OPTION>", "")
            sOutput = Replace(sOutput, "</SELECT>", "")
            sOutput = Replace(sOutput, "VALUE", "")
            sOutput = Replace(sOutput, "=", "")
            sOutput = Replace(sOutput, ">", "-")
            
            Debug.Print sOutput
        Next j
    
    Next i
    
    
    
    
    '----------------------------------------------------------------------------
    
    
    'input ã��
    Dim iFindInputCnt As Integer
    Dim iFindInputStr() As String
    bFind = False
    
    For i = 1 To iCnt
        
        If InStr(1, UCase(sLine(i)), "RADIO") > 0 Then
            
            iFindInputCnt = iFindInputCnt + 1
            ReDim Preserve iFindInputStr(iFindInputCnt)
            
            bFind = True
        End If
    
        If bFind = True Then
            iFindInputStr(iFindInputCnt) = iFindInputStr(iFindInputCnt) & sLine(i)
        End If
        
        

    
    
    Next i
    
    'INPUT �м�
    For i = 1 To iFindCnt
            
        '�ΰ��̻� ���� �ϳ��� ����
        sFindStr(i) = Replace(sFindStr(i), "  ", " ")
        
        'SELECT TAG FIND
        sSelect = Left(sFindStr(i), InStr(1, UCase(sFindStr(i)), "<OPTION") - 1)
    
        'SELECT TAG ����
        sFindStr(i) = Mid(sFindStr(i), InStr(1, UCase(sFindStr(i)), "<OPTION"), Len(sFindStr(i)))
        
        'OPTION TAG�� SPLIT
        sOption = Split(sFindStr(i), "</option> <option")
    
    
        '�׸� ���
        Debug.Print Trim(sSelect)
        For j = 0 To UBound(sOption)
            sOutput = UCase(sOption(j))
            
            
            sOutput = Replace(sOutput, "<OPTION", "")
            sOutput = Replace(sOutput, "</OPTION>", "")
            sOutput = Replace(sOutput, "</SELECT>", "")
            sOutput = Replace(sOutput, "VALUE", "")
            sOutput = Replace(sOutput, "=", "")
            sOutput = Replace(sOutput, ">", "-")
            
            Debug.Print sOutput
        Next j
    
    Next i
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
End Sub





