VERSION 5.00
Begin VB.Form frmVlookup 
   Appearance      =   0  '평면
   BackColor       =   &H80000005&
   BorderStyle     =   1  '단일 고정
   Caption         =   "요구사항 Vlookup"
   ClientHeight    =   10305
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17580
   Icon            =   "frmVlookup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10305
   ScaleWidth      =   17580
   Begin VB.TextBox txtTgt 
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
      Left            =   2520
      TabIndex        =   8
      Top             =   1260
      Width           =   2115
   End
   Begin VB.TextBox txtOrg 
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
      Left            =   240
      TabIndex        =   7
      Top             =   1260
      Width           =   2115
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Height          =   615
      Left            =   4920
      TabIndex        =   6
      Top             =   1140
      Width           =   1995
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   915
      Left            =   13620
      TabIndex        =   5
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox txtRightResult 
      Height          =   8175
      Left            =   13260
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   4
      Top             =   1860
      Width           =   3735
   End
   Begin VB.TextBox txtLeftResult 
      Height          =   8175
      Left            =   9360
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   3
      Top             =   1860
      Width           =   3735
   End
   Begin VB.CommandButton cmdStep1 
      Caption         =   "☞"
      Height          =   2175
      Left            =   8220
      TabIndex        =   2
      Top             =   3660
      Width           =   735
   End
   Begin VB.TextBox txtRight 
      Height          =   8175
      Left            =   4200
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   1
      Top             =   1920
      Width           =   3735
   End
   Begin VB.TextBox txtLeft 
      Height          =   8175
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   0
      Top             =   1920
      Width           =   3735
   End
End
Attribute VB_Name = "frmVlookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    txtLeft.Text = ""
    txtRight.Text = ""
    txtLeftResult.Text = ""
    txtRightResult.Text = ""
End Sub

Private Sub cmdReplace_Click()
    txtLeft.Text = Replace(txtLeft.Text, Trim(txtOrg.Text), Trim(txtTgt.Text))
    txtRight.Text = Replace(txtRight.Text, Trim(txtOrg.Text), Trim(txtTgt.Text))
    
End Sub

Private Sub cmdStep1_Click()
'On Error GoTo Err_Handler
    Dim sLeft() As String
    Dim sRight() As String
    Dim i, j As Integer
    
    sLeft = Split(txtLeft.Text, vbCrLf)
    sRight = Split(txtRight.Text, vbCrLf)
    
    
    txtLeftResult.Text = ""
    txtRightResult.Text = ""
    
    
    Dim bFind As Boolean
    bFind = False
    For i = LBound(sLeft) To UBound(sLeft)
        For j = LBound(sRight) To UBound(sRight)
            
            If Trim(sLeft(i)) = Trim(sRight(j)) Then
                bFind = True
                Exit For
            End If
            
        Next j
        
        If Not bFind Then
            txtLeftResult.Text = txtLeftResult.Text & sLeft(i) & vbCrLf
            
        End If
        
        bFind = False
    Next i
    

    bFind = False
    For i = LBound(sRight) To UBound(sRight)
        For j = LBound(sLeft) To UBound(sLeft)
            
            If Trim(sRight(i)) = Trim(sLeft(j)) Then
                bFind = True
                Exit For
            End If
            
        Next j
        
        If Not bFind Then
            txtRightResult.Text = txtRightResult.Text & sRight(i) & vbCrLf
            
        End If
        
        bFind = False
    Next i
    

    '중복제거
    If Trim(txtLeftResult.Text) <> "" Then
        sLeft = Split(txtLeftResult.Text, vbCrLf)
        
    
        Dim sResult() As String
        Dim iCnt As Integer
        
        ReDim Preserve sResult(0)
        iCnt = 0
        sResult(0) = sLeft(0)
        For i = LBound(sLeft) To UBound(sLeft)
            
            bFind = False
            For j = LBound(sResult) To UBound(sResult)
                If Trim(sLeft(i)) = Trim(sResult(j)) Then
                    bFind = True
                    Exit For
                End If
            Next j
            
            If Not bFind Then
                iCnt = iCnt + 1
                ReDim Preserve sResult(iCnt)
                sResult(iCnt) = sLeft(i)
            End If
        Next i
        
        txtLeftResult.Text = ""
        
        For i = LBound(sResult) To UBound(sResult)
            
            txtLeftResult.Text = txtLeftResult.Text & sResult(i) & vbCrLf
        
        Next i
            
    End If
    
    
    If Trim(txtRightResult.Text) <> "" Then
    
        sRight = Split(txtRightResult.Text, vbCrLf)
        
        Dim sResult2() As String
        
        ReDim Preserve sResult2(0)
        iCnt = 0
        sResult2(0) = sRight(0)
        For i = LBound(sRight) To UBound(sRight)
            
            bFind = False
            For j = LBound(sResult2) To UBound(sResult2)
                If Trim(sRight(i)) = Trim(sResult2(j)) Then
                    bFind = True
                    Exit For
                End If
            Next j
            
            If Not bFind Then
                iCnt = iCnt + 1
                ReDim Preserve sResult2(iCnt)
                sResult2(iCnt) = sRight(i)
            End If
        Next i
        
        txtRightResult.Text = ""
        
        For i = LBound(sResult2) To UBound(sResult2)
            
            txtRightResult.Text = txtRightResult.Text & sResult2(i) & vbCrLf
        
        Next i
        
    End If
    
    

    Exit Sub
Err_Handler:
    gfMsgWrn (Err.Description)
    
    
    
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    txtOrg.Text = """"
End Sub
