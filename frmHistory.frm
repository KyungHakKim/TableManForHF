VERSION 5.00
Begin VB.Form frmHistory 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   Caption         =   "Form1"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10980
   Icon            =   "frmHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdClose 
      Caption         =   "닫기"
      Height          =   435
      Left            =   7740
      Style           =   1  '그래픽
      TabIndex        =   1
      Top             =   7920
      Width           =   3195
   End
   Begin VB.TextBox txtHistory 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7755
      IMEMode         =   8  '영문
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   0
      Top             =   60
      Width           =   10815
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    subAddHistroy ("ver 1.36    2011.01.06 ")
    subAddHistroy (" ")
    subAddHistroy ("   공통코드 메타 보기 TRIM 처리 ")
    subAddHistroy (" ")
    
    subAddHistroy ("ver 1.35    2011.01.05 ")
    subAddHistroy (" ")
    subAddHistroy ("   프로시저 소스 보기 기능 추가 ")
    subAddHistroy (" ")
    subAddHistroy ("ver 1.35    2011.01.05 ")
    subAddHistroy (" ")
    subAddHistroy ("   프로시저 목록 콤보박스에 re-load 버튼 기능 추가 ")
    subAddHistroy (" ")
    subAddHistroy ("ver 1.34    2011.01.03 ")
    subAddHistroy (" ")
    subAddHistroy ("   디버그 탭에 Stored Procedure명 콤보박스로 변경 ")
    
End Sub

Private Sub subAddHistroy(sText As String)
    txtHistory.Text = txtHistory.Text + sText + vbCrLf

End Sub
