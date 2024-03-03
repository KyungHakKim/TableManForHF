VERSION 5.00
Begin VB.Form frmSrc 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '단일 고정
   Caption         =   "소스보기"
   ClientHeight    =   10395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16950
   Icon            =   "frmSrc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10395
   ScaleWidth      =   16950
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame frmSrc 
      BackColor       =   &H00FFFFFF&
      Height          =   10335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16875
      Begin VB.TextBox txtSrc 
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
         Height          =   9555
         IMEMode         =   8  '영문
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   1
         Top             =   600
         Width           =   16575
      End
      Begin VB.Image Image2 
         Height          =   405
         Left            =   16020
         Picture         =   "frmSrc.frx":030A
         Stretch         =   -1  'True
         Top             =   120
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmSrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Image2_Click()
    Unload Me
End Sub
