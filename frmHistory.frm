VERSION 5.00
Begin VB.Form frmHistory 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '���� ����
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
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ݱ�"
      Height          =   435
      Left            =   7740
      Style           =   1  '�׷���
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
      IMEMode         =   8  '����
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  '����
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
    subAddHistroy ("   �����ڵ� ��Ÿ ���� TRIM ó�� ")
    subAddHistroy (" ")
    
    subAddHistroy ("ver 1.35    2011.01.05 ")
    subAddHistroy (" ")
    subAddHistroy ("   ���ν��� �ҽ� ���� ��� �߰� ")
    subAddHistroy (" ")
    subAddHistroy ("ver 1.35    2011.01.05 ")
    subAddHistroy (" ")
    subAddHistroy ("   ���ν��� ��� �޺��ڽ��� re-load ��ư ��� �߰� ")
    subAddHistroy (" ")
    subAddHistroy ("ver 1.34    2011.01.03 ")
    subAddHistroy (" ")
    subAddHistroy ("   ����� �ǿ� Stored Procedure�� �޺��ڽ��� ���� ")
    
End Sub

Private Sub subAddHistroy(sText As String)
    txtHistory.Text = txtHistory.Text + sText + vbCrLf

End Sub
