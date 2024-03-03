VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Begin VB.Form frmDong 
   Caption         =   "동음이의어"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12225
   Icon            =   "frmDong.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   12225
   StartUpPosition =   3  'Windows 기본값
   Begin FPSpreadADO.fpSpread fpDong 
      Height          =   6015
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   11895
      _Version        =   393216
      _ExtentX        =   20981
      _ExtentY        =   10610
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   14737632
      MaxCols         =   3
      MaxRows         =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      ShadowColor     =   12648447
      ShadowDark      =   12632256
      SpreadDesigner  =   "frmDong.frx":048A
      ScrollBarTrack  =   1
   End
End
Attribute VB_Name = "frmDong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
On Error GoTo Err_Handler
    Dim i As Integer
    
    fpDong.MaxRows = 0


    For i = LBound(gsDongEum1) To UBound(gsDongEum1)
        fpDong.MaxRows = fpDong.MaxRows + 1
        With fpDong
            .Row = i + 1
            .Col = 1: .Value = gsDongEum1(i)
            .Col = 2: .Value = gsDongEum2(i)
            .Col = 3: .Value = gsDongEum3(i)
        End With
    Next i
    
    Exit Sub
Err_Handler:
    gfMsgInf (Err.Description)
End Sub
