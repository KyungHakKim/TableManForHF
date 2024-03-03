VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Begin VB.Form frmVarList 
   Caption         =   "변수 치환 목록"
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11355
   Icon            =   "frmVarList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   11355
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdClose 
      Caption         =   "닫기"
      Height          =   375
      Left            =   9240
      TabIndex        =   3
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "행추가"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "저장"
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   6600
      Width           =   1815
   End
   Begin FPSpreadADO.fpSpread fpVariable 
      Height          =   6015
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   10995
      _Version        =   393216
      _ExtentX        =   19394
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
      MaxCols         =   4
      MaxRows         =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      ShadowColor     =   12648447
      ShadowDark      =   12632256
      SpreadDesigner  =   "frmVarList.frx":048A
      ScrollBarTrack  =   1
   End
End
Attribute VB_Name = "frmVarList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    fpVariable.MaxRows = fpVariable.MaxRows + 1
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrHandler
    Dim ffile As Integer
    Dim sTemp  As String
    Dim sTemp2() As String
    Dim iCnt As Integer
    Dim i As Integer
    
    

    ffile = FreeFile

    Open App.Path & "\" & "변수치환.txt" For Output As #ffile   ' 파일을 엽니다.
    
    
    For i = 1 To fpVariable.MaxRows
        
        fpVariable.Row = i
        fpVariable.Col = 1
        sTemp = fpVariable.Value
    
        fpVariable.Row = i
        fpVariable.Col = 2
        sTemp = sTemp & ";" & fpVariable.Value
    
        fpVariable.Row = i
        fpVariable.Col = 3
        sTemp = sTemp & ";" & fpVariable.Value
    
        fpVariable.Row = i
        fpVariable.Col = 4
        sTemp = sTemp & ";" & fpVariable.Value
    
        Print #ffile, sTemp
    Next i
    
    
    Close #ffile
    
    gfMsgInf ("저장 되었습니다.!!")
    Exit Sub
ErrHandler:
    gfMsgInf (Err.Description)
End Sub

Private Sub Form_Load()

    Call ReadFile
    

End Sub


Private Sub ReadFile()
On Error GoTo ErrHandler
    Dim ffile As Integer
    Dim sTemp  As String
    Dim sTemp2() As String
    Dim iCnt As Integer
    Dim i As Integer
    
    

    ffile = FreeFile

    Open App.Path & "\" & "변수치환.txt" For Input As #ffile   ' 파일을 엽니다.
    
    fpVariable.MaxRows = 0
    Do
    
        Input #ffile, sTemp
        
        sTemp2 = Split(sTemp, ";")
        
        fpVariable.MaxRows = fpVariable.MaxRows + 1
        For i = LBound(sTemp2) To UBound(sTemp2)
            
            fpVariable.Row = fpVariable.MaxRows
            
            fpVariable.Col = i + 1
            fpVariable.Value = sTemp2(i)
        Next i
        
    Loop Until EOF(1)
    
    Close #ffile
    Exit Sub
ErrHandler:
    gfMsgInf (Err.Description)

End Sub
