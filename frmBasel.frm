VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBasel 
   BackColor       =   &H8000000E&
   Caption         =   "바젤용"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   12360
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox txtLog 
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
      Left            =   180
      TabIndex        =   8
      Top             =   7980
      Width           =   9135
   End
   Begin VB.CommandButton cmdReg 
      Caption         =   "등록"
      Height          =   555
      Left            =   10500
      TabIndex        =   7
      Top             =   840
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11280
      Top             =   180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FPSpreadADO.fpSpread fpBasel 
      Height          =   6375
      Left            =   120
      TabIndex        =   6
      Top             =   1500
      Width           =   12075
      _Version        =   393216
      _ExtentX        =   21299
      _ExtentY        =   11245
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   8
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmBasel.frx":0000
   End
   Begin VB.TextBox txtSheet 
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
      Left            =   1320
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdBasel 
      Caption         =   "upload"
      Height          =   555
      Left            =   8760
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdFileSearch 
      Caption         =   "..."
      Height          =   315
      Left            =   10560
      Style           =   1  '그래픽
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox txtFileName 
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
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   9135
   End
   Begin VB.Label Label1 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "sheet 위치 : "
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   780
      Width           =   1155
   End
   Begin VB.Label Label7 
      Appearance      =   0  '평면
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      Caption         =   "파일명 :"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   540
      TabIndex        =   1
      Top             =   420
      Width           =   705
   End
End
Attribute VB_Name = "frmBasel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bOK As Boolean

Private Sub cmdBasel_Click()
    Screen.MousePointer = vbHourglass
        'Call subExcel
        Call subExcelLoad
    Screen.MousePointer = vbDefault
End Sub



Sub upLoad()
    ' Declare variables
    Dim y As Boolean, z As Boolean
    Dim Var As Variant
    Dim x As Integer, listcount As Integer, handle As Integer
    Dim List(10) As String

    ' Check if file is an Excel file and set result to x
    x = fpBasel.IsExcelFile(txtFileName.Text)

    ' If file is Excel file, tell user, import sheet
    ' list, and set result to y
    If x = 1 Then
        'MsgBox "File is an Excel file.", , "File Type"
        y = fpBasel.GetExcelSheetList(txtFileName.Text, List, listcount, "C:\ILOGFILE.TXT", handle, True)
        
MsgBox List(0)
MsgBox List(1)
MsgBox List(2)

        ' If received sheet list, tell user, import file,
        ' and set result to z
        If y = True Then
            'MsgBox "Got sheet list.", , "Status"
            z = fpBasel.ImportExcelSheet(handle, CInt(txtSheet.Text))
            ' Tell user result based on T/F value of z
            If z = True Then
                MsgBox "Import complete.", , "Result"
            Else
                MsgBox "Import did not succeed.", , "Result"
            End If
        Else
            ' Tell user cannot obtain sheet list
            MsgBox "Cannot return information for Excel file.", , "Result"
        End If
    Else
        ' Tell user file is not Excel file or is locked
        MsgBox "File is not an Excel file or is locked and cannot be imported.", , "Invalid File Type or Locked"
    End If
End Sub

Private Sub cmdFileSearch_Click()
    CommonDialog1.Filter = "엑셀파일|*.xls"
    
    CommonDialog1.Action = 1
    
     
    
     
    
    If CommonDialog1.FileTitle = "" Then
    
        Exit Sub
    
    End If
    
     
    
     
    
    txtFileName.Text = Trim(CommonDialog1.FileName)
    
     

End Sub

Private Sub Command1_Click()
    Call subExcel
End Sub

Private Sub cmdReg_Click()
Dim conn As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim sSQL As String
    

On Error GoTo ErrorHandler
    
   
    
    If GCONN_USE = False Then
        Set conn = New ADODB.Connection
        conn.ConnectionString = GetConnString
    
        conn.Open
        conn.BeginTrans
    Else
        gConn.BeginTrans
        
    End If
    
    Dim i As Integer
    
    
    Dim sColid As String
    Dim sColnm As String
    Dim sCodeDomain As String
    Dim sColBigo As String
    Dim sColBigo2 As String
    
    
    Dim EXCEPT As String
    Dim BASE_YM As String
    Dim CO_CD As String
    Dim AC_CLCD As String
    Dim CAP_ITCD As String
    Dim RWA_CALMT_TPCD As String
    Dim AMT As String


    If fpBasel.MaxRows <= 1 Then
        MsgBox "데이터를 그리드에 적재 후 등록 하세요", vbInformation
        Exit Sub
    End If
    
    If bOK = False Then
    
        MsgBox "데이터가 정상적으로 적재 되지 않았습니다.", vbInformation
        Exit Sub
    End If
    
     Screen.MousePointer = vbHourglass
    
    For i = 1 To fpBasel.MaxRows
        
        
        
        fpBasel.Row = i
        
        fpBasel.Col = 1: EXCEPT = fpBasel.Value
        fpBasel.Col = 2: BASE_YM = fpBasel.Value
        fpBasel.Col = 3: CO_CD = fpBasel.Value
        fpBasel.Col = 4: AC_CLCD = fpBasel.Value
        fpBasel.Col = 5: CAP_ITCD = fpBasel.Value
        fpBasel.Col = 6: RWA_CALMT_TPCD = fpBasel.Value
        fpBasel.Col = 7: AMT = fpBasel.Value
        
    
    
    
        If EXCEPT <> "1" Then
    
        
                sSQL = ""
                sSQL = sSQL & vbCrLf & " INSERT INTO CM_C019_TB ( BASE_YM, CO_CD , AC_CLCD , CAP_ITCD, RWA_CALMT_TPCD, AMT, CGDD) "
                sSQL = sSQL & vbCrLf & "  VALUES ('" & BASE_YM & "' ,"
                sSQL = sSQL & vbCrLf & "          '" & CO_CD & "' ,"
                sSQL = sSQL & vbCrLf & "          '" & AC_CLCD & "' ,"
                sSQL = sSQL & vbCrLf & "          '" & CAP_ITCD & "' ,"
                sSQL = sSQL & vbCrLf & "          '" & RWA_CALMT_TPCD & "' ,"
                sSQL = sSQL & vbCrLf & "           " & AMT & " ,"
                sSQL = sSQL & vbCrLf & "              to_char(sysdate,'yyyymmdd')  )"
                
                
                
                If GCONN_USE = True Then
                    gConn.Execute sSQL
                Else
                    conn.Execute sSQL
                End If
         End If
    Next i
    
    
    gfMsg ("등록되었습니다.")

    'commit
    If GCONN_USE = True Then
        gConn.CommitTrans
    Else
        conn.CommitTrans
    End If
    
    
    Screen.MousePointer = vbDefault

    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    Exit Sub
ErrorHandler:
    
    If GCONN_USE = True Then
        gConn.RollbackTrans
        
    Else
        conn.RollbackTrans
        
    End If
    
    
    Set conn = Nothing
    Set rsTemp = Nothing
    
    
    
    Screen.MousePointer = vbDefault
    
    gfMsg "데이터 등록 에러 : " & Err.Description

End Sub

Private Sub Form_Load()
    txtFileName.Text = "F:\06.VB소스\바젤\자회사 자기자본 현황 템플릿 변경_v1.xls"
    txtSheet.Text = 3
    bOK = False
End Sub


Private Sub subExcel()




Dim xlapp As New Excel.Application

Dim xlapp_worksheet As Worksheet

'
'
Dim sheet_count As Long
'
Dim sheet_col_count(100) As Long

Dim i, j, K As Long

Dim dummy As String

Dim row_data As Variant

Dim row_cnt As Long

Dim chk_str As String

Dim dummy_max As Long

Dim tot_col_count As Long

Dim tot_row_count As Long

'
'
'
'
'FG_execl.Clear
'
'FG_execl.Rows = 2
'
'FG_execl.Cols = 2
'
'
'
'tot_col_count = 0
'
'tot_row_count = 0
'
'
'
'SSTab1.Tab = 0
'
'
'
'
'
'
'
'
'
'
'
'
'
xlapp.Workbooks.Open (Trim(txtFileName.Text))
sheet_count = xlapp.Worksheets.Count

'MsgBox sheet_count
'
'
'
''전체 워크시트 불러오기와서 temp?.txt 파일로 저장
'
'
'
For i = 1 To sheet_count



    Set xlapp_worksheet = xlapp.Worksheets(i)



    xlapp_worksheet.Activate

    sheet_col_count(i) = xlapp_worksheet.UsedRange.Columns.Count

    xlapp.DisplayAlerts = False



    xlapp.ActiveWorkbook.SaveAs App.Path & "\" & Trim(i) & ".txt", xlText, "", "", False, False '==>2000 + 2003 공용
    
txtLog.Text = App.Path & "\" & Trim(i) & ".txt"
    
    'xlapp.ActiveWorkbook.SaveAs "c:\" & Trim(i) & ".txt", xlText, "", "", False, False  '==>2000 + 2003 공용

    'xlapp_worksheet.SaveAs App.Path & "\temp\temp" & Trim(i) & ".txt", xlText, "", "", False, False ==>엑셀 2000용

    'ActiveWorkbook.SaveAs App.Path & "\temp\temp" & Trim(i) & ".txt", xlText, "", "", False, False  ===>엑셀 2003용

Next i

'
'
'
'
xlapp.Quit
'
Set xlapp_worksheet = Nothing
'
Set xlapp = Nothing
'
'
'
'
'
'
'
'전체 엑셀의 MAX cols값 추출

dummy_max = 0

For i = 1 To sheet_count

    If sheet_col_count(i) >= dummy_max Then

        dummy_max = sheet_col_count(i)

    End If

Next i

tot_col_count = dummy_max

'
'
'
'
''전체 row값 추출

For i = CInt(txtSheet.Text) To CInt(txtSheet.Text)


    Open (App.Path & "\" & Trim(i) & ".txt") For Input As #1

        While Not EOF(1)

           Line Input #1, dummy

                row_data = Split(Trim(dummy), Chr(9))

                chk_str = "": For j = 0 To UBound(row_data): chk_str = chk_str & row_data(j): Next j



                If Len(Trim(dummy)) > 0 Then

                    tot_row_count = tot_row_count + 1

                End If

        Wend

    Close #1

Next i







'그리드 초기화

'
'
'FG_execl.Cols = tot_col_count + 7
'
'FG_execl.Rows = tot_row_count + 1
'
'FG_execl.FixedCols = 3
'


fpBasel.MaxRows = 0
'fpBasel.MaxRows = tot_row_count

'그리드에 출력

Dim iEx As Boolean

For i = CInt(txtSheet.Text) To CInt(txtSheet.Text)

    Open (App.Path & "\" & Trim(i) & ".txt") For Input As #1

        While Not EOF(1)

            Line Input #1, dummy

                row_data = Split(Trim(dummy), Chr(9))

                chk_str = "": For j = 0 To UBound(row_data): chk_str = chk_str & row_data(j): Next j



                If Len(chk_str) > 0 Then

                    row_cnt = row_cnt + 1


'If row_cnt = 237 Then
'    MsgBox row_cnt
'End If

'                    FG_execl.ColWidth(1) = 1300: FG_execl.ColAlignment(1) = flexAlignLeftCenter

'                    FG_execl.ColWidth(2) = 2000: FG_execl.ColAlignment(1) = flexAlignLeftCenter
                    
                    fpBasel.Row = row_cnt
                    
                    fpBasel.MaxRows = fpBasel.MaxRows + 1
                    
                    If row_cnt = 1 Then
                        fpBasel.Col = 1
                        fpBasel.Value = 1
                    End If
                    
                    
                    iEx = False
                    For j = 0 To UBound(row_data)

                        'FG_execl.TextMatrix(row_cnt, j + 1) = row_data(j)
                        'Debug.Print row_data(j)
                        
                        fpBasel.Col = j + 2
                        fpBasel.Value = row_data(j)
                        
                        If j = 0 Then
                            If Trim(row_data(j)) = "" Then
                                iEx = True
                            End If
                        End If

                    Next j
                    
                    If iEx = True Then
                        fpBasel.Col = 1
                        fpBasel.Value = 1
                    End If







                End If

        Wend

    Close #1

Next i

End Sub




























'==============================================================================
' 엑셀파일불러오기 버튼을 누른다.
'==============================================================================
Private Sub subExcelLoad()

    Dim iRowCnt_1 As Long
    
    Dim wkbObj As Workbook
    Dim i As Integer
    Dim j As Integer
    
    ' 에러고객 엑셀저장 ---------------------------------
    Dim XL As Object
    Dim WB As Object
    Dim WS As Object
    
    Dim first As Integer, last As Integer
    ' ---------------------------------------------------
    bOK = False
    
On Error GoTo ErrorHandler
    
    
    j = 0
    
    
    
    Set wkbObj = GetObject((Trim(txtFileName.Text)))

'    iRowCnt_1 = 1
'    ' 엑셀의 A1 데이터가 없을경우 정보가 없는걸로 인식해서 ERROR 처리
'    If Trim(wkbObj.Worksheets(1).Range("A" & iRowCnt_1).Value) = "" Then
'        gfMsgWrn ("등록할 내용이 없습니다.")
'        GoTo ErrorExit
'    End If
    
    '---------------------------------
'    Set XL = GetObject(, "Excel.Application")
'    Set WB = XL.Workbooks.Add()
'    Set WS = WB.Worksheets.Add
    '----------------------------------------------------
    
    Screen.MousePointer = vbHourglass

    
    Dim Data1() As String
    Dim K As Double
    
    ReDim Data1(0) As String
    
    Dim iRow As Integer
    K = 3
    iRow = 0
    fpBasel.MaxRows = 0
    Do
        K = K + 1
        
        'AX 값이 "" 일경우 마지막으로 인식한다
        If K > 5 Then
            If Trim(wkbObj.Worksheets(CInt(txtSheet)).Range("A" & K).Value) = "" Then
                'K = K - 1
                Exit Do
            End If
        End If
        

        fpBasel.MaxRows = fpBasel.MaxRows + 1
        iRow = iRow + 1
        fpBasel.Row = iRow
        fpBasel.Col = 2
        fpBasel.Value = Trim(wkbObj.Worksheets(CInt(txtSheet)).Range("A" & K).Value)
        
        fpBasel.Col = 3
        fpBasel.Value = Trim(wkbObj.Worksheets(CInt(txtSheet)).Range("B" & K).Value)
        
        fpBasel.Col = 4
        fpBasel.Value = Trim(wkbObj.Worksheets(CInt(txtSheet)).Range("C" & K).Value)
        
        fpBasel.Col = 5
        fpBasel.Value = Trim(wkbObj.Worksheets(CInt(txtSheet)).Range("D" & K).Value)
        
        fpBasel.Col = 6
        fpBasel.Value = Trim(wkbObj.Worksheets(CInt(txtSheet)).Range("E" & K).Value)
        
        fpBasel.Col = 7

        fpBasel.Value = wkbObj.Worksheets(CInt(txtSheet)).Range("F" & K).Value
        
        fpBasel.Col = 8
        fpBasel.Value = Trim(wkbObj.Worksheets(CInt(txtSheet)).Range("G" & K).Value)

   
                        
    Loop
    bOK = True
    MsgBox "정상적으로 로드 되었습니다.", vbInformation
    
ErrorExit:

    Screen.MousePointer = vbDefault
    Set WS = Nothing
    Set WB = Nothing
    Set XL = Nothing
    'Set obj = Nothing
    Set wkbObj = Nothing
    Exit Sub

ErrorHandler:
    Screen.MousePointer = vbDefault
    Set WS = Nothing
    Set WB = Nothing
    Set XL = Nothing
    Set wkbObj = Nothing
    MsgBox "로드중 에러가 발생하였습니다. 데이터를 확인하세요 " & vbCrLf & "에러메세지 : " & Err.Description, vbCritical
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If frmBasel.Tag = "YES" Then
        End
    End If
End Sub
