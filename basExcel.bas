Attribute VB_Name = "Module1"
Option Explicit

'Dim xlapp As New Excel.Application
'
'Dim xlapp_worksheet As Worksheet

'
'
'Dim sheet_count As Long
'
'Dim sheet_col_count(100) As Long
'
'Dim i, j, k As Long
'
'Dim dummy As String
'
'Dim row_data As Variant
'
'Dim row_cnt As Long
'
'Dim chk_str As String
'
'Dim dummy_max As Long
'
'Dim tot_col_count As Long
'
'Dim tot_row_count As Long
'
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
''엑셀 열기
'
'CommonDialog1.Filter = "엑셀파일|*.xls"
'
'CommonDialog1.Action = 1
'
'
'
'
'
'If CommonDialog1.FileTitle = "" Then
'
'    Exit Sub
'
'End If
'
'
'
'
'
'Label2.Caption = Trim(CommonDialog1.FileName)
'
'
'
'
'
'xlapp.Workbooks.Open (Trim(CommonDialog1.FileName))
'
'sheet_count = xlapp.Worksheets.Count
'
'
'
''전체 워크시트 불러오기와서 temp?.txt 파일로 저장
'
'
'
'For i = 1 To sheet_count
'
'
'
'    Set xlapp_worksheet = xlapp.Worksheets(i)
'
'
'
'    xlapp_worksheet.Activate
'
'    sheet_col_count(i) = xlapp_worksheet.UsedRange.Columns.Count
'
'    xlapp.DisplayAlerts = False
'
'
'
'    xlapp.ActiveWorkbook.SaveAs App.Path & "\temp\temp" & Trim(i) & ".txt", xlText, "", "", False, False '==>2000 + 2003 공용
'
'    'xlapp_worksheet.SaveAs App.Path & "\temp\temp" & Trim(i) & ".txt", xlText, "", "", False, False ==>엑셀 2000용
'
'    'ActiveWorkbook.SaveAs App.Path & "\temp\temp" & Trim(i) & ".txt", xlText, "", "", False, False  ===>엑셀 2003용
'
'Next i
'
'
'
'
'
'xlapp.Quit
'
'Set xlapp_worksheet = Nothing
'
'Set xlapp = Nothing
'
'
'
'
'
'
'
''전체 엑셀의 MAX cols값 추출
'
'dummy_max = 0
'
'For i = 1 To sheet_count
'
'    If sheet_col_count(i) >= dummy_max Then
'
'        dummy_max = sheet_col_count(i)
'
'    End If
'
'Next i
'
'tot_col_count = dummy_max
'
'
'
'
'
''전체 row값 추출
'
'For i = 1 To sheet_count
'
'    Open (App.Path & "\temp\temp" & Trim(i) & ".txt") For Input As #1
'
'        While Not EOF(1)
'
'           Line Input #1, dummy
'
'                row_data = Split(Trim(dummy), Chr(9))
'
'                chk_str = "": For j = 0 To UBound(row_data): chk_str = chk_str & row_data(j): Next j
'
'
'
'                If Len(Trim(dummy)) > 0 Then
'
'                    tot_row_count = tot_row_count + 1
'
'                End If
'
'        Wend
'
'    Close #1
'
'Next i
'
'
'
'
'
'
'
''그리드 초기화
'
'
'
'FG_execl.Cols = tot_col_count + 7
'
'FG_execl.Rows = tot_row_count + 1
'
'FG_execl.FixedCols = 3
'
'
'
''그리드에 출력
'
'For i = 1 To sheet_count
'
'    Open (App.Path & "\temp\temp" & Trim(i) & ".txt") For Input As #1
'
'        While Not EOF(1)
'
'            Line Input #1, dummy
'
'                row_data = Split(Trim(dummy), Chr(9))
'
'                chk_str = "": For j = 0 To UBound(row_data): chk_str = chk_str & row_data(j): Next j
'
'
'
'                If Len(chk_str) > 0 Then
'
'                    row_cnt = row_cnt + 1
'
'                    FG_execl.ColWidth(1) = 1300: FG_execl.ColAlignment(1) = flexAlignLeftCenter
'
'                    FG_execl.ColWidth(2) = 2000: FG_execl.ColAlignment(1) = flexAlignLeftCenter
'
'
'
'                    For j = 0 To UBound(row_data)
'
'                        FG_execl.TextMatrix(row_cnt, j + 1) = row_data(j)
'
'                    Next j
'
'
'
'
'
'
'
'                End If
'
'        Wend
'
'    Close #1
'
'Next i
'
