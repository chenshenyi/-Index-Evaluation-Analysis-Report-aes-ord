' Function merge_worksheet_in_multiple_workbooks(workbook_path_list As Variant, _
'                                                 worksheet_name As String, _
'                                                 output_workbook As Workbook, _
'                                                 output_worksheet_name_list As Variant, _
'                                                 Optional output_worksheet_name_prefix As String = "", _
'                                                 Optional output_worksheet_name_suffix As String = "")
'     Dim output_worksheet As Worksheet
'     Dim output_worksheet_name As String
    
'     For i = 0 To UBound(workbook_path_list)
'         Set input_workbook = Workbooks.Open(workbook_path_list(i))
'         Set input_worksheet = input_workbook.Worksheets(worksheet_name)
        
'         output_worksheet_name = output_worksheet_name_prefix & input_workbook.Name & output_worksheet_name_suffix
        
'         ' If the output worksheet name is already in the output workbook, then delete it
'         If Not IsError(Application.Match(output_worksheet_name, output_worksheet_name_list, 0)) Then
'             output_workbook.Worksheets(output_worksheet_name).Delete
'         End If

'         ' Copy the input worksheet to the end of output workbook
'         input_worksheet.Copy After:=output_workbook.Worksheets(output_workbook.Worksheets.Count)
'         Set output_worksheet = output_workbook.Worksheets(output_workbook.Worksheets.Count)
'         output_worksheet.Name = output_worksheet_name

'         input_workbook.Close
'     Next i
' End Function

Function reformulate_worksheets(wb As Workbook, _
                                ws_name As String, _
                                avg_or_sum As String)
    Dim ws As Worksheet
    Dim last_row As Long
    
    Set ws = wb.Worksheets(ws_name)
    last_row = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    ' Re-formulate the cells in Column C to N
    ' * The range of the cells should according to the source data
    For i = 3 To 14
        For j = 9 To last_row ' 9 is the first row of the data
            ' Test if " /" exists in the cell
            If InStr(ws.Cells(j, i).Value, " /") = 0 Then
                ' Change the cell to the number format
                ' e.g. "123" -> 123, "123.45" -> 123.45, "23.02%" -> 23.02%

                If ws.Cells(j, i).Value <> "" Then
                    ws.Cells(j, i).Value = ws.Cells(j, i).Value
                End If
            
                GoTo next_j
            End If

            ' Re-formulate the cell with two datas
            ' `345.00 /8.82%` -> `345.00` if avg_or_sum = "加總"
            ' `345.00 /8.82%` -> `8.82%` if avg_or_sum = "均值" 

            If avg_or_sum = "加總" Then
                ws.Cells(j, i).Value = Split(ws.Cells(j, i).Value, " /")(0)
            ElseIf avg_or_sum = "均值" Then
                ws.Cells(j, i).Value = Split(ws.Cells(j, i).Value, " /")(1)
            End If
next_j:
        Next j
    Next i

End Function

Sub test_reformulate_worksheets()
    Dim wb As Workbook

    Set wb = Workbooks.Open(ThisWorkbook.path & "\origin-data-rearrangement.xlsx")
    Call reformulate_worksheets(wb, "近三年比較", "加總")
End Sub