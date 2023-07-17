Attribute VB_Name = "import_data"
' 2023-07-10 12:00:08


' Passed all test

' �s�J�Ҧ����мƾڪ��r��
Dim evaluation_items_value_dict As Scripting.Dictionary


' ================================================= import data =================================================

Sub import_all()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    argument_init
    Dim college_list As Collection
    Dim evaluation_item_list As Collection
    
    Set college_list = New Collection
    For Each college In college_department_dict.Keys
        college_list.Add college
    Next college

    Set evaluation_item_list = New Collection
    For Each evaluation_item In evaluation_item_dict.Keys
        evaluation_item_list.Add evaluation_item
    Next evaluation_item
    
    import_data college_list, evaluation_item_list

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

' * Import data from "0. ��l���" to "1. �U�|�J����"
Function import_data(college_list As Collection, evaluation_item_list As Collection)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
        
    Dim argument_wb As Workbook
    Dim destin_wb As Workbook

    Set evaluation_items_value_dict = evaluation_items_value_dict_init(evaluation_item_list)
        
    For Each college In college_list
        Set destin_wb = Workbooks.Open(college_excel_path(college))
        import_data_to_wb destin_wb, college, evaluation_item_list
        write_summary_table destin_wb, college
        destin_wb.Close True
    Next college

        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Function

' Passed test
Private Sub test_import_data()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    argument_init

    Dim college_list As Collection
    Set college_list = New Collection
    college_list.Add "�F�v�j��"
    college_list.Add "��ǰ|"
    college_list.Add "�z�ǰ|"

    Dim evaluation_item_list As Collection
    Set evaluation_item_list = New Collection
    evaluation_item_list.Add "�Ǥh�Z�c�P���ˤJ�ǿ����v"
    evaluation_item_list.Add "�դh�Z�ۦ��ꤺ���I�j�ǲ��~�ͤ�v"
    evaluation_item_list.Add "�Ǥh�Z����U�Ǫ��������B"

    import_data college_list, evaluation_item_list

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub


' * Import data
' Parameter:
'   wb: the workbook to import data
'   college: the college name
'   evaluation_item_list: the list of evaluation item
Function import_data_to_wb( wb as Workbook, _
                            ByVal college as String, _
                            evaluation_item_list as Collection)
    Dim ws As Worksheet
    Dim shtname As String
    Dim evaluation_item As Variant
    Dim evaluation_item_dict As Scripting.Dictionary
    Dim college_value_dict As Scripting.Dictionary

    ' �N��ƶפJ��U�Ӥu�@��(�U�ӫ��е�Ų����)
    For Each evaluation_item In evaluation_item_list
        
        ' �ӫ��Ъ����
        Set evaluation_item_dict = evaluation_items_value_dict(evaluation_item)

        ' �ӫ��иӾǰ|�����
        college_id = college_department_dict(college)(1)("id")
        college_with_id = college_id & " " & college
        Set college_value_dict = evaluation_item_dict("evaluation_value_dict")(college_with_id)
        
        ' �ӫ��ЩҦb���u�@��
        shtname = evaluation_item_dict("id") & " " & evaluation_item
        Set ws = wb.Worksheets(shtname)

        ' �N��ƶפJ��u�@��
        import_data_to_ws ws, college_value_dict, evaluation_item_dict("format")

        ' �ͦ����R��r
        generate_analytic_text ws, college, evaluation_item        
    Next evaluation_item
    
End Function


' Passed test
Private Sub test_import_data_to_wb()

    Dim evaluation_item_list As Collection
    Set evaluation_item_list = New Collection
    evaluation_item_list.Add "�Ǥh�Z�c�P���ˤJ�ǿ����v"
    evaluation_item_list.Add "�դh�Z�ۦ��ꤺ���I�j�ǲ��~�ͤ�v"
    
    argument_init

    Set evaluation_items_value_dict = evaluation_items_value_dict_init(evaluation_item_list)

    from_path = ThisWorkbook.path & "\1. �U�|�J����\�z�ǰ|.xlsx"
    to_path = ThisWorkbook.path & "\output\import_data_to_wb.xlsx"
    copy_file from_path, to_path

    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = Workbooks.Open(to_path)
    college = "�z�ǰ|"
    import_data_to_wb wb, college, evaluation_item_list
End Sub


' * Save the data to the worksheet
' Column A: department_name
' Column C: avg
' Column D: year3
' Column E: year2
' Column F: year1
' Column G: rank
' Parameter:
'   ws: the worksheet to save data
'   college_value_dict: the dict of the college's data
'       key: department_name
'       value: {avg, year3, year2, year1, rank}
'   format: the format of the data
Function import_data_to_ws(ws as Worksheet, college_value_dict as Scripting.Dictionary, format as String)
    
    Dim department_name As Variant
    Dim department_value_dict As Scripting.Dictionary
    Dim row As Integer

    last_row = college_value_dict.Count + 1

    For row = 2 to last_row
        if ws.Cells(row, 1).Value = "" Then
            last_row = row - 1
            Exit For
        End If

        ' ��ڲĤ@�檺�Ǩt�W�١A���o�ӾǨt�����
        department_name = ws.Cells(row, 1).Value
        Set department_value_dict = college_value_dict(department_name)
                
        ws.Cells(row, 3).Value = department_value_dict("avg")
        ws.Cells(row, 4).Value = department_value_dict("year3")
        ws.Cells(row, 5).Value = department_value_dict("year2")
        ws.Cells(row, 6).Value = department_value_dict("year1")
        ws.Cells(row, 7).Value = department_value_dict("rank")

    Next row

    Dim rng As Range
    Set rng = ws.Range("C2:F" & last_row)

    ' �۰ʽվ���e
    rng.Columns.AutoFit

    ' �]�w�榡
    set_number_format rng, format

    ' �]�w�Ϫ� y �b���ƭȮ榡
    set_axis_number_format ws, format

    Set rng = ws.Range("A2:G" & last_row)

    ' �M�οz��
    filter_rng rng
End Function

' Passed test
Private Sub test_import_data_to_ws()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim evaluation_item_list As Collection
    Set evaluation_item_list = New Collection
    evaluation_item_list.Add "�Ǥh�Z�c�P���ˤJ�ǿ����v"
    
    Dim evaluation_items_value_dict As Scripting.Dictionary
    Set evaluation_items_value_dict = evaluation_items_value_dict_init(evaluation_item_list)

    Dim college_value_dict As Scripting.Dictionary
    Set college_value_dict = evaluation_items_value_dict("�Ǥh�Z�c�P���ˤJ�ǿ����v")("evaluation_value_dict")("700 �z�ǰ|")

    from_path = ThisWorkbook.path & "\1. �U�|�J����\�z�ǰ|.xlsx"
    to_path = ThisWorkbook.path & "\output\import_data_to_ws.xlsx"
    copy_file from_path, to_path

    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = Workbooks.Open(to_path)
    Set ws = wb.Worksheets("1.1.1.1 �Ǥh�Z�c�P���ˤJ�ǿ����v")
    
    import_data_to_ws ws, college_value_dict
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

' =============================================== format =====================================================

Function set_number_format(rng As Range, ByVal number_format As String)
    
    Select Case number_format
        Case "��Ƽƭ�"
            rng.NumberFormat = "#,##0;-#,###;0"
        Case "�ƭ�"
            rng.NumberFormat = "#,##0.##;-#,##0.##;0"
        Case "�ʤ���"
            rng.NumberFormat = "###.##%;-###.##%;0%"
    End Select

    For Each c In rng
        If c.Value = "�X" Then
            c.NumberFormat = """�X"""
        ElseIf right(c.Text, 2) = ".%" Then
            c.NumberFormat = "###%;###%;0%"
        ElseIf right(c.Text, 1) = "." Then
            c.NumberFormat = "#,###;#,###;0"
        End If
    Next c
End Function

Function set_axis_number_format(ws As Worksheet, ByVal number_format As String)

    Set y_axis = ws.ChartObjects(1).Chart.Axes(xlValue)
    unit = y_axis.MajorUnit

    Select Case number_format
        Case "��Ƽƭ�"
            y_axis.TickLabels.NumberFormat = "#,##0;-#,###;0"
        Case "�ƭ�"
            If unit < 0.1 Then
                y_axis.TickLabels.NumberFormat = "0.00;-0.00;0"
            ElseIf unit<1 Then
                y_axis.TickLabels.NumberFormat = "0.0;-0.0;0"
            Else
                y_axis.TickLabels.NumberFormat = "#,##0;-#,###;0"
            End If
        Case "�ʤ���"
            If unit < 0.001 Then
                y_axis.TickLabels.NumberFormat = "0.00%;-0.00%;0"
            ElseIf unit<0.01 Then
                y_axis.TickLabels.NumberFormat = "0.0%;-0.0%;0"
            Else
                y_axis.TickLabels.NumberFormat = "0%;-0%;0"
            End If
    End Select

End Function



' =============================================== Filter =====================================================

' * Filter in the worksheet
' Filter the data in the worksheet by rank(Col G)
' Hide the rows which rank = '�X' and sort the data by rank
' Parameter:
'   ws: the worksheet to filter
Function filter_rng(rng As Range)
    rng.Sort Key1:=rng.Range("G1"), Order1:=xlAscending
    rng.AutoFilter Field:=7, Criteria1:="<>�X"
End Function

' Passed test
Private Sub test_filter_ws()
    Dim wb As Workbook
    Dim ws As Worksheet
    from_path = ThisWorkbook.path & "\input\filter_ws.xlsx"
    to_path = ThisWorkbook.path & "\output\filter_ws.xlsx"
    copy_file from_path, to_path
    Set wb = Workbooks.Open(to_path)
    Set ws = wb.Worksheets("1.1.1.1 �Ǥh�Z�c�P���ˤJ�ǿ����v")
    filter_ws ws
    wb.Close True
End Sub

' =============================================== Analytic Text =====================================================


' * �ھګ��мƾڥͦ����R��r
' Parameter:
'   ws: the worksheet to generate analytic text
'   college: the name of the college
'   evaluation_item: the name of the evaluation item
Function generate_analytic_text(ws As worksheet, ByVal college As String, ByVal evauation_item As String)

    sortBy = evaluation_item_dict(evauation_item)("sortBy")
    summarize = evaluation_item_dict(evauation_item)("summarize")

    Set result_range = ws.Range("I1")

    If college = "�F�v�j��" Then
        level = "��"
        child_level = "�|"
    Else
        level = "�|"
        child_level = "�t"
    End If

    Dim department_dict As Scripting.Dictionary
    college_id = college_department_dict(college)(1)("id")
    college_with_id = college_id & " " & college
    
    ' ���լO�_�����
    n = college_department_dict(college).Count
    empty_count = 0
    zero_count = 0
    For r = 2 To n+1
        If ws.Cells(r, 1) = "" Then
            Exit For
        End If
        If ws.Cells(r, 3) = "�X" Then
            empty_count = empty_count + 1
        End If
        If ws.Cells(r, 3) = 0 Then
            zero_count = zero_count + 1
        End If
    Next r
    If empty_count >= n - 1 Then
        result_range.Value = "�L��ơC"
        Exit Function
    End If
    college_avg = ws.Cells(n - empty_count + 1, 3).Value
    college_avg_text = ws.Cells(n - empty_count + 1, 3).Text

    ' �ھګ��бƧǤ覡�A�o�X�ݦC�X���t�Ҽƶq
    If summarize = "�[�`" Then
        
        result = ""
        If college_avg_text = "0" Then
            result =  level & "�[�`�T�~���Ȭ�0�C"
            result_range.Value = result
            Exit Function
        End If
        
        For r = 2 To n - empty_count - zero_count - 1
            curr_dept = ws.Cells(r, 2)
            curr_str = ws.Cells(r, 3).Text
            next_str = ws.Cells(r+1, 3).Text
            result = result & curr_dept & curr_str & "�B"
        Next r
        r = n - empty_count - zero_count
        result = result & ws.Cells(r, 2) & ws.Cells(r, 3).Text & "�C"
        
        result = level & "�[�`�T�~���Ȭ�" & college_avg_text & "�A�]�t" & result
        result_range.Value = result
        Exit Function
    End If

    ' summarize = "����"
    bigger_count = 0
    smaller_count = 0
    equal_count = 0
    For r = 2 To n - empty_count
        curr_avg = ws.Cells(r, 3).Value
        If curr_avg > college_avg Then
            bigger_count = bigger_count + 1
        ElseIf curr_avg < college_avg Then
            smaller_count = smaller_count + 1
        Else
            equal_count = equal_count + 1
        End If
    Next r
    
    result = level & "�T�~���Ȭ�" & college_avg_text & "�A"
    If sortBy = "���W" And smaller_count + equal_count = 0 Then
        result = result & "�L�C��]�ε���^" & level & "�T�~���Ȫ̡C"
        result_range.Value = result
        Exit Function
    ElseIf sortBy = "����" And bigger_count + equal_count = 0 Then
        result = result & "�L����]�ε���^" & level & "�T�~���Ȫ̡C"
        result_range.Value = result
        Exit Function
    End If

    If sortBy = "���W" Then
        result = result & "�C��"
        If equal_count <> 0 Then
            result = result & "�]�ε���^"
        End If
        result = result &  level & "�T�~���Ȫ̭p��" & smaller_count + equal_count & "��" & child_level & "�A��"

        For r = 2 To smaller_count + equal_count
            curr_dept = ws.Cells(r, 2)
            curr_str = ws.Cells(r, 3).Text
            next_str = ws.Cells(r+1, 3).Text
            If ws.Cells(r, 3).Value < college_avg Then
                result = result & curr_dept
                If curr_str <> next_str Then
                    result = result & curr_str
                End If
                result = result & "�B"
            End If
        Next r
        r = smaller_count + equal_count + 1
        result = result & ws.Cells(r, 2) & ws.Cells(r, 3).Text & "�C"
    Else
        result = result & "����"
        If equal_count <> 0 Then
            result = result & "�]�ε���^"
        End If
        result = result &  level & "�T�~���Ȫ̭p��" & bigger_count + equal_count & "��" & child_level & "�A��"

        For r = 2 To bigger_count + equal_count
            curr_dept = ws.Cells(r, 2)
            curr_str = ws.Cells(r, 3).Text
            next_str = ws.Cells(r+1, 3).Text
            If ws.Cells(r, 3).Value > college_avg Then
                result = result & curr_dept
                If curr_str <> next_str Then
                    result = result & curr_str
                End If
                result = result & "�B"
            End If
        Next r
        r = bigger_count + equal_count + 1
        result = result & ws.Cells(r, 2) & ws.Cells(r, 3).Text & "�C"
    End If
    result_range.Value = result
    
End Function
    
' Passed test
Private Sub test_generate_analytic_text()
    argument_init

    from_path = ThisWorkbook.path & "\input\generate_analytic_text.xlsx"
    to_path = ThisWorkbook.path & "\output\generate_analytic_text.xlsx"
    copy_file from_path, to_path

    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = Workbooks.Open(to_path)
    Set ws = wb.Sheets("1.1.1.1 �Ǥh�Z�c�P���ˤJ�ǿ����v")

    Dim items As Collection
    Set items = New Collection
    items.Add "�Ǥh�Z�c�P���ˤJ�ǿ����v"

    generate_analytic_text ws, "��ǰ|", "�Ǥh�Z�c�P���ˤJ�ǿ����v"
End Sub


' ============================================================= �p����� =============================================================
' TODO: �ͦ��p�����
' ���D�C�O���ЦW�١B���D��O�Ǩt�W��
' �U�檺�ȬO�ӾǨt�� rank
' Parameter:
'   wb: the workbook to generate summary table
'   evaluation_items_value_dict: the dict of the evaluation items' data 
Function write_summary_table(wb As Workbook, ByVal college As String)
    ' column 1 is evaluation item id, column 2 is evaluation item name
    ' row 1 is department full name, row 2 is department abbreviation
    ' write the rank of each department in each evaluation item
    ws_name = "�p��"
    Dim ws As Worksheet
    Set ws = wb.Sheets(ws_name)

    r = 3
    max_row = ws.Cells(2, 2).End(xlDown).Row

    c = 3
    max_column = ws.Cells(2,2).End(xlToRight).Column

    college_id = college_department_dict(college)(1)("id")
    college_with_id = college_id & " " & college


    start = 3
    For r = 3 To max_row
        evaluation_item = ws.Cells(r, 2).Value

        If ws.Cells(r, 1).Value <> "����" Then
            Set department_ranks = evaluation_items_value_dict(evaluation_item)("evaluation_value_dict")(college_with_id)
        End if

        For c = 3 To max_column
            If ws.Cells(r, 1).Value = "����" Then
                On Error GoTo SetAvgDash
                avg = Application.WorksheetFunction.Average(ws.Range(ws.Cells(start, c), ws.Cells(r-1, c)))
                avg = Round(avg, 2)
                ws.Cells(r, c).Value = avg
                GoTo NoError
SetAvgDash:
                ws.Cells(r, c).Value = "-"
NoError:
            Else
                department = ws.Cells(1, c).Value
                ws.Cells(r, c).Value = department_ranks(department)("rank")
            End If
        Next c
    Next r
    
    ws.Range(ws.Cells(3, 3), ws.Cells(max_row, max_column)).VerticalAlignment = xlCenter
    ws.Range(ws.Cells(3, 3), ws.Cells(max_row, max_column)).HorizontalAlignment = xlCenter

End Function

Private Sub test_write_summary_table()
    argument_init

    Dim evaluation_item_list As Collection
    Set evaluation_item_list = New Collection

    For Each evaluation_item In evaluation_item_dict.Keys
        evaluation_item_list.Add evaluation_item
    Next evaluation_item

    Set evaluation_items_value_dict = evaluation_items_value_dict_init(evaluation_item_list)

    from_path = ThisWorkbook.path & "\input\write_summary_table.xlsx"
    to_path = ThisWorkbook.path & "\output\write_summary_table.xlsx"
    copy_file from_path, to_path

    Dim wb As Workbook
    Set wb = Workbooks.Open(to_path)

    write_summary_table wb, "��ǰ|"
End Sub

    

    