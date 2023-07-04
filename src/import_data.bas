Attribute VB_Name = "import_data"

' Passed all test

' TODO: add filter to each worksheet


' ================================================= import data =================================================

' * Import data from "0. ��l���" to "1. �U�|�J����"
Function import_data(college_list As Collection, evaluation_item_list As Collection)
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Dim argument_wb As Workbook
    Dim college_department_dict As Scripting.Dictionary
    Dim evaluation_items_value_dict As Scripting.Dictionary
    Dim destin_wb As Workbook
    
    ' Create the dictionary of college and evaluation item by "B �Ѽ�.xlsx"
    Set argument_wb = Workbooks.Open(ThisWorkbook.path & "/B �Ѽ�.xlsx")
    Set college_department_dict = college_department_dict_init(argument_wb)
    Set evaluation_items_value_dict = evaluation_items_value_dict_init(argument_wb, evaluation_item_list)
    argument_wb.Close

    For Each college In college_list
        Set destin_wb = Workbooks.Open(college_excel_path(college))
        
        import_data_to_wb destin_wb, college, evaluation_item_list, evaluation_items_value_dict
        destin_wb.Close True
    Next college
End Function

' Passed test
Private Sub test_import_data()
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
End Sub


' * Import data
' Parameter:
'   wb: the workbook to import data
'   college_with_id: the college id and name
'   evaluation_item_list: the list of evaluation item
'   evaluation_items_value_dict: the dict initialized by evaluation_items_value_dict_init (����`�M��)
Function import_data_to_wb( wb as Workbook, _
                            ByVal college as String, _
                            evaluation_item_list as Collection, _
                            evaluation_items_value_dict as Scripting.Dictionary)
    Dim ws As Worksheet
    Dim shtname As String
    Dim evaluation_item As Variant
    Dim evaluatation_item_dict As Scripting.Dictionary
    Dim college_value_dict As Scripting.Dictionary

    ' �N��ƶפJ��U�Ӥu�@��(�U�ӫ��е�Ų����)
    For Each evaluation_item In evaluation_item_list
        
        ' �ӫ��Ъ����
        Set evaluatation_item_dict = evaluation_items_value_dict(evaluation_item)

        ' �ӫ��иӾǰ|�����
        college_id = college_department_dict(college)(1)("id") ' Caution: the college's information is in the first item of the dict
        college_with_id = college_id & " " & college
        Set college_value_dict = evaluatation_item_dict("evaluation_value_dict")(college_with_id)
        
        ' �ӫ��ЩҦb���u�@��
        shtname = evaluatation_item_dict("id") & " " & evaluation_item
        Set ws = wb.Worksheets(shtname)

        ' �N��ƶפJ��u�@��
        import_data_to_ws ws, college_value_dict

        ' �M�οz�ﾹ
        Call filter_ws ws

        ' �ͦ����R��r
        Call generate_analysis_text ws, college
        
    Next evaluation_item
    
End Function

' Passed test
Private Sub test_import_data_to_wb()
    Dim argument_wb As Workbook
    Set argument_wb = Workbooks.Open(ThisWorkbook.path & "/B �Ѽ�.xlsx")

    Dim evaluation_item_list As Collection
    Set evaluation_item_list = New Collection
    evaluation_item_list.Add "�Ǥh�Z�c�P���ˤJ�ǿ����v"
    evaluation_item_list.Add "�դh�Z�ۦ��ꤺ���I�j�ǲ��~�ͤ�v"
    
    Dim evaluation_items_value_dict As Scripting.Dictionary
    Set evaluation_items_value_dict = evaluation_items_value_dict_init(argument_wb, evaluation_item_list)
    argument_wb.Close

        Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = Workbooks.Open(ThisWorkbook.path & "\1. �U�|�J����\�z�ǰ|.xlsx")
    college_with_id = "700 �z�ǰ|"
    import_data_to_wb wb, college_with_id, evaluation_item_list, evaluation_items_value_dict
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
Function import_data_to_ws(ws as Worksheet, college_value_dict as Scripting.Dictionary)
    Dim department_name As Variant
    Dim department_value_dict As Scripting.Dictionary
    Dim row As Integer
    
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For row = 2 to last_row

        ' ��ڲĤ@�檺�Ǩt�W�١A���o�ӾǨt�����
        department_name = ws.Cells(row, 1).Value
        Set department_value_dict = college_value_dict(department_name)
                
        ws.Cells(row, 3).Value = department_value_dict("avg")
        ws.Cells(row, 4).Value = department_value_dict("year3")
        ws.Cells(row, 5).Value = department_value_dict("year2")
        ws.Cells(row, 6).Value = department_value_dict("year1")
        ws.Cells(row, 7).Value = department_value_dict("rank")

    Next row
End Function

' Passed test
Private Sub test_import_data_to_ws()
    Dim argument_wb As Workbook
    Set argument_wb = Workbooks.Open(ThisWorkbook.path & "/B �Ѽ�.xlsx")

    Dim evaluation_item_list As Collection
    Set evaluation_item_list = New Collection
    evaluation_item_list.Add "�Ǥh�Z�c�P���ˤJ�ǿ����v"
    
    Dim evaluation_items_value_dict As Scripting.Dictionary
    Set evaluation_items_value_dict = evaluation_items_value_dict_init(argument_wb, evaluation_item_list)
    argument_wb.Close

    Dim college_value_dict As Scripting.Dictionary
    Set college_value_dict = evaluation_items_value_dict("�Ǥh�Z�c�P���ˤJ�ǿ����v")("evaluation_value_dict")("700 �z�ǰ|")

    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = Workbooks.Open(ThisWorkbook.path & "\1. �U�|�J����\�z�ǰ|.xlsx")
    Set ws = wb.Worksheets("1.1.1.1 �Ǥh�Z�c�P���ˤJ�ǿ����v")
    
    import_data_to_ws ws, college_value_dict
End Sub


' =============================================== Filter =====================================================

' * Filter in the worksheet
' Filter the data in the worksheet by rank(Col G)
' Hide the rows which rank = '-' and sort the data by rank
' Parameter:
'   ws: the worksheet to filter
Function filter_ws(ws as Worksheet)
    Dim last_row As Integer
    Dim filter_range As Range

    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Set filter_range = ws.Range(ws.Cells(2, 1), ws.Cells(last_row, 7))
    filter_range.AutoFilter Field:=7, Criteria1:="<>-"
    filter_range.Sort Key1:=ws.Range("G2"), Order1:=xlAscending
End Function

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
Function generate_analytic_text(ws As workbook, college As String, evaluatation_item_dict As Scripting.Dictionary)
    format = evaluatation_item_dict("format")
    sortBy = evaluatation_item_dict("sortBy")
    summarize = evaluatation_item_dict("summarize")

    Set result_range = ws.Range("Q1")

    Dim department_dict As Scripting.Dictionary
    college_id = college_department_dict(college)(1)("id")
    college_with_id = college_id & " " & college

    Dim college_value_dict As Scripting.Dictionary
    Set college_value_dict = evaluatation_item_dict("evaluation_value_dict")(college_with_id)
    college_avg = college_value_dict(1)("avg")
    
    If college = "�F�v�j��" Then
        level = "��"
        child_level = "�|"
    Else
        level = "�|"
        child_level = "�t"
    End If

    ' ���լO�_�����
    n = college_department_dict(college).Count
    empty_count = 0
    For r = 2 To n+1
        If ws.Cells(r, 1) = "" Then
            Exit For
        End If
        If ws.Cells(r, 3) = "-" Then
            empty_count = empty_count + 1
        End If
    Next r
    If empty_count = n Then
        result_range.Value = "�L��ơC"
        Exit Function
    End If

    ' �ھګ��бƧǤ覡�A�o�X�ݦC�X���t�Ҽƶq
    If summarize = "�[�`" Then
        result = ""
        
        For r = 2 To n - empty_count - 1
            curr_dept = ws.Cells(r, 2)
            curr_str = ws.Cells(r, 3).Text
            next_str = ws.Cells(r+1, 3).Text
            result = result & curr_dept
            If curr_str <> next_str Then
                result = result & curr_str
            End If
            result = result & "�B"
        Next r
        r = n - empty_count
        result = result & ws.Cells(r, 2) & ws.Cells(r, 3).Text & "�C"
        
        result = level & "�[�`�T�~���Ȭ�" & college_avg & "�A�]�t" & result
        result_range.Value = result
        Exit Function
    End If

    ' summarize = "����"
    bigger_count = 0
    smaller_count = 0
    equal_count = 0
    For r = 2 To n - empty_count
        curr_avg = ws.Cells(r, 3)
        If curr_avg > college_avg Then
            bigger_count = bigger_count + 1
        ElseIf curr_avg < college_avg Then
            smaller_count = smaller_count + 1
        Else
            equal_count = equal_count + 1
        End If
    Next r
    
    
    If sortBy = "���W" Then
        result = "�C��"
        If equal_count <> 0 Then
            result = result & "�]�ε���^"
        End If
        result = result &  level & "�T�~���Ȫ̭p��" & smaller_count + equal_count & "��" & child_level & "�A��"

        For r = 2 To smaller_count + equal_count
            curr_dept = ws.Cells(r, 2)
            curr_str = ws.Cells(r, 3).Text
            next_str = ws.Cells(r+1, 3).Text
            If ws.Cells(r, 3) < college_avg Then
                result = result & curr_dept
                If curr_str <> next_str Then
                    result = result & curr_str
                End If
                result = result & "�B"
            End If
        Next r
        r = smaller_count + equal_count
        result = result & ws.Cells(r, 2) & ws.Cells(r, 3).Text & "�C"
    Else
        result = "����"
        If equal_count <> 0 Then
            result = result & "�]�ε���^"
        End If
        result = result &  level & "�T�~���Ȫ̭p��" & bigger_count + equal_count & "��" & child_level & "�A��"

        For r = 2 To bigger_count + equal_count
            curr_dept = ws.Cells(r, 2)
            curr_str = ws.Cells(r, 3).Text
            next_str = ws.Cells(r+1, 3).Text
            If ws.Cells(r, 3) > college_avg Then
                result = result & curr_dept
                If curr_str <> next_str Then
                    result = result & curr_str
                End If
                result = result & "�B"
            End If
        Next r
        r = bigger_count + equal_count
        result = result & ws.Cells(r, 2) & ws.Cells(r, 3).Text & "�C"
    End If
    
End Function
    
' TODO: ���g���դ��



' ============================================================= �p�`��� =============================================================
' TODO: �ͦ��p�`���
' ���D�C�O���ЦW�١B���D��O�Ǩt�W��
' �U�檺�ȬO�ӾǨt�� rank
' Parameter:
'   wb: the workbook to generate summary table
'   evaluation_items_value_dict: the dict of the evaluation items' data 
Function write_summary_table(wb As Workbook, evaluation_items_value_dict As Scripting.Dictionary, college As String)
    For Each evaluation_item_name In evaluation_items_value_dict

        Set evaluation_item_dict = evaluation_items_value_dict(evaluation_item_name)

        college_with_id = college_department_dict(college)(1)("id") & " " & college
        
        Set college_value_dict = evaluation_item_dict("evaluation_value_dict")(college_with_id)
        Set department_dict = college_department_dict(college_id & " " & college_name)

        Set ws = wb.Worksheets
    Next evaluation_item_name
End Function
    