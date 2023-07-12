Attribute VB_Name = "evaluation_value_dictionary"
' 2023-07-11 17:14:41

' Passed All Test

' The format of the dictionary is as follows:
' evaluation_items_value_dict:
'   key: evaluation_item_name
'   value: {id, format, sortBy, summarize, evaluation_value_dict}

'   evaluation_value_dict:
'       key: college_name
'       value: {department_name: department_value_dict}

'       department_value_dict: {avg, year3, year2, year1, rank}

' The data is stored in the following format:
' The first row is row 9
' Columns:
'   A: college name(only the first row of each college)
'   B: department name
'   E: avg
'   H: year3
'   K: year2
'   N: year1

' evaluation_items_value_dict:
'   key: evaluation_item_name
'   value: {id, format, sortBy, summarize, evaluation_value_dict}
' Parameter:
'   argument_wb: the workbook of the argument file 'B �Ѽ�.xlsx'
'   evaluation_item_list: the list of evaluation item names
Function evaluation_items_value_dict_init(evaluation_item_list As Collection) As Scripting.Dictionary
    Dim evaluation_items_value_dict As Scripting.Dictionary
    Dim evaluation_item As Scripting.Dictionary
    Dim id As String
    Dim summarize As String
    Dim sortBy As String
    Dim wb As Workbook
    Dim ws As Worksheet

    Dim argument_wb As Workbook
    Set argument_wb = Workbooks.Open(ThisWorkbook.path & "/B �Ѽ�.xlsx")
    Set evaluation_items_value_dict = evaluation_item_dict_init(argument_wb)
    argument_wb.Close

    For Each evaluation_item_name In evaluation_item_list
        Set evaluation_item = evaluation_items_value_dict(evaluation_item_name)
        
        summarize = evaluation_item("summarize")
        id = evaluation_item("id")
        sortBy = evaluation_item("sortBy")
 
        Set wb = Workbooks.Open(evaluation_item_source_data_path(id))
        Set ws = wb.Worksheets("��T�~���")

        evaluation_item.add "evaluation_value_dict", evaluation_value_dict_init(ws, summarize)        
        wb.Close

        Call evaluation_items_rank_eq(evaluation_item("evaluation_value_dict"), sortBy)

    Next evaluation_item_name

    Set evaluation_items_value_dict_init = evaluation_items_value_dict
End Function

' Passed test
Private Sub test_evaluation_items_value_dict_init()
    Application.ScreenUpdating = False
    application.DisplayAlerts = False

    Dim evaluation_item_list As Collection
    Set evaluation_item_list = New Collection
    evaluation_item_list.Add "�Ǥh�Z�c�P���ˤJ�ǿ����v"
    evaluation_item_list.Add "�դh�Z�ۦ��ꤺ���I�j�ǲ��~�ͤ�v"
    ' evaluation_item_list.Add "�Ǥh�Z����U�Ǫ��������B"
    ' evaluation_item_list.Add "�Ӥh�Z�����׷~�~��"
    ' evaluation_item_list.Add "�����ӳդh�Z�׽ҾǥͤH��"
    ' evaluation_item_list.Add "�U�t�ұЮv�ݥ����դG�žǳN��F�D�ޤH��"
    ' evaluation_item_list.Add "�|���ھǳN��Q�|��"

    Dim evaluation_items_value_dict As Scripting.Dictionary
    Set evaluation_items_value_dict = evaluation_items_value_dict_init(evaluation_item_list)
    
    Dim file_path As String
    file_path = ThisWorkbook.path & "/output/evaluation_items_value_dict1.csv"
    print_to_file  file_path, csv_expand_last_str(evaluation_items_value_dict("�Ǥh�Z�c�P���ˤJ�ǿ����v")("evaluation_value_dict"))

    file_path = ThisWorkbook.path & "/output/evaluation_items_value_dict2.csv"
    print_to_file  file_path, csv_expand_last_str(evaluation_items_value_dict("�դh�Z�ۦ��ꤺ���I�j�ǲ��~�ͤ�v")("evaluation_value_dict"))

    file_path = ThisWorkbook.path & "/output/evaluation_items_value_dict.json"
    print_to_file  file_path, json_str(evaluation_items_value_dict)

    argument_wb.Close
End Sub

' evaluation_value_dict:
'   key: college_name
'   value: {department_name: department_value_dict}
' parameter:
'   ws: '��T�~���' of original data
'   summarize: "�[�`" or "����"
Function evaluation_value_dict_init(ws As Worksheet, ByVal summarize As String) As Scripting.Dictionary
    Dim evaluation_value_dict As Scripting.Dictionary
    Dim school_dict As Scripting.Dictionary

    Dim college_name As String
    Dim college_dict As Scripting.Dictionary

    Dim department_name As String
    Dim department_value_dict As Scripting.Dictionary

    ' �կŸ�Ʀr��
    Set school_dict = New Scripting.Dictionary
    school_dict.Add ws.Range("B9").Text, department_value_dict_init(ws, 9, summarize)

    ' �N�կŸ�Ʀr���x�s�ܫ��мƾڦr�夤
    Set evaluation_value_dict = New Scripting.Dictionary
    evaluation_value_dict.Add ws.Range("A9").Text, school_dict
    
    ' �N�U�t����x�s�ܰ|���r�夤
    Dim row As Integer
    row = 10
    Do While ws.Range("B" & row) <> ""
        
        department_name = ws.Range("B" & row)
        Set department_value_dict = department_value_dict_init(ws, row, summarize)

        ' �p�G�O�s���|
        If ws.Range("A" & row) <> "" Then
            college_name = ws.Range("A" & row)
            school_dict.Add department_name, department_value_dict_init(ws, row, summarize)' �N�U�|����x�s�ܮժ��r�夤
            Set college_dict = New Scripting.Dictionary
            evaluation_value_dict.Add college_name, college_dict
        End If

        ' �N�U�t����x�s�ܰ|���r�夤
        evaluation_value_dict(college_name).Add department_name, department_value_dict_init(ws, row, summarize)

        ' Caution: department_value_dict should be different in each place of dictionary

        row = row + 1
    Loop

    Set evaluation_value_dict_init = evaluation_value_dict
End Function

' Passed test
Private Sub test_evaluation_value_dict()
    Dim wb As Workbook
    Dim ws As worksheet

    pth = ThisWorkbook.Path & "\0. ��l���\output-1.1.1.1_data.xls"
    Set wb = Workbooks.Open(pth)
    Set ws = wb.Worksheets("��T�~���")

    Dim evaluation_value_dict As Scripting.Dictionary
    Set evaluation_value_dict = evaluation_value_dict_init(ws, "�[�`")

    Dim file_path As String
    file_path = ThisWorkbook.path & "/output/evaluation_value_dict.json"
    print_to_file  file_path, json_str(evaluation_value_dict)
End Sub

' department_value_dict: {avg, year3, year2, year1}
' parameter:
'   ws: '��T�~���' of original data
'   row: row number of data
'   summarize: "�[�`" or "����"
Function department_value_dict_init(ws As Worksheet, ByVal row As Integer, ByVal summarize As String) As Scripting.Dictionary
    Dim department_value_dict As Scripting.Dictionary
    Set department_value_dict = New Scripting.Dictionary

    ' �N�U������x�s�ܨt���r�夤
    department_value_dict.Add "avg", reformulate_value(ws.Range("E" & row), summarize)
    department_value_dict.Add "year3", reformulate_value(ws.Range("H" & row), summarize)
    department_value_dict.Add "year2", reformulate_value(ws.Range("K" & row), summarize)
    department_value_dict.Add "year1", reformulate_value(ws.Range("N" & row), summarize)

    Set department_value_dict_init = department_value_dict
End Function

' Passed test
Private Sub test_department_value_dict_init()
    Dim wb As Workbook
    Dim ws As worksheet

    pth = ThisWorkbook.Path & "\0. ��l���\output-1.1.1.1_data.xls"
    Set wb = Workbooks.Open(pth)
    Set ws = wb.Worksheets("��T�~���")

    Dim department_value_dict As Scripting.Dictionary
    Set department_value_dict = department_value_dict_init(ws, 13, "�[�`")

    Debug.Print json_str(department_value_dict)
End Sub

' if " /" exists in the cell
' `345.00 /8.82%` -> `345.00` if avg_or_sum = "�[�`"
' `345.00 /8.82%` -> `0.0882` if avg_or_sum = "����"
' parameter:
'   value: cell value
'   summarize: "�[�`" or "����"
Function reformulate_value(ByVal value As String, ByVal summarize As String) As String
    On Error GoTo ErrorHandler

    ' �p�G�� " /"�A�h�ھ� summarize ���X�ƭ�
    If InStr(value, " /") Then
        Select Case summarize
            Case "�[�`"
                value = Split(value, " /")(0)
            Case "����"
                value = Split(value, " /")(1)
        End Select
    End If
    
    If value = "-1" or value = "" Then
        value = "�X"
    End If

    ' �N�ʤ����ഫ���p��
    If InStr(value, "%") Then
        value = Left(value, Len(value) - 1) / 100
    End If

    reformulate_value = value
    Exit Function
ErrorHandler:
    reformulate_value = "�X"
    Resume Next
End Function

' Passed test
Private Sub test_reformulate_value()
    Debug.Print reformulate_value("345.00 /8.82%", "�[�`")
    Debug.Print reformulate_value("345.00 /8.82%", "����")
    Debug.Print reformulate_value("345.00", "�[�`")
End Sub

' =========================================== rank ===========================================

' * Apply college_rank_eq to each college
' Parameter:
'   evaluation_value_dict: the evaluation_value_dict initialized by `evaluation_value_dict_init`
'   sortBy: '���W' or '����'
Function evaluation_items_rank_eq(evaluation_value_dict As Scripting.Dictionary, ByVal sortBy As String)
    For Each college_name In evaluation_value_dict
        college_rank_eq evaluation_value_dict, college_name, sortBy
    Next college_name
End Function

' Passed test
Private Sub test_evaluation_items_rank_eq()
    Dim wb As Workbook
    Dim ws As worksheet

    pth = ThisWorkbook.Path & "\0. ��l���\output-1.1.1.1_data.xls"
    Set wb = Workbooks.Open(pth)
    Set ws = wb.Worksheets("��T�~���")

    Dim evaluation_value_dict As Scripting.Dictionary
    Set evaluation_value_dict = evaluation_value_dict_init(ws, "����")

    evaluation_items_rank_eq evaluation_value_dict, "���W"

    file_path = ThisWorkbook.path & "/output/evaluation_items_rank.json"
    print_to_file  file_path, json_str(evaluation_value_dict)
End Sub

' * Add 'rank' to each department
' * (comparing to the other departments in the same college)
' Parameter:
'   evaluation_value_dict: the evaluation_value_dict initialized by `evaluation_value_dict_init`
'   college_name: the name of the college with college id(eg. '100 ��ǰ|')
'   sortBy: '���W' or '����'
Function college_rank_eq(evaluation_value_dict As Scripting.Dictionary, ByVal college_name As String, ByVal sortBy As String)
    On Error GoTo ErrorHandler

    Dim department_avg_list As Collection
    Set department_avg_list = New Collection

    For Each department_name In evaluation_value_dict(college_name)
        ' �p�G department_name �e�T�X�M college_name �e�T�X�ۦP �άO avg �O "�X" �N���[�J
        avg = evaluation_value_dict(college_name)(department_name)("avg")
        If Left(department_name, 3) <> Left(college_name, 3) And avg<>"�X" Then
            department_avg_list.Add avg
        End if
    Next department_name

    Dim department_dict As Scripting.Dictionary
    For Each department_name In evaluation_value_dict(college_name)
        Set department_dict = evaluation_value_dict(college_name)(department_name)
        avg = department_dict("avg")

        If Left(department_name, 3) = Left(college_name, 3) Then
            department_dict.Add "rank", 999
        ElseIf avg = "�X" Then
           department_dict.Add "rank", "�X"
        Else
            rank = rank_eq(department_avg_list, avg, sortBy)
            department_dict.Add "rank", rank
        End if
    Next department_name

    Exit Function
    
ErrorHandler:
    MsgBox "Error: " & "college_rank_eq" & vbCrLf & "college_name: " & college_name
End Function

' Passed test
Private Sub test_college_rank_eq()

    Dim evaluation_value_dict As Scripting.Dictionary
    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = Workbooks.Open(ThisWorkbook.path & "/0. ��l���/output-1.1.1.1_data.xls")
    Set ws = wb.Worksheets("��T�~���")
    Set evaluation_value_dict = evaluation_value_dict_init(ws, "����")

    Dim college_name1 As String
    Dim college_name2 As String
    Dim college_name3 As String
    college_name1 = "100 ��ǰ|"
    college_name2 = "700 �z�ǰ|"
    college_name3 = "200 ���|��ǰ|"
    

    college_rank_eq evaluation_value_dict, college_name1, "���W"
    college_rank_eq evaluation_value_dict, college_name2, "����"
    college_rank_eq evaluation_value_dict, college_name3, "����"

    Dim file_path As String
    file_path = ThisWorkbook.path & "/output/college_rank_eq.json"
    print_to_file file_path, json_str(evaluation_value_dict(college_name1))
    file_path = ThisWorkbook.path & "/output/college_rank_eq.csv"
    print_to_file file_path, csv_str(evaluation_value_dict(college_name1)) & vbCrLf & vbCrLf & _
                            csv_str(evaluation_value_dict(college_name2)) & vbCrLf & vbCrLf & _
                            csv_str(evaluation_value_dict(college_name3))

End Sub

' Parameter:
'   value_list: collection of all values
'   v: value to be ranked
'   sortBy: "���W" or "����"
Function rank_eq(value_list As Collection, ByVal v As Double, ByVal sortBy As String) As Integer
    On Error GoTo ErrorHandler
    

    Dim i As Integer
    Dim rank As Integer
    Dim list_length As Integer
    Dim list_item As Double

    list_length = value_list.Count
    rank = 0

    For i = 1 To list_length
        list_item = value_list(i)
        If sortBy = "���W" Then
            If v > list_item Then
                rank = rank + 1
            End If
        ElseIf sortBy = "����" Then
            If v < list_item Then
                rank = rank + 1
            End If
        End If
    Next i
    rank_eq = rank +1
    Exit Function
ErrorHandler:
    MsgBox "Error in rank_eq: " & vbCrLf & _
            "Can't rank " & v & " in " & json_str(value_list) & vbCrLf

    rank_eq = 0
End Function

' Passed test
Private Sub test_rank_eq()
    Dim sorted_list As Collection
    Set sorted_list = New Collection
    sorted_list.Add 1
    sorted_list.Add 2
    sorted_list.Add 3
    sorted_list.Add 3
    sorted_list.Add 5

    Debug.Print rank_eq(sorted_list, 1, "���W")
    Debug.Print rank_eq(sorted_list, 2, "���W")
    Debug.Print rank_eq(sorted_list, 3, "����")
    Debug.Print rank_eq(sorted_list, 4, "����")
    Debug.Print rank_eq(sorted_list, 5, "����")
End Sub

