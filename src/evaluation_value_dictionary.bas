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
'   argument_wb: the workbook of the argument file 'B 參數.xlsx'
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
    Set argument_wb = Workbooks.Open(ThisWorkbook.path & "/B 參數.xlsx")
    Set evaluation_items_value_dict = evaluation_item_dict_init(argument_wb)
    argument_wb.Close

    For Each evaluation_item_name In evaluation_item_list
        Set evaluation_item = evaluation_items_value_dict(evaluation_item_name)
        
        summarize = evaluation_item("summarize")
        id = evaluation_item("id")
        sortBy = evaluation_item("sortBy")
 
        Set wb = Workbooks.Open(evaluation_item_source_data_path(id))
        Set ws = wb.Worksheets("近三年比較")

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
    evaluation_item_list.Add "學士班繁星推薦入學錄取率"
    evaluation_item_list.Add "博士班招收國內重點大學畢業生比率"
    ' evaluation_item_list.Add "學士班獲獎助學金平均金額"
    ' evaluation_item_list.Add "碩士班平均修業年限"
    ' evaluation_item_list.Add "平均碩博士班修課學生人數"
    ' evaluation_item_list.Add "各系所教師兼任本校二級學術行政主管人次"
    ' evaluation_item_list.Add "舉辦國際學術研討會數"

    Dim evaluation_items_value_dict As Scripting.Dictionary
    Set evaluation_items_value_dict = evaluation_items_value_dict_init(evaluation_item_list)
    
    Dim file_path As String
    file_path = ThisWorkbook.path & "/output/evaluation_items_value_dict1.csv"
    print_to_file  file_path, csv_expand_last_str(evaluation_items_value_dict("學士班繁星推薦入學錄取率")("evaluation_value_dict"))

    file_path = ThisWorkbook.path & "/output/evaluation_items_value_dict2.csv"
    print_to_file  file_path, csv_expand_last_str(evaluation_items_value_dict("博士班招收國內重點大學畢業生比率")("evaluation_value_dict"))

    file_path = ThisWorkbook.path & "/output/evaluation_items_value_dict.json"
    print_to_file  file_path, json_str(evaluation_items_value_dict)

    argument_wb.Close
End Sub

' evaluation_value_dict:
'   key: college_name
'   value: {department_name: department_value_dict}
' parameter:
'   ws: '近三年比較' of original data
'   summarize: "加總" or "均值"
Function evaluation_value_dict_init(ws As Worksheet, ByVal summarize As String) As Scripting.Dictionary
    Dim evaluation_value_dict As Scripting.Dictionary
    Dim school_dict As Scripting.Dictionary

    Dim college_name As String
    Dim college_dict As Scripting.Dictionary

    Dim department_name As String
    Dim department_value_dict As Scripting.Dictionary

    ' 校級資料字典
    Set school_dict = New Scripting.Dictionary
    school_dict.Add ws.Range("B9").Text, department_value_dict_init(ws, 9, summarize)

    ' 將校級資料字典儲存至指標數據字典中
    Set evaluation_value_dict = New Scripting.Dictionary
    evaluation_value_dict.Add ws.Range("A9").Text, school_dict
    
    ' 將各系資料儲存至院的字典中
    Dim row As Integer
    row = 10
    Do While ws.Range("B" & row) <> ""
        
        department_name = ws.Range("B" & row)
        Set department_value_dict = department_value_dict_init(ws, row, summarize)

        ' 如果是新的院
        If ws.Range("A" & row) <> "" Then
            college_name = ws.Range("A" & row)
            school_dict.Add department_name, department_value_dict_init(ws, row, summarize)' 將各院資料儲存至校的字典中
            Set college_dict = New Scripting.Dictionary
            evaluation_value_dict.Add college_name, college_dict
        End If

        ' 將各系資料儲存至院的字典中
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

    pth = ThisWorkbook.Path & "\0. 原始資料\output-1.1.1.1_data.xls"
    Set wb = Workbooks.Open(pth)
    Set ws = wb.Worksheets("近三年比較")

    Dim evaluation_value_dict As Scripting.Dictionary
    Set evaluation_value_dict = evaluation_value_dict_init(ws, "加總")

    Dim file_path As String
    file_path = ThisWorkbook.path & "/output/evaluation_value_dict.json"
    print_to_file  file_path, json_str(evaluation_value_dict)
End Sub

' department_value_dict: {avg, year3, year2, year1}
' parameter:
'   ws: '近三年比較' of original data
'   row: row number of data
'   summarize: "加總" or "均值"
Function department_value_dict_init(ws As Worksheet, ByVal row As Integer, ByVal summarize As String) As Scripting.Dictionary
    Dim department_value_dict As Scripting.Dictionary
    Set department_value_dict = New Scripting.Dictionary

    ' 將各欄位資料儲存至系的字典中
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

    pth = ThisWorkbook.Path & "\0. 原始資料\output-1.1.1.1_data.xls"
    Set wb = Workbooks.Open(pth)
    Set ws = wb.Worksheets("近三年比較")

    Dim department_value_dict As Scripting.Dictionary
    Set department_value_dict = department_value_dict_init(ws, 13, "加總")

    Debug.Print json_str(department_value_dict)
End Sub

' if " /" exists in the cell
' `345.00 /8.82%` -> `345.00` if avg_or_sum = "加總"
' `345.00 /8.82%` -> `0.0882` if avg_or_sum = "均值"
' parameter:
'   value: cell value
'   summarize: "加總" or "均值"
Function reformulate_value(ByVal value As String, ByVal summarize As String) As String
    On Error GoTo ErrorHandler

    ' 如果有 " /"，則根據 summarize 取出數值
    If InStr(value, " /") Then
        Select Case summarize
            Case "加總"
                value = Split(value, " /")(0)
            Case "均值"
                value = Split(value, " /")(1)
        End Select
    End If
    
    If value = "-1" or value = "" Then
        value = "—"
    End If

    ' 將百分比轉換為小數
    If InStr(value, "%") Then
        value = Left(value, Len(value) - 1) / 100
    End If

    reformulate_value = value
    Exit Function
ErrorHandler:
    reformulate_value = "—"
    Resume Next
End Function

' Passed test
Private Sub test_reformulate_value()
    Debug.Print reformulate_value("345.00 /8.82%", "加總")
    Debug.Print reformulate_value("345.00 /8.82%", "均值")
    Debug.Print reformulate_value("345.00", "加總")
End Sub

' =========================================== rank ===========================================

' * Apply college_rank_eq to each college
' Parameter:
'   evaluation_value_dict: the evaluation_value_dict initialized by `evaluation_value_dict_init`
'   sortBy: '遞增' or '遞減'
Function evaluation_items_rank_eq(evaluation_value_dict As Scripting.Dictionary, ByVal sortBy As String)
    For Each college_name In evaluation_value_dict
        college_rank_eq evaluation_value_dict, college_name, sortBy
    Next college_name
End Function

' Passed test
Private Sub test_evaluation_items_rank_eq()
    Dim wb As Workbook
    Dim ws As worksheet

    pth = ThisWorkbook.Path & "\0. 原始資料\output-1.1.1.1_data.xls"
    Set wb = Workbooks.Open(pth)
    Set ws = wb.Worksheets("近三年比較")

    Dim evaluation_value_dict As Scripting.Dictionary
    Set evaluation_value_dict = evaluation_value_dict_init(ws, "均值")

    evaluation_items_rank_eq evaluation_value_dict, "遞增"

    file_path = ThisWorkbook.path & "/output/evaluation_items_rank.json"
    print_to_file  file_path, json_str(evaluation_value_dict)
End Sub

' * Add 'rank' to each department
' * (comparing to the other departments in the same college)
' Parameter:
'   evaluation_value_dict: the evaluation_value_dict initialized by `evaluation_value_dict_init`
'   college_name: the name of the college with college id(eg. '100 文學院')
'   sortBy: '遞增' or '遞減'
Function college_rank_eq(evaluation_value_dict As Scripting.Dictionary, ByVal college_name As String, ByVal sortBy As String)
    On Error GoTo ErrorHandler

    Dim department_avg_list As Collection
    Set department_avg_list = New Collection

    For Each department_name In evaluation_value_dict(college_name)
        ' 如果 department_name 前三碼和 college_name 前三碼相同 或是 avg 是 "—" 就不加入
        avg = evaluation_value_dict(college_name)(department_name)("avg")
        If Left(department_name, 3) <> Left(college_name, 3) And avg<>"—" Then
            department_avg_list.Add avg
        End if
    Next department_name

    Dim department_dict As Scripting.Dictionary
    For Each department_name In evaluation_value_dict(college_name)
        Set department_dict = evaluation_value_dict(college_name)(department_name)
        avg = department_dict("avg")

        If Left(department_name, 3) = Left(college_name, 3) Then
            department_dict.Add "rank", 999
        ElseIf avg = "—" Then
           department_dict.Add "rank", "—"
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

    Set wb = Workbooks.Open(ThisWorkbook.path & "/0. 原始資料/output-1.1.1.1_data.xls")
    Set ws = wb.Worksheets("近三年比較")
    Set evaluation_value_dict = evaluation_value_dict_init(ws, "均值")

    Dim college_name1 As String
    Dim college_name2 As String
    Dim college_name3 As String
    college_name1 = "100 文學院"
    college_name2 = "700 理學院"
    college_name3 = "200 社會科學院"
    

    college_rank_eq evaluation_value_dict, college_name1, "遞增"
    college_rank_eq evaluation_value_dict, college_name2, "遞減"
    college_rank_eq evaluation_value_dict, college_name3, "遞減"

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
'   sortBy: "遞增" or "遞減"
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
        If sortBy = "遞增" Then
            If v > list_item Then
                rank = rank + 1
            End If
        ElseIf sortBy = "遞減" Then
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

    Debug.Print rank_eq(sorted_list, 1, "遞增")
    Debug.Print rank_eq(sorted_list, 2, "遞增")
    Debug.Print rank_eq(sorted_list, 3, "遞減")
    Debug.Print rank_eq(sorted_list, 4, "遞減")
    Debug.Print rank_eq(sorted_list, 5, "遞減")
End Sub

