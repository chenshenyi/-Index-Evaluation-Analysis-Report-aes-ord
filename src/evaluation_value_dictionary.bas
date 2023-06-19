Attribute VB_Name = "evaluation_value_dictionary"

' The format of the dictionary is as follows:
' evaluation_items_value_dict:
'   key: evaluation_item_name
'   value: {id, format, sort, summarize, evaluation_value_dict}

'   evaluation_value_dict:
'       key: college_name
'       value: {department_name: department_value_dict}

'       department_value_dict: {avg, year3, year2, year1}

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
'   value: {id, format, sort, summarize, evaluation_value_dict}
Function evaluation_items_value_dict_init(argument_wb As Workbook, evaluation_item_list As Collection) As Scripting.Dictionary
    Dim evaluation_items_value_dict As Scripting.Dictionary
    Dim evaluation_item As Scripting.Dictionary
    Dim id As String
    Dim summarize As String
    Dim wb As Workbook
    Dim ws As Worksheet

    Set evaluation_items_value_dict = evaluation_item_dict_init(argument_wb)

    For Each evaluation_item_name In evaluation_item_list
        Set evaluation_item = evaluation_items_value_dict(evaluation_item_name)
        summarize = evaluation_item("summarize")
        id = evaluation_item("id")
 
        Set wb = Workbooks.Open(source_path(id))
        Set ws = wb.Worksheets("近三年比較")

        evaluation_item.add "evaluation_value_dict", evaluation_value_dict_init(ws, summarize)

        wb.Close
    Next evaluation_item_name

    Set evaluation_items_value_dict_init = evaluation_items_value_dict
End Function

Private Sub test_evaluation_items_value_dict_init()
    Application.ScreenUpdating = False
    application.DisplayAlerts = False

    Dim argument_wb As Workbook
    Set argument_wb = Workbooks.Open(ThisWorkbook.path & "/B 參數.xlsx")

    Dim evaluation_item_list As Collection
    Set evaluation_item_list = New Collection
    evaluation_item_list.Add "學士班繁星推薦入學錄取率"
    evaluation_item_list.Add "博士班招收國內重點大學畢業生比率"
    evaluation_item_list.Add "學士班獲獎助學金平均金額"
    evaluation_item_list.Add "碩士班平均修業年限"
    evaluation_item_list.Add "平均碩博士班修課學生人數"
    evaluation_item_list.Add "各系所教師兼任本校二級學術行政主管人次"
    evaluation_item_list.Add "舉辦國際學術研討會數"

    Dim evaluation_items_value_dict As Scripting.Dictionary
    Set evaluation_items_value_dict = evaluation_items_value_dict_init(argument_wb, evaluation_item_list)
    
    Dim file_path As String
    file_path = ThisWorkbook.path & "/output/evaluation_items_value_dict.json"
 
    print_to_file  file_path, json_str(evaluation_items_value_dict)

    argument_wb.Close
End Sub

' evaluation_value_dict:
'   key: college_name
'   value: {department_name: department_value_dict}
Function evaluation_value_dict_init(ws As Worksheet, summarize As Variant) As Scripting.Dictionary
    Dim evaluation_value_dict As Scripting.Dictionary
    Dim school_dict As Scripting.Dictionary

    Dim college_name As String
    Dim college_dict As Scripting.Dictionary

    Dim department_name As String
    Dim department_value_dict As Scripting.Dictionary

    ' 特別標出校級資料
    Set school_dict = New Scripting.Dictionary
    school_dict.Add ws.Range("B9").Text, department_value_dict_init(ws, 9, summarize)

    Set evaluation_value_dict = New Scripting.Dictionary
    evaluation_value_dict.Add ws.Range("A9").Text, school_dict

    Dim row As Integer
    row = 10

    Do While ws.Range("B" & row) <> ""
        
        department_name = ws.Range("B" & row)
        Set department_value_dict = department_value_dict_init(ws, row, summarize)

        If ws.Range("A" & row) <> "" Then
            college_name = ws.Range("A" & row)
            school_dict.Add department_name, department_value_dict ' 將各院資料儲存至校的字典中
            Set college_dict = New Scripting.Dictionary
            evaluation_value_dict.Add college_name, college_dict
        End If

        evaluation_value_dict(college_name).Add department_name, department_value_dict

        row = row + 1
    Loop

    Set evaluation_value_dict_init = evaluation_value_dict
End Function

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
Function department_value_dict_init(ws As Worksheet, row As Integer, summarize As Variant) As Scripting.Dictionary
    Dim department_value_dict As Scripting.Dictionary
    Set department_value_dict = New Scripting.Dictionary

    department_value_dict.Add "avg", reformulate_value(ws.Range("E" & row), summarize)
    department_value_dict.Add "year3", reformulate_value(ws.Range("H" & row), summarize)
    department_value_dict.Add "year2", reformulate_value(ws.Range("K" & row), summarize)
    department_value_dict.Add "year1", reformulate_value(ws.Range("N" & row), summarize)

    Set department_value_dict_init = department_value_dict
End Function

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
Function reformulate_value(value As String, summarize As Variant) As String
    If Not InStr(value, " /") = 0 Then
        ' use split, " /" as delimiter
        Select Case summarize
            Case "加總"
                value = Split(value, " /")(0)
            Case "均值"
                value = Split(value, " /")(1)
        End Select
    End If

    ' Test if "%" exists in the cell
    If InStr(value, "%") Then
        value = Left(value, Len(value) - 1) / 100
    End If

    reformulate_value = value
End Function

Private Sub test_reformulate_value()
    Debug.Print reformulate_value("345.00 /8.82%", "加總")
    Debug.Print reformulate_value("345.00 /8.82%", "均值")
    Debug.Print reformulate_value("345.00", "加總")
End Sub

