' filename: evaluation_value_dictionary.cls

' The format of the dictionary is as follows:
'   college evaluation value dictionary:
'       key:   college name
'       value:  {department name: department evaluation value}
'           department evaluation value: {avg, one_year_ago, two_year_ago, three_year_ago}

' The data is stored in the following format:
'   A: college name
'   B: department name
'   E: avg
'   H: one_year_ago
'   K: two_year_ago
'   N: three_year_ago
'The first row is row 9

' This module use the following dictionary:
' college_department_dict
'   key: college name
'   value: [{id, name, abbr}] (list of department dictionary)
' evaluation_item_dict
'   key: evaluation item name
'   value: {id, format, sort, summarize}
'       id: String
'       format: "整數數值" | "數值" | "百分比"
'       sort: "遞增" | "遞減"
'       summarize: "均值" | "加總"

' This

Function create_department_evaluation_value_dict(source_ws As worksheet, row As Integer, summarize As Variant) As Scripting.Dictionary
    Dim department_evaluation_value_dict As Scripting.Dictionary
    Set department_evaluation_value_dict = New Scripting.Dictionary

    department_evaluation_value_dict.Add "avg", reformulate_value(source_ws.Range("E" & row), summarize)
    department_evaluation_value_dict.Add "one_year_ago", reformulate_value(source_ws.Range("H" & row), summarize)
    department_evaluation_value_dict.Add "two_year_ago", reformulate_value(source_ws.Range("K" & row), summarize)
    department_evaluation_value_dict.Add "three_year_ago", reformulate_value(source_ws.Range("N" & row), summarize)

    Set create_department_evaluation_value_dict = department_evaluation_value_dict
End Function

Private Sub test_create_department_evaluation_value_dict()
    Dim source_ws As worksheet
    Dim source_wb As Workbook

    pth = ThisWorkbook.Path & "\0. 原始資料\output-1.1.1.1_data.xls"
    Set source_wb = Workbooks.Open(pth)
    Set source_ws = source_wb.Worksheets("近三年比較")

    Dim department_evaluation_value_dict As Scripting.Dictionary
    Set department_evaluation_value_dict = create_department_evaluation_value_dict(source_ws, 9, "加總")

    Debug.Print department_evaluation_value_dict("avg")
    Debug.Print department_evaluation_value_dict("one_year_ago")
    Debug.Print department_evaluation_value_dict("two_year_ago")
    Debug.Print department_evaluation_value_dict("three_year_ago")
End Sub

Function reformulate_value(value As String, summarize As Variant) As Double
    ' Test if " /" exists in the cell
    If Not InStr(value, " /") = 0 Then
        ' `345.00 /8.82%` -> `345.00` if avg_or_sum = "加總"
        ' `345.00 /8.82%` -> `8.82%` if avg_or_sum = "均值"
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

