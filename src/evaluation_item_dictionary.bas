Attribute VB_Name = "evaluation_item_dictionary"

' Passed all test

' The format of the dictionary is as follows:
' evaluation_item_dict
'   key: evaluation item name
'   value: {id, format, sortBy, summarize, source}
'       id: String
'       format: "整數數值" | "數值" | "百分比"
'       sortBy: "遞增" | "遞減"
'       summarize: "均值" | "加總" 

' The data is stored in the worksheet "評鑑指標" in "B 參數.xlsx"
'   First row is the header
'   Column A: Evaluation item id
'   Column B: Evaluation item name
'   Column C: Evaluation item format
'   Column D: Evaluation item sortBy
'   Column E: Evaluation item summarize

Function evaluation_item_dict_init(argument_wb As Workbook) As Scripting.Dictionary
    Dim evaluation_item_dict As Scripting.Dictionary
    Dim evaluation_item_id As Variant
    Dim evaluation_item_name As Variant
    Dim evaluation_item_format As Variant
    Dim evaluation_item_sort As Variant
    Dim evaluation_item_summarize As Variant
    Dim ws As Worksheet
    Dim last_row As Long
    
    Set evaluation_item_dict = New Scripting.Dictionary
    Set ws = argument_wb.Worksheets("評鑑指標")
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To last_row
        evaluation_item_id = ws.Cells(i, 1).Value
        evaluation_item_name = ws.Cells(i, 2).Value
        evaluation_item_format = ws.Cells(i, 3).Value
        evaluation_item_sortBy = ws.Cells(i, 4).Value
        evaluation_item_summarize = ws.Cells(i, 5).Value
        
        Set evaluation_item = New Scripting.Dictionary
        evaluation_item.Add "id", evaluation_item_id
        evaluation_item.Add "format", evaluation_item_format
        evaluation_item.Add "sortBy", evaluation_item_sortBy
        evaluation_item.Add "summarize", evaluation_item_summarize
        
        evaluation_item_dict.Add evaluation_item_name, evaluation_item
    Next i
    
    Set evaluation_item_dict_init = evaluation_item_dict
End Function

' Passed test
Private Sub test_create_evaluation_item_dict()
    Dim evaluation_item_dict As Scripting.Dictionary
    Dim evaluation_item_name As Variant
    Dim evaluation_item As Variant
    Dim argument_wb As Workbook

    Set argument_wb = Workbooks.Open(ThisWorkbook.path & "/B 參數.xlsx")
    Set evaluation_item_dict = evaluation_item_dict_init(argument_wb)
    argument_wb.Close
    
    Dim file_path As String
    file_path = ThisWorkbook.path & "/output/evaluation_item_dict.json"
    print_to_file  file_path, json_str(evaluation_item_dict)
End Sub

Function evaluation_item_source_data_path(ByVal evaluation_id As String) As String
    evaluation_item_source_data_path = ThisWorkbook.path & "/0. 原始資料/output-" & evaluation_id & "_data.xls"
End Function