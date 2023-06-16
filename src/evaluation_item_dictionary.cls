' filename: evaluation_item_dictionary

' The format of the dictionary is as follows:
' evaluation_item_dict
'   key: evaluation item name
'   value: {id, format, sort, summarize}
'       id: String
'       format: "整數數值" | "數值" | "百分比"
'       sort: "遞增" | "遞減"
'       summarize: "均值" | "加總"

' The data is stored in the worksheet "評鑑指標" in "B 參數.xlsx"
'   First row is the header
'   Column A: Evaluation item id
'   Column B: Evaluation item name
'   Column C: Evaluation item format
'   Column D: Evaluation item sort
'   Column E: Evaluation item summarize

Function create_evaluation_item_dict(wb As Workbook) As Scripting.Dictionary
    Dim evaluation_item_dict As Scripting.Dictionary
    Dim evaluation_item_id As Variant
    Dim evaluation_item_name As Variant
    Dim evaluation_item_format As Variant
    Dim evaluation_item_sort As Variant
    Dim evaluation_item_summarize As Variant
    Dim ws As Worksheet
    Dim last_row As Long
    
    Set evaluation_item_dict = New Scripting.Dictionary
    Set ws = wb.Worksheets("評鑑指標")
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To last_row
        evaluation_item_id = ws.Cells(i, 1).Value
        evaluation_item_name = ws.Cells(i, 2).Value
        evaluation_item_format = ws.Cells(i, 3).Value
        evaluation_item_sort = ws.Cells(i, 4).Value
        evaluation_item_summarize = ws.Cells(i, 5).Value
        
        Set evaluation_item = New Scripting.Dictionary
        evaluation_item.Add "id", evaluation_item_id
        evaluation_item.Add "format", evaluation_item_format
        evaluation_item.Add "sort", evaluation_item_sort
        evaluation_item.Add "summarize", evaluation_item_summarize
        
        evaluation_item_dict.Add evaluation_item_name, evaluation_item
    Next i
    
    Set create_evaluation_item_dict = evaluation_item_dict
End Function

Private Sub test_create_evaluation_item_dict()
    Dim evaluation_item_dict As Scripting.Dictionary
    Dim evaluation_item_name As Variant
    Dim evaluation_item As Variant
    Dim wb As Workbook

    Set wb = Workbooks.Open(ThisWorkbook.path & "/B 參數.xlsx")
    Set evaluation_item_dict = create_evaluation_item_dict(wb)
    
    For Each evaluation_item_name In evaluation_item_dict.Keys
        Set evaluation_item = evaluation_item_dict(evaluation_item_name)
        MsgBox evaluation_item("id") & " " & evaluation_item_name & " " & evaluation_item("format") & " " & evaluation_item("sort") & " " & evaluation_item("summarize")
    Next evaluation_item_name
End Sub