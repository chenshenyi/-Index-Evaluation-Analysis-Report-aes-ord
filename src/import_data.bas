Attribute VB_Name = "import_data"

Function import_data(college_list As Collection, evaluation_item_list As Collection)
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Dim argument_wb As Workbook
    Dim college_department_dict As Scripting.Dictionary
    Dim evaluation_items_value_dict As Scripting.Dictionary
    Dim destin_wb As Workbook
    
    ' Create the dictionary of college and evaluation item by "B 把计.xlsx"
    Set argument_wb = Workbooks.Open(ThisWorkbook.path & "/B 把计.xlsx")
    Set college_department_dict = college_department_dict_init(argument_wb)
    Set evaluation_items_value_dict = evaluation_items_value_dict_init(argument_wb, evaluation_item_list)
    argument_wb.Close

    For Each college In college_list
        Set destin_wb = Workbooks.Open(path_dir1(college))
        
        destin_wb.Close
    Next college
End Function


Function import_data_to_wb(wb as Workbook, ByVal college as String, evaluation_item_list as Collection, evaluation_items_value_dict as Scripting.Dictionary)

End Function

' Save the data to the worksheet
' Column A: department_name
' Column C: avg
' Column D: year3
' Column E: year2
' Column F: year1

' college_value_dict:
'   key: department_name
'   value: {avg, year3, year2, year1, rank}
Function import_data_to_ws(ws as Worksheet, ByVal college as String, evaluation_item_list as Collection, college_value_dict as Scripting.Dictionary)
    
End Function