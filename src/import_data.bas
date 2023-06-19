Attribute VB_Name = "import_data"

Function import_data(college_list As Collection, evaluation_item_list As Collection)
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Dim argument_wb As Workbook
    Dim college_department_dict As Scripting.Dictionary
    Dim evaluation_items_value_dict As Scripting.Dictionary
    
    ' Create the dictionary of college and evaluation item by "B �Ѽ�.xlsx"
    Set argument_wb = Workbooks.Open(ThisWorkbook.path & "/B �Ѽ�.xlsx")
    Set college_department_dict = college_department_dict_init(argument_wb)
    Set evaluation_items_value_dict = evaluation_items_value_dict_init(argument_wb, evaluation_item_list)
    argument_wb.Close

    For Each college in college_list
        ' Create the dictionary of evaluation item by "A ��Ų����.xlsx"
        Set evaluation_item_dict = evaluation_item_dict_init(college, evaluation_item_list)
        
        ' Import the data from "C ��Ų���.xlsx"
        Set evaluation_data_wb = Workbooks.Open(ThisWorkbook.path & "/C ��Ų���.xlsx")
        import_data_from_evaluation_data_wb college, evaluation_data_wb, college_department_dict, evaluation_item_dict, evaluation_items_value_dict
        evaluation_data_wb.Close

End Function

