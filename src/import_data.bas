Attribute VB_Name = "import_data"

Function import_data(college_list As Variant, evaluation_item_list As Variant)
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Dim argument_wb As Workbook
    Dim college_department_dict As Scripting.Dictionary
    Dim evaluation_item_dict As Scripting.Dictionary

    Dim wb As Workbook
    Dim sht As Worksheet
    Dim year As Integer
    Dim worksheet_name As String
    Dim college_number As Integer
    
    ' Create the dictionary of college and evaluation item by "B 把计.xlsx"
    Set argument_wb = Workbooks.Open(ThisWorkbook.path & "/B 把计.xlsx")
    Set college_department_dict = college_department_dict_init(argument_wb)
    Set evaluation_item_dict = evaluation_item_dict_init(argument_wb)
    argument_wb.Close

End Function

