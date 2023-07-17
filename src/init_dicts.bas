
Attribute VB_Name = "init_dicts"
' 2023-07-10 11:59:55


Global college_department_dict As Scripting.Dictionary
Global evaluation_item_dict As Scripting.Dictionary

Function argument_init()

    Dim argument_wb As Workbook
    If college_department_dict Is Nothing Then
        Set argument_wb = Workbooks.Open(ThisWorkbook.Path & "\B 把计.xlsx")
        Set college_department_dict = college_department_dict_init(argument_wb)
        argument_wb.Close False
    End If

    If evaluation_item_dict Is Nothing Then
        Set argument_wb = Workbooks.Open(ThisWorkbook.Path & "\B 把计.xlsx")
        Set evaluation_item_dict = evaluation_item_dict_init(argument_wb)
        argument_wb.Close False
    End If

End Function