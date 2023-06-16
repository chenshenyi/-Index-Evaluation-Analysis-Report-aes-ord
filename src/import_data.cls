' filename: import_data.cls

' This module use the following dictionary:
' college_department_dict
'   key: college name
'   value: [{id, name, abbr}] (list of department dictionary)
' evaluation_item_dict
'   key: evaluation item name
'   value: {id, format, sort, summarize}
'       id: String
'       format: "��Ƽƭ�" | "�ƭ�" | "�ʤ���"
'       sort: "���W" | "����"
'       summarize: "����" | "�[�`"

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
    
    ' Create the dictionary of college and evaluation item by "B �Ѽ�.xlsx"
    Set argument_wb = Workbooks.Open(ThisWorkbook.path & "/B �Ѽ�.xlsx")
    Set college_department_dict = create_college_department_dict(argument_wb)
    Set evaluation_item_dict = create_evaluation_item_dict(argument_wb)
    argument_wb.Close

End Function

