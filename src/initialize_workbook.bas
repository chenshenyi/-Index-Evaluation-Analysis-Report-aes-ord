Attribute VB_Name = "initialize_workbook"

' This module use the following dictionary:
' college_department_dict
'   key: college name
'   value: [{id, name, abbr}] (list of department dictionary)
' evaluation_item_dict
'   key: evaluation item name
'   value: {id, format, sortBy, summarize}
'       id: String
'       format: "��Ƽƭ�" | "�ƭ�" | "�ʤ���"
'       sortBy: "���W" | "����"
'       summarize: "����" | "�[�`"


Sub initialize_all_workbooks()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Dim argument_wb As Workbook
    Dim college_department_dict As Scripting.Dictionary
    Dim evaluation_item_dict As Scripting.Dictionary
    Dim wb As Workbook
    Dim sht As Worksheet
    Dim year As Integer
    Dim worksheet_name As String
    Dim department_number As Integer
    
    ' Create the dictionary of college and evaluation item by "B �Ѽ�.xlsx"
    Set argument_wb = Workbooks.Open(ThisWorkbook.path & "/B �Ѽ�.xlsx")
    Set college_department_dict = college_department_dict_init(argument_wb)
    Set evaluation_item_dict = evaluation_item_dict_init(argument_wb)
    argument_wb.Close
    
    year = get_year()

    ' Initialize the template worksheet
    ThisWorkbook.worksheets("template").Range("D1").Value = year & "�~"
    ThisWorkbook.worksheets("template").Range("E1").Value = year-1 & "�~"
    ThisWorkbook.worksheets("template").Range("F1").Value = year-2 & "�~"

    ' Create the workbooks by college names
    Call create_workbooks_by_college_names(college_department_dict.keys)

    For Each college In college_department_dict.keys
        Set wb = Workbooks.Open(college_excel_path(college))
        Call initialize_summary_worksheet(wb, college_department_dict(college), evaluation_item_dict)
        
        ' Initialize the worksheets by evaluation items
        For Each evaluation_item In evaluation_item_dict.keys

            worksheet_name = evaluation_item_dict(evaluation_item)("id") & " " & evaluation_item
            Call initialize_worksheet(wb, worksheet_name, thisworkbook.worksheets("template"))

            department_number = college_department_dict(college).count - 1 ' First department is the college itself

            ' Insert new rows by the number of departments - 1
            Call insert_new_rows(wb, worksheet_name, department_number - 1)

            ' Set the number format
            Call set_number_format(wb, worksheet_name, evaluation_item_dict(evaluation_item)("format"))

            ' Write the department names
            Call write_department_names(wb, worksheet_name, college_department_dict(college), evaluation_item_dict(evaluation_item)("summarize"))

        Next evaluation_item      
        wb.Save
        wb.Close
    Next college

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

' =========================================== �Τ@��l�� ===========================================

' * Create workbooks by college names
' The workbooks will be created in the target folder
' Parameter:
' `  college_name_list: the list of college names
'   target_folder: the folder to store the workbooks
Function create_workbooks_by_college_names(college_name_list As Variant)

    Dim output_workbook As Workbook
    Dim college_name As Variant
    
    For Each college_name In college_name_list
        output_workbook_name = college_excel_path(college_name)
        
        ' If the output wb already exists, then go to the next college name
        If Dir(output_workbook_name) <> "" Then
            GoTo NextCollegeName
        End If

        Set output_workbook = Workbooks.Add
        output_workbook.SaveAs output_workbook_name
        output_workbook.Close

NextCollegeName:
    Next college_name
End Function

Private Sub test_create_workbooks_by_college_names()
    Dim college_name_list As Variant
    
    college_name_list = Array("College1", "College2", "College3")

    Call create_workbooks_by_college_names(college_name_list, ThisWorkbook.path & "\")
End Sub

' * Initialize worksheets by Worksheet Template
' The worksheets will be created in the target workbook
' Parameter:
'   wb: the target workbook
'   worksheet_name: the name of the worksheet to be created
'   worksheet_template: the template worksheet
Function initialize_worksheet(wb As Workbook, ByVal worksheet_name As String, worksheet_template As Worksheet)

    Dim sht As Worksheet
        
    ' Test if the sht exists
    ' If the sht exists, then delete
    If worksheet_exists(wb, worksheet_name) Then
        wb.Worksheets(worksheet_name).Delete
    End If

    ' Copy the template sht
    worksheet_template.Copy After:=wb.Worksheets(wb.Worksheets.Count)
    Set sht = wb.Worksheets(wb.Worksheets.Count)
    sht.Name = worksheet_name

End Function

Private Sub test_initialize_worksheets()
    Application.DisplayAlerts = False

    Dim workbook_names As Variant
    Dim worksheet_names As Variant
    Dim worksheet_template As Worksheet
    
    workbook_names = Array("College1", "College2", "College3")
    worksheet_names = Array("Sheet1", "Sheet2", "Sheet3")

    Set worksheet_template = ThisWorkbook.Worksheets("Template")

    For Each workbook_name In workbook_names
        Set wb = Workbooks.Open(ThisWorkbook.path & "\" & workbook_name & ".xlsx")
        Call initialize_worksheets(wb, worksheet_names, worksheet_template)
        wb.Save
        wb.Close
    Next workbook_name

    Application.DisplayAlerts = True
End Sub

' * Test if the worksheet exists in the workbook
' Parameter:
'   wb: the target workbook
'   worksheet_name: the name of the worksheet to be tested
Function worksheet_exists(wb As Workbook, ByVal worksheet_name As String) As Boolean
    Dim sht As Worksheet
    
    worksheet_exists = False
    
    For Each sht In wb.Worksheets
        If sht.Name = worksheet_name Then
            worksheet_exists = True
            Exit Function
        End If
    Next sht
End Function

Private Sub test_worksheet_exists()
    Dim wb As Workbook
    Dim sht As Worksheet
    
    Set wb = Workbooks.Open(ThisWorkbook.path & "/College1.xlsx")
    Set sht = wb.Worksheets("Sheet1")
    
    MsgBox worksheet_exists(wb, "Sheet1")
    MsgBox worksheet_exists(wb, "Sheet2")
End Sub

' * Get the year of this report
' The data is stored in the worksheet "�Ѽ�"!B1 in "B �Ѽ�.xlsx"
' Parameter:
'   wb: the workbook containing the worksheet "�Ѽ�"
Function get_year() As Integer
    Dim year As Integer
    
    year = ThisWorkbook.Worksheets("�D���x").Cells(1, 2).Value
    
    get_year = year
End Function

' ======================================== �s��U�|���e ========================================

' * Insert new rows according to the number of departments
' This function will insert empty rows between row 2 and row 3
' Parameter: 
'   wb: the workbook to be modified
'   worksheet_name: the name of the worksheet to be modified
'   number_of_departments: the number of departments
Function insert_new_rows(wb As Workbook, ByVal worksheet_name As String, ByVal number_of_departments As Integer)
    Dim sht As Worksheet
    Dim last_row As Integer
    
    Set sht = wb.Worksheets(worksheet_name)
    sht.Rows(3).Resize(number_of_departments).Insert Shift:=xlDown

End Function

Private Sub test_insert_new_rows()
    Dim wb As Workbook
    
    Set wb = ThisWorkbook
    
    Call insert_new_rows(wb, "test1", 3)
End Sub

' * Set number format of the cells
' The range of the cells is from column C to column F
' The first row of the range is row 2
' The last row of the range is the last row of the worksheet
' Parameter:
'   wb: the workbook to be modified
'   worksheet_name: the name of the worksheet to be modified
'   number_format: the number format of the cells, the option is "��Ƽƭ�", "�ƭ�" and "�ʤ���"
' * When the cells are empty, the number format should be "�X"
Function set_number_format(wb As Workbook, ByVal worksheet_name As String, ByVal number_format As String)
    Dim sht As Worksheet
    Dim last_row As Integer
    Dim rng As Range

    Set sht = wb.Worksheets(worksheet_name)
    last_row = sht.Cells(sht.Rows.Count, 3).End(xlUp).Row
    Set rng = sht.Range("C2:F" & last_row)
   
    Select Case number_format
        Case "��Ƽƭ�"
            rng.NumberFormat = "#,##0;""�X"";""�X"""
        Case "�ƭ�"
            rng.NumberFormat = "#,##0.00;""�X"";""�X"""
        Case "�ʤ���"
            rng.NumberFormat = "0.00%;""�X"";""�X"""
    End Select
End Function

Private Sub test_set_number_format()
    Dim wb As Workbook
    
    Set wb = ThisWorkbook
    
    Call set_number_format(wb, "test1", "��Ƽƭ�")
End Sub

' * Write the department information to worksheet
' In the row 2:
' Collumn A is id & " " & name & "�]�|�[�` / �|���ȡ^"
' Collumn B is "�|�[�`" or "�|����" according to summarize
' Write the information starting from row 3
' Collumn A is the id & " " & name
' Collumn B is the abbr
' Parameter:
'   wb: the workbook to be modified
'   worksheet_name: the name of the worksheet to be modified
'   department_list: the list of the department information, the format is {id, name, abbr}
'   summarize: the option is "�[�`" or "����"
Function write_department_names(wb As Workbook, ByVal worksheet_name As String, department_list As Collection, ByVal summarize As String)
    Dim id As String
    Dim name As String
    Dim abbr As String
    Dim sht As Worksheet

    Set sht = wb.Worksheets(worksheet_name)

    surfix = ""
    ' write row 2
    If department_list(1)("name") = "�F�v�j��" Then
        sht.Cells(2, 1) = department_list(1)("id") & " " & department_list(1)("name") & "�]�ե[�` / �է��ȡ^"
        sht.Cells(2, 2) = "��" & summarize
        surfix = "�]�|�[�` / �|���ȡ^"
    Else
        sht.Cells(2, 1) = department_list(1)("id") & " " & department_list(1)("name") & "�]�|�[�` / �|���ȡ^"
        sht.Cells(2, 2) = "�|" & summarize
    End If

    r = 3
    ' write row 3 to the last row
    ' skip the first department
    ' because the first department is the college
    For Each department In department_list
        If department("name") = department_list(1)("name") Then
            GoTo NextDepartment
        End If
        
        sht.Cells(r, 1) = department("id") & " " & department("name") & surfix
        sht.Cells(r, 2) = department("abbr")
        r = r + 1
NextDepartment:
    Next department
End Function

Private Sub test_write_department_names()

    Dim argument_wb As Workbook
    Dim college_department_dict As Scripting.Dictionary
    Dim department_list As Collection
    
    Set argument_wb = Workbooks.Open(ThisWorkbook.path & "/B �Ѽ�.xlsx")
    Set college_department_dict = college_department_dict_init(argument_wb)
    argument_wb.Close
    
    Set department_list = college_department_dict("�z�ǰ|") 
    
    Call write_department_names(ThisWorkbook, "test1", department_list, "�[�`")
End Sub

' ========================================== �p����� ==========================================

' Initialize the worksheet "�p��"
' The first row is the abbreviation of the colleges
' The first column is the name of the evaluation items
' Parameter:
'   wb: the workbook to be modified
'   college_list: the list of the college information, the format is {id, name, abbr}
'   evaluation_item_dict: the dictionary of the evaluation item, the format is {name: {id, format, sortBy, summarize}}
Function initialize_summary_worksheet(wb As workbook, college_list As Collection, evaluation_item_dict As Scripting.Dictionary)
    
    If Not worksheet_exists(wb, "�p��") Then
        wb.Worksheets("�u�@��1").name = "�p��"
    Else
        wb.Worksheets("�p��").Cells.Clear
    End If

    c = 1
    For Each college In college_list
        wb.Worksheets("�p��").Cells(1, c) = college("abbr")
        c = c + 1
    Next college

    r = 2
    For Each evaluation_item In evaluation_item_dict.keys
        worksheet_name = evaluation_item_dict(evaluation_item)("id") & " " & evaluation_item
        address = "'" & worksheet_name & "'" & "!A1"
        ' add an navigation link to the worksheet
        wb.Worksheets("�p��").Hyperlinks.Add Anchor:=wb.Worksheets("�p��").Cells(r, 1), Address:="", SubAddress:= address, TextToDisplay:=worksheet_name
        r = r + 1
    Next evaluation_item

    ' auto fit the width of the collumn
    wb.Worksheets("�p��").Columns.AutoFit

    ' Set the font to "�з���"
    wb.Worksheets("�p��").Cells.Font.Name = "�з���"
End Function

Private Sub test_initialize_summary_worksheet()
    Dim argument_wb As Workbook
    Dim college_department_dict As Scripting.Dictionary
    Dim evaluation_item_dict As Scripting.Dictionary
    
    Set argument_wb = Workbooks.Open(ThisWorkbook.path & "/B �Ѽ�.xlsx")
    Set college_department_dict = college_department_dict_init(argument_wb)
    Set evaluation_item_dict = evaluation_item_dict_init(argument_wb)
    argument_wb.Close
    
    Call initialize_summary_worksheet(ThisWorkbook, college_department_dict("�z�ǰ|"), evaluation_item_dict)
End Sub
