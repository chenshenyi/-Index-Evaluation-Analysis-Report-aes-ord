Attribute VB_Name = "initialize_workbook"
' 2023-07-11 17:15:38


' Passed all tests
Sub initialize_all_workbooks()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Dim argument_wb As Workbook
    Dim wb As Workbook
    Dim sht As Worksheet
    Dim year As Integer
    Dim worksheet_name As String
    Dim department_number As Integer
    
    argument_init

    year = get_year()

    ' Initialize the template worksheet
    ThisWorkbook.worksheets("template").Range("D1").Value = year & "年"
    ThisWorkbook.worksheets("template").Range("E1").Value = year-1 & "年"
    ThisWorkbook.worksheets("template").Range("F1").Value = year-2 & "年"

    
    ' Create the workbooks by college names
    create_workbooks_by_college_names college_department_dict.Keys


    For Each college In college_department_dict.Keys
        Set wb = Workbooks.Open(college_excel_path(college))
        Call initialize_summary_worksheet(wb, college_department_dict(college))
        
        ' Initialize the worksheets by evaluation items
        For Each evaluation_item In evaluation_item_dict.Keys

            worksheet_name = evaluation_item_dict(evaluation_item)("id") & " " & evaluation_item
            Call initialize_worksheet(wb, worksheet_name, thisworkbook.worksheets("template"))

            department_number = college_department_dict(college).count - 1 ' First department is the college itself

            ' Insert new rows by the number of departments - 1
            Call insert_new_rows(wb, worksheet_name, department_number - 1)

            ' Write the department names
            Call write_department_names(wb, worksheet_name, college_department_dict(college), evaluation_item_dict(evaluation_item)("summarize"))

        Next evaluation_item      
        wb.Save
        wb.Close
    Next college

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

' =========================================== 統一初始化 ===========================================

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
    Application.ScreenUpdating = False
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
    Application.ScreenUpdating = True
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
' The data is stored in the worksheet "參數"!B1 in "B 參數.xlsx"
' Parameter:
'   wb: the workbook containing the worksheet "參數"
Function get_year() As Integer
    Dim year As Integer
    
    year = ThisWorkbook.Worksheets("主控台").Cells(1, 2).Value
    
    get_year = year
End Function

' ======================================== 編輯各院內容 ========================================

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
'   number_format: the number format of the cells, the option is "整數數值", "數值" and "百分比"

' * Write the department information to worksheet
' In the row 2:
' Collumn A is id & " " & name & "（院加總 / 院均值）"
' Collumn B is "院加總" or "院均值" according to summarize
' Write the information starting from row 3
' Collumn A is the id & " " & name
' Collumn B is the abbr
' Parameter:
'   wb: the workbook to be modified
'   worksheet_name: the name of the worksheet to be modified
'   department_list: the list of the department information, the format is {id, name, abbr}
'   summarize: the option is "加總" or "均值"
Function write_department_names(wb As Workbook, ByVal worksheet_name As String, department_list As Collection, ByVal summarize As String)
    Dim id As String
    Dim name As String
    Dim abbr As String
    Dim sht As Worksheet

    Set sht = wb.Worksheets(worksheet_name)

    surfix = ""
    ' write row 2
    If department_list(1)("name") = "政治大學" Then
        sht.Cells(2, 1) = department_list(1)("id") & " " & department_list(1)("name") & "（校加總 / 校均值）"
        sht.Cells(2, 2) = "校" & summarize
        surfix = "（院加總 / 院均值）"
    Else
        sht.Cells(2, 1) = department_list(1)("id") & " " & department_list(1)("name") & "（院加總 / 院均值）"
        sht.Cells(2, 2) = "院" & summarize
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
    
    Set argument_wb = Workbooks.Open(ThisWorkbook.path & "/B 參數.xlsx")
    Set college_department_dict = college_department_dict_init(argument_wb)
    argument_wb.Close
    
    Set department_list = college_department_dict("理學院") 
    
    Call write_department_names(ThisWorkbook, "test1", department_list, "加總")
End Sub

' ========================================== 小結表格 ==========================================

' Initialize the worksheet "小結"
' The first row is the abbreviation of the colleges
' The first column is the name of the evaluation items
' Parameter:
'   wb: the workbook to be modified
'   departments: the list of the department information, the format is {id, name, abbr, fullname}
Function initialize_summary_worksheet(wb As workbook, departments As Collection)
    
    If Not worksheet_exists(wb, "小結") Then
        wb.Worksheets("工作表1").name = "小結"
    Else
        wb.Worksheets("小結").Cells.Clear
    End If

    argument_init

    c = 2
    For Each department In departments
        wb.Worksheets("小結").Cells(1, c) = department("fullname")
        wb.Worksheets("小結").Cells(2, c) = department("abbr")
        ' set the column width
        wb.Worksheets("小結").Columns(c).ColumnWidth = 7
        c = c + 1
    Next department

    r = 3
    privious_group = -1
    current_group = -1
    For Each evaluation_item In evaluation_item_dict.keys
        current_group = evaluation_item_dict(evaluation_item)("group")
        If current_group <> privious_group And privious_group <> -1 Then
            wb.Worksheets("小結").Cells(r, 1) = "平均"
            wb.Worksheets("小結").Cells(r, 2) = "平均 " & privious_group
            r = r + 1
        End If
        privious_group = current_group

        worksheet_name = evaluation_item_dict(evaluation_item)("id") & " " & evaluation_item
        address = "'" & worksheet_name & "'" & "!A1"
        ' add an navigation link to the worksheet
        wb.Worksheets("小結").Cells(r, 1) = evaluation_item_dict(evaluation_item)("id")
        wb.Worksheets("小結").Hyperlinks.Add Anchor:=wb.Worksheets("小結").Cells(r, 2), Address:="", SubAddress:= address, TextToDisplay:= evaluation_item
        r = r + 1
    Next evaluation_item

    wb.Worksheets("小結").Cells(r, 1) = "平均"
    wb.Worksheets("小結").Cells(r, 2) = "小結 " & current_group


    ' auto fit the width of the first 2 column
    wb.Worksheets("小結").Columns(1).AutoFit
    wb.Worksheets("小結").Columns(2).AutoFit

    ' Set the font to "標楷體"
    wb.Worksheets("小結").Cells.Font.Name = "標楷體"
End Function

Private Sub test_initialize_summary_worksheet()
    Dim argument_wb As Workbook
    Dim college_department_dict As Scripting.Dictionary
    Dim evaluation_item_dict As Scripting.Dictionary
    
    Set argument_wb = Workbooks.Open(ThisWorkbook.path & "/B 參數.xlsx")
    Set college_department_dict = college_department_dict_init(argument_wb)
    Set evaluation_item_dict = evaluation_item_dict_init(argument_wb)
    argument_wb.Close
    
    Call initialize_summary_worksheet(ThisWorkbook, college_department_dict("理學院"), evaluation_item_dict)
End Sub
