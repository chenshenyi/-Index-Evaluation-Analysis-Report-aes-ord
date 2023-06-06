' This code has following functions:
' 1. Create workbooks by college names

Function create_workbooks_by_college_names(college_name_list As Variant, target_folder As String)

    Dim output_workbook As Workbook
    Dim college_name As Variant
    
    For Each college_name In college_name_list
        output_workbook_name = target_folder & college_name & ".xlsx"
        
        ' If the output workbook already exists, then go to the next college name
        If Dir(output_workbook_name) <> "" Then
            GoTo NextCollegeName
        End If

        Set output_workbook = Workbooks.Add
        output_workbook.SaveAs output_workbook_name
        output_workbook.Close

NextCollegeName:
    Next college_name
End Function

Function initialize_worksheets(workbook_names As Variant, worksheet_names As Variant, worksheet_template As Worksheet)

    Dim workbook_name As Variant
    Dim worksheet_name As Variant
    Dim workbook As Workbook
    Dim worksheet As Worksheet
    
    For Each workbook_name In workbook_names
        Set workbook = Workbooks.Open(ThisWorkbook.path & "/" & workbook_name)
        
        For Each worksheet_name In worksheet_names
            ' Test if the worksheet exists
            ' If the worksheet exists, then delete
            If worksheet_exists(workbook, worksheet_name) Then
                workbook.Worksheets(worksheet_name).Delete
            End If

            ' Copy the template worksheet
            worksheet_template.Copy After:=workbook.Worksheets(workbook.Worksheets.Count)
            Set worksheet = workbook.Worksheets(workbook.Worksheets.Count)
            worksheet.Name = worksheet_name
        Next worksheet_name
        
        workbook.Save
        workbook.Close
    Next workbook_name
End Function

Function worksheet_exists(workbook As Workbook, worksheet_name As Variant) As Boolean
    Dim worksheet As Worksheet
    
    worksheet_exists = False
    
    For Each worksheet In workbook.Worksheets
        If worksheet.Name = worksheet_name Then
            worksheet_exists = True
            Exit Function
        End If
    Next worksheet
End Function

Sub test_create_workbooks_by_college_names()
    Dim college_name_list As Variant
    
    college_name_list = Array("College1", "College2", "College3")

    Call create_workbooks_by_college_names(college_name_list, ThisWorkbook.path & "\")
End Sub

Sub test_initialize_worksheets()
Application.DisplayAlerts = False

    Dim workbook_names As Variant
    Dim worksheet_names As Variant
    Dim worksheet_template As Worksheet
    
    workbook_names = Array("College1.xlsx", "College2.xlsx", "College3.xlsx")
    worksheet_names = Array("Sheet1", "Sheet2", "Sheet3")
    Set worksheet_template = ThisWorkbook.Worksheets("Template")
    
    Call initialize_worksheets(workbook_names, worksheet_names, worksheet_template)

Application.DisplayAlerts = True
End Sub

Sub test_worksheet_exists()
    Dim workbook As Workbook
    Dim worksheet As Worksheet
    
    Set workbook = Workbooks.Open(ThisWorkbook.path & "/College1.xlsx")
    Set worksheet = workbook.Worksheets("Sheet1")
    
    MsgBox worksheet_exists(workbook, "Sheet1")
    MsgBox worksheet_exists(workbook, "Sheet2")
End Sub