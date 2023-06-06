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

Sub test_create_workbooks_by_college_names()
    Dim college_name_list As Variant
    
    college_name_list = Array("College1", "College2", "College3")

    Call create_workbooks_by_college_names(college_name_list, ThisWorkbook.path & "\")
End Sub