Function merge_worksheet_in_multiple_workbooks(workbook_path_list As Variant, _
                                                worksheet_name As String, _
                                                output_workbook As Workbook, _
                                                output_worksheet_name_list As Variant, _
                                                Optional output_worksheet_name_prefix As String = "", _
                                                Optional output_worksheet_name_suffix As String = "")
    Dim output_worksheet As Worksheet
    Dim output_worksheet_name As String
    
    For i = 0 To UBound(workbook_path_list)
        Set input_workbook = Workbooks.Open(workbook_path_list(i))
        Set input_worksheet = input_workbook.Worksheets(worksheet_name)
        
        output_worksheet_name = output_worksheet_name_prefix & input_workbook.Name & output_worksheet_name_suffix
        
        ' If the output worksheet name is already in the output workbook, then delete it
        If Not IsError(Application.Match(output_worksheet_name, output_worksheet_name_list, 0)) Then
            output_workbook.Worksheets(output_worksheet_name).Delete
        End If

        ' Copy the input worksheet to the end of output workbook
        input_worksheet.Copy After:=output_workbook.Worksheets(output_workbook.Worksheets.Count)
        Set output_worksheet = output_workbook.Worksheets(output_workbook.Worksheets.Count)
        output_worksheet.Name = output_worksheet_name

        input_workbook.Close
    Next i
End Function

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



Sub test_merge_worksheet_in_multiple_workbooks()
    Dim workbook_path_list As Variant
    Dim output_workbook As Workbook
    Dim output_worksheet_name_list As Variant
    
    Application.DisplayAlerts = False

    workbook_path_list = Array(ThisWorkbook.path & "\Book1.xlsx", _
                               ThisWorkbook.path & "\Book2.xlsx", _
                               ThisWorkbook.path & "\Book3.xlsx")
    Set output_workbook = Workbooks.Open(ThisWorkbook.path & "\origin-data-rearrangement.xlsx")
    output_worksheet_name_list = Array("Book1", "Book2", "Book3")
    
    Call merge_worksheet_in_multiple_workbooks(workbook_path_list, _
                                               "sheet1", _
                                               output_workbook, _
                                               output_worksheet_name_list, _
                                               "Book", _
                                               "")
    Application.DisplayAlerts = True
End Sub

Sub test_create_workbooks_by_college_names()
    Dim college_name_list As Variant
    
    college_name_list = Array("College1", "College2", "College3")

    Call create_workbooks_by_college_names(college_name_list, ThisWorkbook.path & "\")
End Sub