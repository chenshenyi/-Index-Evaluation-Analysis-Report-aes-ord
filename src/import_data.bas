Attribute VB_Name = "import_data"

' Passed all test

' TODO: add filter to each worksheet


' ================================================= import data ==================

' * Import data from "0. 原始資料" to "1. 各院彙整資料"
Function import_data(college_list As Collection, evaluation_item_list As Collection)
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Dim argument_wb As Workbook
    Dim college_department_dict As Scripting.Dictionary
    Dim evaluation_items_value_dict As Scripting.Dictionary
    Dim destin_wb As Workbook
    
    ' Create the dictionary of college and evaluation item by "B 參數.xlsx"
    Set argument_wb = Workbooks.Open(ThisWorkbook.path & "/B 參數.xlsx")
    Set college_department_dict = college_department_dict_init(argument_wb)
    Set evaluation_items_value_dict = evaluation_items_value_dict_init(argument_wb, evaluation_item_list)
    argument_wb.Close

    For Each college In college_list
        Set destin_wb = Workbooks.Open(college_excel_path(college))
        college_id = college_department_dict(college)(1)("id") ' Caution: the college's information is in the first item of the dict
        
        import_data_to_wb destin_wb, college_id & " " & college, evaluation_item_list, evaluation_items_value_dict
        destin_wb.Close True
    Next college
End Function

' Passed test
Private Sub test_import_data()
    Dim college_list As Collection
    Set college_list = New Collection
    college_list.Add "政治大學"
    college_list.Add "文學院"
    college_list.Add "理學院"

    Dim evaluation_item_list As Collection
    Set evaluation_item_list = New Collection
    evaluation_item_list.Add "學士班繁星推薦入學錄取率"
    evaluation_item_list.Add "博士班招收國內重點大學畢業生比率"
    evaluation_item_list.Add "學士班獲獎助學金平均金額"

    import_data college_list, evaluation_item_list
End Sub


' * Import data
' Parameter:
'   wb: the workbook to import data
'   college_with_id: the college id and name
'   evaluation_item_list: the list of evaluation item
'   evaluation_items_value_dict: the dict initialized by evaluation_items_value_dict_init (資料總清單)
Function import_data_to_wb( wb as Workbook, _
                            ByVal college_with_id as String, _
                            evaluation_item_list as Collection, _
                            evaluation_items_value_dict as Scripting.Dictionary)
    Dim ws As Worksheet
    Dim shtname As String
    Dim evaluation_item As Variant
    Dim evaluatation_item_dict As Scripting.Dictionary
    Dim college_value_dict As Scripting.Dictionary

    ' 將資料匯入到各個工作表(各個指標評鑑項目)
    For Each evaluation_item In evaluation_item_list
        
        ' 該指標的資料
        Set evaluatation_item_dict = evaluation_items_value_dict(evaluation_item)

        ' 該指標該學院的資料
        Set college_value_dict = evaluatation_item_dict("evaluation_value_dict")(college_with_id)
        
        ' 該指標所在的工作表
        shtname = evaluatation_item_dict("id") & " " & evaluation_item
        Set ws = wb.Worksheets(shtname)

        ' 將資料匯入到工作表
        import_data_to_ws ws, college_value_dict
        
    Next evaluation_item
    
End Function

' Passed test
Private Sub test_import_data_to_wb()
    Dim argument_wb As Workbook
    Set argument_wb = Workbooks.Open(ThisWorkbook.path & "/B 參數.xlsx")

    Dim evaluation_item_list As Collection
    Set evaluation_item_list = New Collection
    evaluation_item_list.Add "學士班繁星推薦入學錄取率"
    evaluation_item_list.Add "博士班招收國內重點大學畢業生比率"
    
    Dim evaluation_items_value_dict As Scripting.Dictionary
    Set evaluation_items_value_dict = evaluation_items_value_dict_init(argument_wb, evaluation_item_list)
    argument_wb.Close

        Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = Workbooks.Open(ThisWorkbook.path & "\1. 各院彙整資料\理學院.xlsx")
    college_with_id = "700 理學院"
    import_data_to_wb wb, college_with_id, evaluation_item_list, evaluation_items_value_dict
End Sub


' * Save the data to the worksheet
' Column A: department_name
' Column C: avg
' Column D: year3
' Column E: year2
' Column F: year1
' Column G: rank
' Parameter:
'   ws: the worksheet to save data
'   college_value_dict: the dict of the college's data
'       key: department_name
'       value: {avg, year3, year2, year1, rank}
Function import_data_to_ws(ws as Worksheet, college_value_dict as Scripting.Dictionary)
    Dim department_name As Variant
    Dim department_value_dict As Scripting.Dictionary
    Dim row As Integer
    
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For row = 2 to last_row

        ' 跟據第一欄的學系名稱，取得該學系的資料
        department_name = ws.Cells(row, 1).Value
        Set department_value_dict = college_value_dict(department_name)
                
        ws.Cells(row, 3).Value = department_value_dict("avg")
        ws.Cells(row, 4).Value = department_value_dict("year3")
        ws.Cells(row, 5).Value = department_value_dict("year2")
        ws.Cells(row, 6).Value = department_value_dict("year1")
        ws.Cells(row, 7).Value = department_value_dict("rank")
    Next row
End Function

' Passed test
Private Sub test_import_data_to_ws()
    Dim argument_wb As Workbook
    Set argument_wb = Workbooks.Open(ThisWorkbook.path & "/B 參數.xlsx")

    Dim evaluation_item_list As Collection
    Set evaluation_item_list = New Collection
    evaluation_item_list.Add "學士班繁星推薦入學錄取率"
    
    Dim evaluation_items_value_dict As Scripting.Dictionary
    Set evaluation_items_value_dict = evaluation_items_value_dict_init(argument_wb, evaluation_item_list)
    argument_wb.Close

    Dim college_value_dict As Scripting.Dictionary
    Set college_value_dict = evaluation_items_value_dict("學士班繁星推薦入學錄取率")("evaluation_value_dict")("700 理學院")

    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = Workbooks.Open(ThisWorkbook.path & "\1. 各院彙整資料\理學院.xlsx")
    Set ws = wb.Worksheets("1.1.1.1 學士班繁星推薦入學錄取率")
    
    import_data_to_ws ws, college_value_dict
End Sub


' =============================================== Filter =====================================================