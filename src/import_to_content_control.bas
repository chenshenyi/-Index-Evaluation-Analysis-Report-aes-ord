Attribute VB_Name = "import_to_content_control"
' 2023-07-12 16:17:31

Private wd As Object

Function word_Open(ByVal path As String) As Object

    If wd Is Nothing Then
        Set wd = CreateObject("Word.Application")
        wd.Visible = True
    End If

    Set word_Open = wd.Documents.Open(path)

End Function

' =========================================== create content controls ===========================================

Sub create_all_content_controls()

    Application.ScreenUpdating = False
    argument_init
    Application.ScreenUpdating = True

    fromfolder = ThisWorkbook.Path & "\2. 各院報告書模板"
    tofolder = ThisWorkbook.Path & "\3. 各院報告書（輸入）"
    
    ' colleges = Collection(college_department_dict.Keys)
    Dim colleges As Collection
    Set colleges = New Collection
    For Each college In college_department_dict.Keys
        colleges.Add college
    Next college
    
    Dim found As Collection
    Set found = detect_files_in_folder(fromfolder, colleges, ".docx")

    ' copy found files to tofolder
    For Each file In found
        copy_file fromfolder & "\" & file & ".docx", tofolder & "\" & file & ".docx"
    Next file
    
    ' create content controls by pattern
    pattern1 = "$_[0-9.]@_chart_$"
    pattern2 = "$_[0-9.]@_text_$"

    Dim doc As Object
    For Each file In found
        Set doc = word_Open(tofolder & "\" & file & ".docx")
        create_content_controls doc, pattern1
        create_content_controls doc, pattern2
        doc.Close SaveChanges:=True
    Next file

    wd.Quit

End Sub

Function create_content_controls(wdDoc As Object, ByVal pattern As String)

    Dim wdRng As Object
    Set wdRng = wdDoc.Range
    found = True
    Do While found
        with wdRng.Find
            .Text = pattern
            .MatchWildcards = True
            .Wrap = wdFindStop
            .Execute
        End With

        found = wdRng.Find.Found
        If found Then
            wdRng.Select
            title = wdRng.Text
            wdDoc.ContentControls.Add wdContentControlText, wdRng
            wdRng.ContentControls(1).Title = title
            wdRng.ContentControls(1).Range.Text = ""
        End If
        Set wdRng = wdDoc.Range
    Loop

End Function

' Passed test
Private Sub test_create_content_controls()

    Dim wdApp As Object
    Dim wdDoc As Object
    
    from_path = ThisWorkbook.path & "\input\create_content_controls.docx"
    to_path = ThisWorkbook.path & "\output\create_content_controls.docx"
    copy_file from_path, to_path

    Set wdApp = CreateObject("Word.Application")
    Set wdDoc = wdApp.Documents.Open(to_path)

    wdApp.Visible = True

    pattern1 = "$_[0-9.]@_chart_$"
    pattern2 = "$_[0-9.]@_text_$"

    Call create_content_controls(wdDoc, pattern1)
    Call create_content_controls(wdDoc, pattern2)

End Sub

' =========================================== import data to content controls ===========================================

Sub import_all_data_to_content_controls()
    argument_init

    excelfolder = ThisWorkbook.Path & "\1. 各院彙整資料"
    fromfolder = ThisWorkbook.Path & "\3. 各院報告書（輸入）"
    tofolder = ThisWorkbook.Path & "\4. 各院報告書（輸出）"

    ' colleges = Collection(college_department_dict.Keys)
    Dim colleges As Collection
    Set colleges = New Collection
    For Each college In college_department_dict.Keys
        colleges.Add college
    Next college

    Dim found As Collection
    Set found = detect_files_in_folder(fromfolder, colleges, ".docx")

    ' copy found files to tofolder
    For Each file In found
        copy_file fromfolder & "\" & file & ".docx", tofolder & "\" & file & ".docx"
    Next file

    Dim doc As Object
    Dim wb As Workbook
    For Each file In found
        Set doc = word_Open(tofolder & "\" & file & ".docx")
        Set wb = Workbooks.Open(excelfolder & "\" & file & ".xlsx")
        
        import_data_to_content_controls doc, wb

        doc.Close SaveChanges:=True
        wb.Close SaveChanges:=False
    Next file

    wd.Quit

End Sub

Function import_data_to_content_controls(wdDoc As Object, xlWb As Workbook)
    argument_init

    For Each cc In wdDoc.ContentControls
        If cc.Title Like "$_[0-9.][0-9.]*_chart_$" Then
            id = Split(cc.Title, "_")(1)
            For Each sheet In xlWb.Sheets
                If sheet.Name Like id & "?*" Then
                    sheet.ChartObjects(1).Chart.ChartArea.Copy
                    cc.range.Paste
                    Exit For
                End If
            Next sheet
        ElseIf cc.Title Like "$_[0-9.][0-9.]*_text_$" Then
            id = Split(cc.Title, "_")(1)
            For Each sheet In xlWb.Sheets
                If sheet.Name Like id & "?*" Then
                    cc.range.Text = sheet.range("I1").Value
                    Exit For
                End If
            Next sheet
        End If
    Next cc

End Function

Private Sub test_import_data_to_content_controls()

    Dim xlWb As Workbook
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim cc As Object

    from_path = ThisWorkbook.path & "\input\import_data_to_content_controls.docx"
    to_path = ThisWorkbook.path & "\output\import_data_to_content_controls.docx"
    copy_file from_path, to_path

    Set wdDoc = word_Open(to_path)
    Set xlWb = Workbooks.Open(ThisWorkbook.path & "\1. 各院彙整資料\政治大學.xlsx")
    Call import_data_to_content_controls(wdDoc, xlWb)

End Sub
