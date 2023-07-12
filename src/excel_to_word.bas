Attribute VB_Name = "excel_to_word"
' 2023-07-11 17:20:24

Private wd As Object

Function word_Open(ByVal path As String) As Object

    If wd Is Nothing Then
        Set wd = CreateObject("Word.Application")
    End If

    Set word_Open = wd.Documents.Open(path)

End Function

Sub create_content_controls()

    Application.ScreenUpdating = False
    
    
    argument_init

    Dim colleges As Collection
    ' 偵測檔案是否存在
    Set colleges = New Collection
    For Each college In college_department_dict.Keys
        colleges.Add college
    Next college
    
    Dim files As Collection
    Set files = New Collection
    
    folder = ThisWorkbook.Path & "\2. 各院報告書模板"
    Set found = detect_files_in_folder( _
        folder, _
        colleges, _
        ".docx")

    Application.DisplayAlerts = False

    pattern1 = "$_[0-9.]@_chart_$"
    pattern2 = "$_[0-9.]@_text_$"

    For Each file In found
        Set doc = word_Open(folder & "\" & file & ".docx")
        ' if doc is protected, unprotect it
        If doc.ProtectionType <> wdNoProtection Then
            doc.Unprotect
        End If
        Set rng = doc.Range
        with rng.Find
            .Text = pattern1
            .MatchWildcards = True
            .Wrap = wdFindStop
            .Execute
            Do While .Found
                rng.Select
                title = rng.Text
                doc.ContentControls.Add wdContentControlText, rng
                rng.ContentControls(1).Title = title
                rng.ContentControls(1).Range.Text = ""
            Loop
        End With
        with rng.Find
            .Text = pattern2
            .MatchWildcards = True
            .Wrap = wdFindStop
            .Execute
            Do While .Found
                rng.Select
                title = rng.Text
                doc.ContentControls.Add wdContentControlRichText, rng
                rng.ContentControls(1).Title = title
                rng.ContentControls(1).Range.Text = ""
            Loop
        End With
        doc.Close SaveChanges:=True
    Next file

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub