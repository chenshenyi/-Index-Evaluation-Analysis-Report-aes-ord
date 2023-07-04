Attribute VB_Name = "content_control"


Function create_content_control_by_place_hold(wdDoc As Object, ByVal placehold As String)

    Dim wdRng As Object
    Set wdRng = wdDoc.Range
    wdRng.Find.Execute FindText:=placehold, ReplaceWith:="", Replace:=wdReplaceAll
    wdRng.Select
    wdDoc.ContentControls.Add wdContentControlText, wdRng
    wdRng.ContentControls(1).Range.Text = placehold
    wdRng.ContentControls(1).Title = placehold

End Function

Private Sub test_create_content_control_by_place_hold()

    Dim wdApp As Object
    Dim wdDoc As Object

    from_path = ThisWorkbook.path & "\input\create_content_control_by_place_hold.docx"
    to_path = ThisWorkbook.path & "\output\create_content_control_by_place_hold.docx"
    copy_file from_path, to_path

    Set wdApp = CreateObject("Word.Application")
    Set wdDoc = wdApp.Documents.Open(to_path)

    wdApp.Visible = True

    Call create_content_control_by_place_hold(wdDoc, "a-1")

End Sub

Function paste_clipboard_to_contentcontrol(wdDoc As Object, cc As Object)

    Dim wdRng As Object
    Set wdRng = cc.Range
    wdRng.Paste

End Function

Private Sub test_paste_clipboard_to_contentcontrol()

    Dim wdApp As Object
    Dim wdDoc As Object
    Dim cc As Object

    from_path = ThisWorkbook.path & "\input\paste_clipboard_to_contentcontrol.docx"
    to_path = ThisWorkbook.path & "\output\paste_clipboard_to_contentcontrol.docx"
    copy_file from_path, to_path

    Set wdApp = CreateObject("Word.Application")
    Set wdDoc = wdApp.Documents.Open(to_path)

    wdApp.Visible = True

    For Each cc In wdDoc.ContentControls
        If cc.Title = "a-1" Then
            Call paste_clipboard_to_contentcontrol(wdDoc, cc)
        End If
    Next cc

End Sub

Function paste_text_to_contentcontrol(wdDoc As Object, cc As Object, ByVal text As String)

    Dim wdRng As Object
    Set wdRng = cc.Range
    wdRng.Text = text

End Function

Private Sub test_paste_text_to_contentcontrol()

    Dim wdApp As Object
    Dim wdDoc As Object
    Dim cc As Object

    from_path = ThisWorkbook.path & "\input\paste_text_to_contentcontrol.docx"
    to_path = ThisWorkbook.path & "\output\paste_text_to_contentcontrol.docx"
    copy_file from_path, to_path

    Set wdApp = CreateObject("Word.Application")
    Set wdDoc = wdApp.Documents.Open(to_path)

    wdApp.Visible = True

    For Each cc In wdDoc.ContentControls
        If cc.Title = "a-1" Then
            Call paste_text_to_contentcontrol(wdDoc, cc, "test")
        End If
    Next cc

End Sub

Function clear_clipboard()
    Dim data As New DataObject
    Set data = New DataObject
    data.SetText ""
    data.PutInClipboard
End Function