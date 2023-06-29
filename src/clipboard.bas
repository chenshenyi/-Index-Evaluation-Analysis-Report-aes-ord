Attribute VB_Name = "clipboard"

Function paste_clipboard_to_placehold(wdDoc As Object, placehold As String)

    Dim wdRng As Object

    Set wdRng = wdDoc.Range
    wdRng.Find.Execute FindText:=placehold, ReplaceWith:="", Replace:=wdReplaceAll
    wdRng.Paste

End Function

Private Sub test_paste_clipboard_to_placehold()

    Dim wdApp As Object
    Dim wdDoc As Object

    Set wdApp = CreateObject("Word.Application")
    Set wdDoc = wdApp.Documents.Open(ThisWorkbook.path & "\clipboard-paste-test.docx")

    wdApp.Visible = True

    Call paste_clipboard_to_placehold(wdDoc, "a-1")

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

    Set wdApp = CreateObject("Word.Application")
    Set wdDoc = wdApp.Documents.Open(ThisWorkbook.path & "\clipboard-paste-test.docx")

    wdApp.Visible = True

    For Each cc In wdDoc.ContentControls
        If cc.Title = "a-1" Then
            Call paste_clipboard_to_contentcontrol(wdDoc, cc)
        End If
    Next cc

End Sub


Function create_content_control_by_place_hold(wdDoc As Object, placehold As String)

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

    Set wdApp = CreateObject("Word.Application")
    Set wdDoc = wdApp.Documents.Open(ThisWorkbook.path & "\clipboard-paste-test.docx")

    wdApp.Visible = True

    Call create_content_control(wdDoc, "a-2")

End Sub