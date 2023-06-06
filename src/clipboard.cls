' This code has following functions:
' 1. paste clipboard to placehold
' 2. paste clipboard to content control
' 3. create content control


Function paste_clipboard_to_placehold(wdDoc As Object, placehold As String)

Dim wdRng As Object

Set wdRng = wdDoc.Range
wdRng.Find.Execute FindText:=placehold, ReplaceWith:="", Replace:=wdReplaceAll
wdRng.Paste

End Function

Function paste_clipboard_to_contentcontrol(wdDoc As Object, cc As Object)

Dim wdRng As Object
Set wdRng = cc.Range
wdRng.Paste

End Function

Function create_content_control(wdDoc As Object, placehold As String)

Dim wdRng As Object
Set wdRng = wdDoc.Range
wdRng.Find.Execute FindText:=placehold, ReplaceWith:="", Replace:=wdReplaceAll
wdRng.Select
wdDoc.ContentControls.Add wdContentControlText, wdRng
wdRng.ContentControls(1).Range.Text = placehold
wdRng.ContentControls(1).Title = placehold

End Function


Sub paste_clipboard_to_placehold_test()

Dim wdApp As Object
Dim wdDoc As Object

Set wdApp = CreateObject("Word.Application")
Set wdDoc = wdApp.Documents.Open(ThisWorkbook.path & "\clipboard-paste-test.docx")

wdApp.Visible = True

Call paste_clipboard_to_placehold(wdDoc, "a-1")

End Sub

Sub paste_clipboard_to_contentcontrol_test()

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

Sub create_content_control_test()

Dim wdApp As Object
Dim wdDoc As Object

Set wdApp = CreateObject("Word.Application")
Set wdDoc = wdApp.Documents.Open(ThisWorkbook.path & "\clipboard-paste-test.docx")

wdApp.Visible = True

Call create_content_control(wdDoc, "a-2")

End Sub