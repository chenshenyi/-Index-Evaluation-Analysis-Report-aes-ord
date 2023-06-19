Attribute VB_Name = "print_tool"

' This file print the content of a dictionary or a collection in json format

Function json_str(iterable As Variant, Optional indent As Integer = 0)
    If TypeOf iterable Is Scripting.Dictionary Then
        json_str = dict_json(iterable, indent)
    ElseIf TypeOf iterable Is Collection Then
        json_str = collection_json(iterable, indent)
    ElseIf IsNumeric(iterable) Then 
        json_str = iterable
    Else
        json_str = """" & iterable & """"
    End If
End Function

Function dict_json(dict As Variant, indent As Integer)
    dict_json = vbCrLf & Space(indent) & "{"
    For Each key In dict
        dict_json = dict_json & vbCrLf & Space(indent+2) & """" & key & """: " & json_str(dict(key), indent+2) & ","
    Next
    dict_json = Left(dict_json, Len(dict_json) - 1) & vbCrLf & Space(indent) & "}"
End Function

Function collection_json(clcn As Variant, indent As Integer)
    collection_json = vbCrLf & Space(indent) & "["
    For Each item In clcn
        collection_json = collection_json & vbCrLf & Space(indent+2) & json_str(item, indent+2) & ","
    Next
    collection_json = Left(collection_json, Len(collection_json) - 1) & vbCrLf & Space(indent) & "]"
End Function

Private Sub test_json_str()
    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary
    dict.Add "key1", "value1"
    dict.Add "key2", "value2"
    dict.Add "key3", "value3"
    dict.Add "key4", "value4"
    
    Dim clcn As Collection
    Set clcn = New Collection
    clcn.Add "item1"
    clcn.Add "item2"
    clcn.Add "item3"
    clcn.Add "item4"
    clcn.Add "item5"

    Dim multilayer_dict As Scripting.Dictionary
    Set multilayer_dict = New Scripting.Dictionary
    multilayer_dict.Add "key1", "value1"
    multilayer_dict.Add "key2", "value2"
    multilayer_dict.Add "key3", dict
    multilayer_dict.Add "key4", clcn

    Debug.Print json_str(dict)
    Debug.Print json_str(clcn)
    Debug.Print json_str(multilayer_dict)
End Sub

Function print_to_file(file_path As String, content As String)
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    Dim file As Scripting.TextStream
    Set file = fso.CreateTextFile(file_path, True)
    file.Write content
    file.Close
End Function

Private Sub test_print_to_file()
    print_to_file ThisWorkbook.Path & "\HelloWorld.txt", "Hello World!"
End Sub