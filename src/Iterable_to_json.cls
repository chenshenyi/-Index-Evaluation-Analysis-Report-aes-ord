' filename: Iterable_to_json.cls
' This file print the content of a dictionary or a collection in json format

Function Iterable_to_json(var, Optional indent As Integer = 0)
    If TypeOf iterable Is Scripting.Dictionary Then
        Iterable_to_json = print_dict(var, indent)
    ElseIf TypeOf iterable Is Collection Then
        Iterable_to_json = print_collection(var, indent)
    Else
        ' if var is not a dictionary or a collection, return the string representation of the var
        indent_spaces = Space(indent)
        Iterable_to_json = indent_spaces & """" & var & """"
    End If
End Function

Function print_dict(dict As Scripting.Dictionary, indent As Integer)
    indent_spaces = Space(indent)
    print_dict = indent_spaces & "{"
    indent_spaces = Space(indent + 4)
    For Each key In dict
        print_dict = print_dict & vbCrLf & indent_spaces & """" & key & """: " & Iterable_to_json(dict(key), indent + 4) & ","
    Next
    print_dict = Left(print_dict, Len(print_dict) - 1) & vbCrLf & indent_spaces & "}"
End Function

Function print_collection(clcn As Collection, indent As Integer)
    indent_spaces = Space(indent)
    print_collection = indent_spaces & "["
    indent_spaces = Space(indent + 4)
    For Each item In clcn
        print_collection = print_collection & vbCrLf & indent_spaces & Iterable_to_json(item, indent + 4) & ","
    Next
    print_collection = Left(print_collection, Len(print_collection) - 1) & vbCrLf & indent_spaces & "]"
End Function

Private Sub test_iterable_to_json()
    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary
    dict.Add "key1", "value1"
    dict.Add "key2", "value2"
    dict.Add "key3", "value3"
    dict.Add "key4", "value4"
    dict.Add "key5", "value5"
    dict.Add "key6", "value6"
    dict.Add "key7", "value7"
    dict.Add "key8", "value8"
    dict.Add "key9", "value9"
    dict.Add "key10", "value10"
    
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

    Debug.Print Iterable_to_json(dict)
    Debug.Print Iterable_to_json(clcn)
    Debug.Print Iterable_to_json(multilayer_dict)
End Sub