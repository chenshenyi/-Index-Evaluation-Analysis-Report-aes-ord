Attribute VB_Name = "print_tool"
' 2023-07-10 12:00:57

' Passed all tests

' ================================= print =================================
' * 將字串輸出到檔案
Function print_to_file(ByVal file_path As String, content As String)
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    Dim file As Scripting.TextStream
    Set file = fso.CreateTextFile(file_path, True)
    file.Write content
    file.Close
End Function

' Passed test
Private Sub test_print_to_file()
    print_to_file ThisWorkbook.Path & "\output\print_to_file.txt", "Hello World!"
End Sub


' ================================= json =================================

' * 將字典或集合轉換成 JSON 字串
Function json_str(iterable As Variant, Optional indent As Integer = 0)
    ' 遞迴函式，用來處理多層的字典或集合
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

' Passed test
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


' ================================= csv =================================

' * 將字典或集合轉換成 CSV 字串
Function csv_str(iterable As Variant, Optional father As String = "")
    ' 遞迴函式，用來處理多層的字典或集合
    If TypeOf iterable Is Scripting.Dictionary Then
        csv_str = dict_csv(iterable, father)
    ElseIf TypeOf iterable Is Collection Then
        csv_str = collection_csv(iterable, father)
    ElseIf IsNumeric(iterable) Then 
        csv_str = father & iterable & vbCrLf
    Else
        csv_str = father & """" & iterable & """" & vbCrLf
    End If
End Function

Function dict_csv(dict As Variant, Optional father As String = "")
    dict_csv = ""
    For Each key In dict
        dict_csv = dict_csv & csv_str(dict(key), father & key & ", ")
    Next
End Function

Function collection_csv(clcn As Variant, Optional father As String = "")
    collection_csv = ""
    For Each item In clcn
        collection_csv = collection_csv & csv_str(item, father)
    Next
End Function

' Passed test
Private Sub test_csv_str()
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
    multilayer_dict.Add "key8", "value1"
    multilayer_dict.Add "key9", "value2"
    multilayer_dict.Add "key10", dict
    multilayer_dict.Add "key11", clcn

    pth = ThisWorkbook.Path & "\output\csv_str.csv"
    print_to_file pth, csv_str(multilayer_dict)

End Sub


' ================================= csv_expand_last =================================

' * 將字典或集合轉換成 CSV 字串，並將最後一層的容器展開(根據該容器第一個項目的型別判斷)
Function csv_expand_last_str(iterable As Variant, Optional father As String = "")

    If TypeOf iterable Is Scripting.Dictionary Then
        csv_expand_last_str = dict_expand_last_csv(iterable, father)
    ElseIf TypeOf iterable Is Collection Then
        csv_expand_last_str = collection_expand_last_csv(iterable, father)
    ElseIf IsNumeric(iterable) Then 
        csv_expand_last_str = father & iterable
    Else
        csv_expand_last_str = father & """" & iterable & """" 
    End If
End Function

Function dict_expand_last_csv(dict As Variant, Optional father As String = "")
    dict_expand_last_csv = ""
    
    k = dict.Keys()(1)
    ' 如果第一個項目不是容器，則直接展開
    If Not TypeOf dict(k) Is Scripting.Dictionary And Not TypeOf dict(k) Is Collection Then
        dict_expand_last_csv = father
        For Each key In dict
            dict_expand_last_csv = dict_expand_last_csv & csv_expand_last_str(dict(key), key & ": ") & ", "
        Next
        dict_expand_last_csv = dict_expand_last_csv & vbCrLf
    ' 如果第一個項目是容器，則遞迴處理
    Else
        For Each key In dict
            dict_expand_last_csv = dict_expand_last_csv & csv_expand_last_str(dict(key), father & key & ", ")
        Next
    End If
End Function

Function collection_expand_last_csv(clcn As Variant, Optional father As String = "")
    collection_expand_last_csv = ""
    
    If Not TypeOf clcn(1) Is Scripting.Dictionary And Not TypeOf clcn(1) Is Collection Then
        collection_expand_last_csv = father
        For Each item In clcn
            collection_expand_last_csv = collection_expand_last_csv & csv_expand_last_str(item, "") & ", "
        Next
        collection_expand_last_csv = collection_expand_last_csv & vbCrLf
    Else
        For Each item In clcn
            collection_expand_last_csv = collection_expand_last_csv & csv_expand_last_str(item, father)
        Next
    End If
End Function

' Passed test
Private Sub test_csv_expand_last_str()
    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary
    dict.Add "key1", "value1"
    dict.Add "key2", "value2"
    dict.Add "key3", "value3"
    dict.Add "key4", "value4"
    
    Dim multilayer_dict As Scripting.Dictionary
    Set multilayer_dict = New Scripting.Dictionary
    multilayer_dict.Add "key8", dict
    multilayer_dict.Add "key9", dict
    multilayer_dict.Add "key10", dict
    multilayer_dict.Add "key11", dict

    pth = ThisWorkbook.Path & "\output\csv_expand_last_str.csv"
    print_to_file pth, csv_expand_last_str(multilayer_dict)

End Sub

' ================================= join =================================

Function join_collection(coll As Collection, _
                        Optional delimiter As String = ",", _
                        Optional ByVal prefix As String = "", _
                        Optional ByVal suffix As String = "") As String

    Dim result As String
    Dim i As Integer
    result = prefix & coll(1) & suffix
    For i = 2 To coll.Count
        result = result & delimiter & prefix & coll(i) & suffix
    Next i
    join_collection = result

End Function