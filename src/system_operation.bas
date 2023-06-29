Attribute VB_Name = "system_operation"

Function copied_to(from_path As String, to_path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile from_path, to_path
End Function

Function copied_to_folder(from_path As String, to_folder As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile from_path, to_folder
End Function

Function copied_folder(from_folder As String, to_folder As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFolder from_folder, to_folder
End Function

Function deleted_file(file_path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.DeleteFile file_path
End Function

Function deleted_folder(folder_path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.DeleteFolder folder_path
End Function

Function created_folder(folder_path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateFolder folder_path
End Function

Function created_file(file_path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateTextFile file_path
End Function
