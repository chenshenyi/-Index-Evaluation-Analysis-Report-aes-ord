Attribute VB_Name = "system_operation"
' 2023-07-11 17:15:58

' Passed all test

' ============================================= Copy =============================================

Function copy_file(ByVal from_path As String, ByVal to_path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile from_path, to_path
End Function

' Passed test
Private Sub test_copy_file()
    Call copy_file( ThisWorkbook.path & "\input\copy_file.txt", ThisWorkbook.path & "\output\copy_file.md")
End Sub


Function copy_folder(ByVal from_path As String, ByVal to_path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFolder from_path, to_path
End Function

' Passed test
Private Sub test_copy_folder()
    Call copy_folder( ThisWorkbook.path & "\input\copy_folder", ThisWorkbook.path & "\output\copy_folder")
End Sub


' ============================================= Create =============================================

Function create_folder(ByVal folder_path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateFolder folder_path
End Function

' Passed test
Private Sub test_create_folder()
    Call create_folder( ThisWorkbook.path & "\output\create_folder")
End Sub


Function create_file(ByVal file_path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateTextFile file_path
End Function

' Passed test
Private Sub test_create_file()
    Call create_file( ThisWorkbook.path & "\output\create_file.txt")
End Sub


' ============================================= Delete =============================================

Function delete_file(ByVal file_path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.DeleteFile file_path
End Function

' Passed test
Private Sub test_delete_file()
    Call create_file( ThisWorkbook.path & "\output\delete_file.txt")
    Call delete_file( ThisWorkbook.path & "\output\delete_file.txt")
End Sub


Function delete_folder(ByVal folder_path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.DeleteFolder folder_path
End Function

' Passed test
Private Sub test_delete_folder()
    Call create_folder( ThisWorkbook.path & "\output\delete_folder")
    Call delete_folder( ThisWorkbook.path & "\output\delete_folder")
End Sub