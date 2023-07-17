
Attribute VB_Name = "system_operation"
' 2023-07-11 17:15:58

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

' ============================================= Existing =============================================

Function detect_files_in_folder(ByVal folder As String, _
                                files As Collection, _
                                Optional ByVal file_type As String = "") As Collection

    Dim found As Collection
    Dim notFound As Collection
    Set found = New Collection
    Set notFound = New Collection

    For Each file In files
        file_path = folder & "\" & file & file_type
        If Dir(file_path) = "" Then
            notFound.Add file
        Else
            found.Add file
        End If
    Next file

    If notFound.Count > 0 Then
        MsgBox "以下檔案不存在：" & vbNewLine & _
            join_collection(notFound, vbNewLine, "◆ ", file_type) & vbNewLine & _
            vbNewLine & _
            "僅匯入以下檔案：" & vbNewLine & _
            join_collection(found, vbNewLine, "◆ ", file_type)
    Else
        MsgBox "匯入以下檔案：" & vbNewLine & _
            join_collection(found, vbNewLine, "◆ ", file_type)
    End If

    Set detect_files_in_folder = found

End Function

' Passed test
Private Sub test_detect_files_in_folder()

    argument_init

    Dim colleges As Collection
    Set colleges = New Collection
    For Each college In college_department_dict.Keys
        colleges.Add college & ".docx"
    Next college

    Dim found As Collection
    Set found = detect_files_in_folder(ThisWorkbook.Path & "\2. 各院報告書模板", colleges)
End Sub

