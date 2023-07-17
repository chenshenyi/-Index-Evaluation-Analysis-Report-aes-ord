
Attribute VB_Name = "import_modules"
' 2023-07-17 11:01:40

Private Sub importSelf()
    CODE_FOLDER = ThisWorkbook.path & "\..\src\"
    Dim project As VBIDE.VBProject
    Set project = ThisWorkbook.VBProject
    For Each component In project.VBComponents
        If component.name = "import_modules" Then
            component.name = "toRemove"
            project.VBComponents.Import CODE_FOLDER & "import_modules.bas"
            project.VBComponents.Remove component
            Exit Sub
        End If
    Next component
    
End Sub

Private Sub Builder()
    from_path = ThisWorkbook.path & "\A 主程式.xlsm"
    to_path = ThisWorkbook.path & "\..\template\A 主程式.xlsm"
    copy_file from_path, to_path

    from_path = ThisWorkbook.path & "\A 主程式.xlsm"
    to_path = ThisWorkbook.path & "\..\example\A%20主程式.xlsm"
    copy_file from_path, to_path

    from_path = ThisWorkbook.path & "\..\docs\自動化程式說明書.pdf"
    to_path = ThisWorkbook.path & "\..\template\自動化程式說明書.pdf"
    copy_file from_path, to_path

    from_path = ThisWorkbook.path & "\..\docs\自動化程式說明書.pdf"
    to_path = ThisWorkbook.path & "\..\example\自動化程式說明書.pdf"
    copy_file from_path, to_path
End Sub

Sub importModules()
    CODE_FOLDER = ThisWorkbook.path & "\.src\"

    Dim project As VBIDE.VBProject
    Set project = ThisWorkbook.VBProject

    Dim component As VBIDE.VBComponent
    For Each component In project.VBComponents
        If component.name <> "import_modules" And _
        component.Type = vbext_ct_StdModule Then
            project.VBComponents.Remove component
        End If
    Next component

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(CODE_FOLDER)

    For Each file In folder.files
        If right(file.name, 4) = ".bas" Then
            If file.name <> "import_modules.bas" Then
                project.VBComponents.Import file.path
            End If
        End If
    Next file
End Sub

Function getFileFromGit(ByVal modulefile As String)
    Dim response As String

    github_path = "https://raw.githubusercontent.com/chenshenyi/Index-Evaluation-Analysis-Report/main/src/"
    address = github_path & modulefile

    With CreateObject("MSXML2.XMLHTTP")
        .Open "GET", address, False
        .setRequestHeader "pragma", "no-cache"
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .setRequestHeader "cache-control", "no-store, must-revalidate, private"
        .send
        response = BinToStr(.responseBody, "big5")
        
    End With

    Open ThisWorkbook.path & "\.src\" & modulefile For Output As #1
    Print #1, response
    Close #1
End Function

Function BinToStr(arrBin, strChrs)
    With CreateObject("ADODB.Stream")
        .Type = 2 ' adModeRead
        .Open
        .Writetext arrBin
        .Position = 0
        .Charset = strChrs
        .Position = 3 ' skip BOM
        BinToStr = .ReadText
        .Close
    End With
End Function

Function readIndex()
    getFileFromGit "index.json"
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file As Object
    Set file = fso.OpenTextFile(ThisWorkbook.path & "\.src\index.json", 1)
    Dim json As String
    json = file.ReadAll
    file.Close
    Set file = Nothing
    Set fso = Nothing
    json = Replace(json, "[", "")
    json = Replace(json, "]", "")
    json = Replace(json, """", "")
    json = Replace(json, vbNewLine, "")
    json = Replace(json, chr(10), "")
    json = Replace(json, chr(13), "")
    json = Replace(json, " ", "")

    readIndex = split(json, ",")
    
End Function

Sub upgrade()
    ' test if .src folder exists, if not, create it
    If Dir(ThisWorkbook.path & "\.src", vbDirectory) = "" Then
        MkDir ThisWorkbook.path & "\.src"
    End If
    For Each modulefile In readIndex
        getFileFromGit modulefile
    Next modulefile
    importModules
End Sub