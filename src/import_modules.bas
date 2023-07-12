Attribute VB_Name = "import_modules"
' 2023-07-12 14:21:38

Sub importModules()
    CODE_FOLDER = ThisWorkbook.Path & "\..\src\"

    Dim project As VBIDE.VBProject
    Set project = ThisWorkbook.VBProjecot


    ' �R���Ҧ��Ҳ�(���F import_modules����)
    Dim component As VBIDE.VBComponent
    For Each component In project.VBComponents
        If component.Name <> "import_modules" And _
        component.Type = vbext_ct_StdModule Then
            project.VBComponents.Remove component
        End If
    Next component

    ' �N src ��Ƨ������Ҧ� .bas �ɮ׶פJ
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(CODE_FOLDER)

    For Each file In folder.Files
        If Right(file.Name, 4) = ".bas" Then
            if file.Name <> "import_modules.bas" Then
                project.VBComponents.Import file.Path
            End If
        End If
    Next file
End Sub

