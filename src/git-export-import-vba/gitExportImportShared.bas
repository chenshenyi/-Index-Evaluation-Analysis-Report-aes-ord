Attribute VB_Name = "gitExportImportShared"
'In order for this code to work properly you must follow these steps:
'Tools -> References -> Add Microsoft Visual Basic for Applications Extensibility 5.3
'Excel Ribbon -> Developer -> Macro Security -> Trust access to the VBA project object model
'
'
'Some code was refactored from https://www.vitoshacademy.com/vba-source-control-with-git-video/

Sub GitSave(gitPath, ignoreFrx As Boolean, wbExport As Boolean, Optional modulesArray As Variant = "")
                 
    If IsArray(modulesArray) = True Then
        
        tempFolder = CreateTempFolder(gitPath)
        
        ExportToTemp tempFolder, modulesArray
        
        CopyFiles tempFolder, gitPath, ignoreFrx, True
    
    End If
    
    If wbExport = True Then
    
        If MsgBox("Do you wish to save your Workbook before exporting?", vbYesNo) = vbYes Then
        
            ThisWorkbook.Save
        
        End If
    
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        Call fso.copyfile(ThisWorkbook.Path & "\" & ThisWorkbook.Name, gitPath, True)
    
    End If
    
    MsgBox "All files have been exported."
    
End Sub

Function OpenFolderDialog()

    On Error GoTo error
  
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Show
                
        OpenFolderDialog = .SelectedItems(1) & "\"
 
    End With
           
    Exit Function
    
error:

    If err.Number = 5 Then
            
        OpenFolderDialog = ""
            
        Exit Function
        
    Else
    
        err.Raise (err.Number)
    
    End If
 
End Function

Function CreateTempFolder(gitPath)
        
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim tempFolder As String: tempFolder = gitPath & "tempVBA"
        
    On Error Resume Next
    fso.DeleteFolder tempFolder
    On Error GoTo 0
    
    MkDir tempFolder
    
    CreateTempFolder = tempFolder & "\"
    
End Function

Sub ExportToTemp(tempFolder, modulesArray As Variant) 'modulesArray must be an Array.
                
    If IsEmpty(modulesArray) Then Exit Sub
    
    Dim wkb As Workbook: Set wkb = Excel.Workbooks(ThisWorkbook.Name)
    
    Dim unitsCount As Long
    Dim filePath As String
    Dim component As VBIDE.VBComponent
    Dim tryExport As Boolean
    
    For Each component In wkb.VBProject.VBComponents
        tryExport = True
        filePath = component.Name
               
        Select Case component.Type
            Case vbext_ct_ClassModule
                filePath = filePath & ".cls"
            Case vbext_ct_MSForm
                filePath = filePath & ".frm"
            Case vbext_ct_StdModule
                filePath = filePath & ".bas"
            Case vbext_ct_Document
                tryExport = False
        End Select
        
        If Not IsError(Application.Match(filePath, modulesArray, 0)) = True Then
        
            If tryExport Then
                Debug.Print unitsCount & " exporting " & filePath
                component.Export tempFolder & "\" & filePath
            End If
            
        End If
    Next
    
    Debug.Print "Exported at " & tempFolder
    
End Sub

Sub CopyFiles(sourceFolder, targetFolder, Optional ignoreFrx As Boolean = False, Optional killSource As Boolean = False)

    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set src = fso.GetFolder(sourceFolder)
    Set dest = fso.GetFolder(targetFolder)
    
    For Each srcFile In src.Files
        
        Dim srcFilepath As String
        srcFilepath = srcFile.Path
        
        If Right(srcFilepath, Len(srcFilepath) - InStrRev(srcFilepath, ".") + 1) = ".frx" Then 'extension includes the "."
                   
            If ignoreFrx = False Then
            
                srcFile.Copy targetFolder, True 'I set Overwrite to True
                
            End If
            
        Else
        
            srcFile.Copy targetFolder, True 'I set Overwrite to True
            
        End If
        
    Next srcFile
    
    If killSource = True Then
        
        On Error Resume Next
        
            fso.DeleteFolder Left(sourceFolder, Len(sourceFolder) - 1) 'Remove "\" at the end of string.
            
        On Error GoTo 0
    
    End If

End Sub

Sub RemoveVBComponents(VBProject As Object, ParamArray ComponentNames() As Variant)
'https://stackoverflow.com/questions/57809025/delete-some-vba-forms-and-modules-leave-others

    Dim Item
    
    For Each Item In ComponentNames
    
        On Error Resume Next
            VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(Item)
        On Error GoTo 0
    
    Next
    
End Sub




