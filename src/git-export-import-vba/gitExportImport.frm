VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} gitExportImport 
   Caption         =   "Git Export / Import"
   ClientHeight    =   8655.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4635
   OleObjectBlob   =   "gitExportImport.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "gitExportImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub updateModuleList(listName As String, gitFolder)

        gitExportImport.Controls(listName).Clear
    
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        On Error GoTo escape
        Set Folder = fso.GetFolder(gitFolder)
        
        For Each file In Folder.Files
        
            extension = fso.GetExtensionName(file)
            
            If extension = "bas" Or extension = "cls" Or extension = "frm" Then
            
                If file.Name <> "gitExportImport.frm" And file.Name <> "gitExportImportShared.bas" Then
                
                    gitExportImport.Controls(listName).AddItem file.Name
                
                End If
                
            End If
        
        Next file
escape:
End Sub

Private Sub navExport_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    MultiPage.Value = 0

End Sub

Private Sub navImport_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    MultiPage.Value = 1
    
    On Error Resume Next
        updateModuleList "moduleListImport", gitFolderLabelImport.Caption
    On Error Resume Next

End Sub

Private Sub UserForm_Initialize()
    
    Set wkb = ThisWorkbook
    Dim component As VBIDE.VBComponent
    
    For Each component In wkb.VBProject.VBComponents
        
        Select Case component.Type
        
            Case vbext_ct_ClassModule
                moduleName = component.Name & ".cls"
                
            Case vbext_ct_MSForm
                moduleName = component.Name & ".frm"
                
            Case vbext_ct_StdModule
                moduleName = component.Name & ".bas"
                
            Case vbext_ct_Document
                moduleName = ""
                                
        End Select
        
        If moduleName <> "" Then moduleListExport.AddItem moduleName

    Next

End Sub

'The code below refers to gitExport page
'--------------------------------------------------------------------------------------------------------------

Private Sub selectAllExport_Click()

    For i = 0 To moduleListExport.ListCount - 1
    
        moduleListExport.Selected(i) = True
    
    Next i

End Sub

Private Sub gitFolderImageExport_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    gitFolder = OpenFolderDialog()
    
    If gitFolder <> "" Then gitFolderLabelExport.Caption = gitFolder
   
End Sub

Private Sub exportFiles_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If MsgBox("Do you wish to proceed?", vbYesNo) = vbNo Then Exit Sub

    Dim ignoreFrx As Boolean
    Dim wbExport As Boolean
    ignoreFrx = False
    wbExport = False
    arrayBound = 0
    noneSelected = True
   
    
    If gitFolderLabelExport.Caption = "Select a directory using the folder icon." Then
    
        MsgBox "Select a valid folder and try again."
        
        Exit Sub
    
    End If
    
    For i = 0 To moduleListExport.ListCount - 1
    
        If moduleListExport.Selected(i) = True Then
                
            Dim modulesArray()
        
            ReDim Preserve modulesArray(0 To arrayBound)
            
            modulesArray(arrayBound) = moduleListExport.List(i)
               
            arrayBound = arrayBound + 1
                
            noneSelected = False
            
            sendArray = True
                  
        End If
    
    Next i
    
    If wbExportCheck.Value = True Then noneSelected = False
    
    If noneSelected = True Then
    
        MsgBox "Please select at least one file to export."
        
        Exit Sub
    
    End If
    
    If ignoreFrxCheck.Value = True Then ignoreFrx = True
           
    
    If wbExportCheck.Value = True Then wbExport = True
    
        
    If sendArray = True Then
    
        GitSave gitFolderLabelExport.Caption, ignoreFrx, wbExport, modulesArray
    
    Else
    
        GitSave gitFolderLabelExport.Caption, ignoreFrx, wbExport
    
    End If
    
End Sub

'The code below refers to gitImport page
'--------------------------------------------------------------------------------------------------------------

Private Sub gitFolderImageImport_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    gitFolder = OpenFolderDialog()
    
    If gitFolder <> "" Then
                 
        gitFolderLabelImport.Caption = gitFolder
    
        updateModuleList "moduleListImport", gitFolder

    End If
    
End Sub

Private Sub selectAllImport_Click()

    For i = 0 To moduleListImport.ListCount - 1
    
        moduleListImport.Selected(i) = True
    
    Next i

End Sub

Private Sub deleteFiles_Click()

    If MsgBox("Do you really wish to delete the selected files from the directory?", vbYesNo) = vbNo Then Exit Sub
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    For i = 0 To moduleListImport.ListCount - 1
    
        If moduleListImport.Selected(i) = True Then
        
            Module = moduleListImport.List(i)
            
            If Right(Module, 4) = ".frm" And fso.FileExists(gitFolderLabelImport & Left(Module, Len(Module) - 4) & ".frx") Then
              
                    If MsgBox("You are trying to delete " & Module & ". Do you wish to delete its .frx reference?", vbYesNo) = vbYes Then
                    
                       fso.deleteFile (gitFolderLabelImport & Left(Module, Len(Module) - 4) & ".frx")
                    
                    End If
            
            End If
            
            fso.deleteFile (gitFolderLabelImport & Module)
                  
        End If
    
    Next i
    
    updateModuleList "moduleListImport", gitFolderLabelImport.Caption

End Sub

Private Sub importFiles_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If gitFolderLabelImport.Caption = "Select a directory using the folder icon." Then
    
        MsgBox "Select a valid folder and try again."
        
        Exit Sub
    
    End If

    If MsgBox("Do you wish to proceed? All conflicting modules will be overwritten.", vbYesNo) = vbNo Then Exit Sub
    
    Dim noneSelected As Boolean
    noneSelected = True
    
    Set Components = ThisWorkbook.VBProject.VBComponents
    
    For i = 0 To moduleListImport.ListCount - 1
    
        If moduleListImport.Selected(i) = True Then
                
            noneSelected = False
    
            Module = moduleListImport.List(i)
                      
            If Right(Module, 4) = ".frm" And fso.FileExists(gitFolderLabelImport & Left(Module, Len(Module) - 4) & ".frx") = False Then
                      
                MsgBox "Could not import " & Module & " because the .frx file is missing." & vbNewLine & "Please check your directory and try again."
            
            Else
            
                RemoveVBComponents ThisWorkbook.VBProject, Left(Module, Len(Module) - 4)
                
                Components.Import gitFolderLabelImport.Caption & Module
            
            End If
            
        End If
    
    Next i
    
    If noneSelected = False Then
    
        MsgBox "Task completed."
    
    Else
        
        MsgBox "Select at least one file and try again."
    
    End If
    
End Sub
