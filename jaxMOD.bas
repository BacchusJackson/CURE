Attribute VB_Name = "jaxMOD"
Option Compare Database

Sub exportMods()
    Dim i As Integer
    
    'find the correct vba project for the file currrently active
    For i = 1 To Application.VBE.VBProjects.Count
        a = Split(Application.VBE.VBProjects(i).FileName, "\")
        B = Split(CurrentDb.Name, "\")
        
        'if there is a match
        If a(UBound(a)) = B(UBound(B)) Then
            
            projID = i
            
        End If
    Next i
    
    'select file path
    fPath = getFilePath() & "\"
    
    With Application.VBE.VBProjects(projID).VBComponents

        For i = 1 To .Count
            
            If .Item(i).Type = 1 Then
                .Item(i).Export fPath & .Item(i).Name & ".bas"
            End If
        
        Next i
        
    End With
    
    
End Sub

Function getFilePath()

    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fDialog
        .AllowMultiSelect = True
        .Title = "Select Folder to Output Source Code"
        .show
    End With
    
    If fDialog.SelectedItems.Count = 1 Then
        getFilePath = fDialog.SelectedItems(1)
    Else
        getFilePath = False
    End If
    

End Function
