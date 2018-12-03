Attribute VB_Name = "jaxMOD"
Option Compare Database

Sub exportMods()
    Dim i As Integer
    
    'find the correct vba project for the file currrently active
    For i = 1 To Application.VBE.VBProjects.Count
        a = Split(Application.VBE.VBProjects(i).FileName, "\")
        b = Split(CurrentDb.Name, "\")
        
        'if there is a match
        If a(UBound(a)) = b(UBound(b)) Then
            
            projID = i
            
        End If
    Next i
    
    'select file path
    fPath = getFilePath() & "\"
    
    'the vba modules
    With Application.VBE.VBProjects(projID).VBComponents

        For i = 1 To .Count
            
            If .Item(i).Type = 1 Then
                .Item(i).Export fPath & .Item(i).Name & ".bas"
            End If
        
        Next i
        
    End With
    
    
End Sub
