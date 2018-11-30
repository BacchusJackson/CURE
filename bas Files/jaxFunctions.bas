Attribute VB_Name = "jaxFunctions"
Option Compare Database

Function jxStringIn(Data)
'Desc: Takes an array and converts it to a string that can be outputted back to an array
    
    jString = ""
    iRow = 1
    Do
        
        For iCol = 1 To UBound(Data, 2)
            'add text and value split diamond
            jString = jString & Data(iRow, iCol) & "<|>"
        Next iCol
        'add block split diamond
            jString = jString & "<||>"
        
        iRow = iRow + 1
    Loop Until iRow = UBound(Data) + 1

    jxStringIn = jString

End Function

Function jxStringOut(jxString, Data)
'Desc: takes a jxString and converts it to a passed empty array
        
    On Error GoTo errorHandler
    blocks = Split(jxString, "<||>")
    
    numOfCols = UBound(Split(blocks(1), "<|>"))
    
    ReDim Data(1 To UBound(blocks), 1 To numOfCols)
    
    For i = 1 To UBound(Data)
        jValues = Split(blocks(i - 1), "<|>")
        
        For E = 1 To UBound(jValues)
            Data(i, E) = jValues(E - 1)
        Next E
        
    Next i
    
    jxStringOut = True
    Exit Function
    
errorHandler:
    Debug.Print "Error #" & Err.Number & ": " & Err.Description

End Function
