Attribute VB_Name = "tests"
Option Compare Database

Sub testInput()
    Dim Data(1 To 5, 1 To 3) As Variant
    
    
    For i = 1 To 5
        Data(i, 1) = i
        Data(i, 2) = "Question? " & i
        Data(i, 3) = "Answer. " & i
    Next i
    
    jxStringIn Data
    
    a = 1
    
End Sub

Sub TestOutput()
    Dim Data() As Variant
    
    jxString = "1<|>Question? 1<|>Answer. 1<|><||>2<|>Question? 2<|>Answer. 2<|><||>3<|>Question? 3<|>Answer. 3<|><||>4<|>Question? 4<|>Answer. 4<|><||>5<|>Question? 5<|>Answer. 5<|><||>"
    jxStringOut jxString, Data
    
    
End Sub

Sub fullTest()
    Dim Data(1 To 5, 1 To 3) As Variant
    Dim data2() As Variant
    Dim data3() As Variant
    
    For i = 1 To 5
        Data(i, 1) = i
        Data(i, 2) = "Question? " & i
        Data(i, 3) = "Answer. " & i
    Next i
    
    jxString = jxStringIn(Data)
    
    response = jxStringOut(jxString, data2)
    
    If response = True Then
        Debug.Print data2(1, 2)
    Else
        Exit Sub
    End If
    
    data2(1, 3) = "This is the answer to question 2"
    
    jxString = jxStringIn(data2)
    
    Debug.Print jxString
    
    
    

End Sub


