Attribute VB_Name = "Functions"
Option Compare Database
Function currentSiteID()
    Dim db As Database
    Dim rs As Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("readCurrentSite", dbOpenSnapshot)
    
    If rs.recordcount >= 1 Then
        currentSiteID = rs.Fields(0)
    Else
        currentSiteID = 1
    End If
    
    rs.Close
    db.Close
    
End Function

Function submitData(typeSelector As String, Data) As String
    
    Select Case typeSelector
        'Submitting data to the tbl1_direCare table
        Case "Direct Care"
            If IsNull(Data(1)) = True Or IsNumeric(Data(1)) = False Then
                submitData = "Missing SiteID"
                Exit Function
            End If
            
            If IsNull(Data(2)) = True Or IsNumeric(Data(2)) = False Then
                submitData = "Missing ActivityID"
                Exit Function
            End If

            If IsNull(Data(3)) = True Or IsNumeric(Data(3)) = False Then
                submitData = "Missing Hours"
                Exit Function
            End If
            
            If IsNull(Data(4)) = True Or IsDate(Data(4)) = False Then
                submitData = "Missing Date"
                Exit Function
            End If
            
            If IsNull(Data(5)) = True Or IsNumeric(Data(5)) = False Then
                submitData = "Missing Members Engaged"
                Exit Function
            End If
            
            If IsNull(Data(6)) Then
                Data(6) = ""
            End If
            
            'create the SQL String to insert the data
            sqlString = "INSERT INTO tbl2_DirectCareEvents(siteID, activityID, hours, eventDate, memberEngaged, notes)" & vbLf & _
            "VALUES(" & Data(1) & ", " & Data(2) & ", " & Data(3) & ", " & Chr(34) & Data(4) & Chr(34) & ", " & _
             Data(5) & ", " & Chr(34) & Data(6) & Chr(34) & ")"
            
            'execute the SQL Command
            CurrentDb.Execute sqlString, dbFailOnError
        
        'submitting data to tbl1_nonpatientCare table
        Case "nonpatientCareEvents"
            If IsNull(Data(1)) = True Or IsNumeric(Data(1)) = False Then
                submitData = "Missing SiteID"
                Exit Function
            End If
            
            If IsNull(Data(2)) = True Or IsNumeric(Data(2)) = False Then
                submitData = "Missing ActivityID"
                Exit Function
            End If

            If IsNull(Data(3)) = True Or IsNumeric(Data(3)) = False Then
                submitData = "Missing Hours"
                Exit Function
            End If
            
            If IsNull(Data(4)) = True Or IsDate(Data(4)) = False Then
                submitData = "Missing Date"
                Exit Function
            End If
            
            If IsNull(Data(5)) = True Then
                Data(5) = ""
            End If
            
            'create the SQL String to insert data
            sqlString = "INSERT INTO tbl2_nonpatientCareEvents(siteID, activityID, hours, eventDate, information)" & vbLf & _
            "VALUES(" & Data(1) & ", " & Data(2) & ", " & Data(3) & ", " & Chr(34) & Data(4) & Chr(34) & ", " & Chr(34) & Data(5) & Chr(34) & ")"
            
            'execute the SQL Command
            CurrentDb.Execute sqlString, dbFailOnError
            
        Case "surveyAnswers"
        
            For i = 1 To UBound(Data)
                'create the sql string
                sqlString = "INSERT INTO tbl2_answerLog(siteID, surveyTypeID, questionID, answer, surveyDate)" & vbLf & _
                "VALUES(" & Data(i, 1) & ", " & Data(i, 2) & ", " & Data(i, 3) & ", " & _
                Chr(34) & Data(i, 4) & Chr(34) & " ," & Chr(34) & Data(i, 5) & Chr(34) & ")"
                
                'execute the command
                CurrentDb.Execute sqlString, dbFailOnError
            Next i
        
        Case "narative"
            If IsNull(Data(1)) = True Or IsNumeric(Data(1)) = False Then
                submitData = "Missing siteID"
                Exit Function
            End If
            If IsNull(Data(2)) = True Or IsDate(Data(2)) = False Then
                submitData = "Missing date"
                Exit Function
            End If
            If IsNull(Data(3)) = True Then
                submitData = "Blank narative"
            End If
            'Create the SQL String to insert Data
            sqlString = "INSERT INTO tbl2_monthlyNaratives(siteID, narativeDate, narative)" & vbLf & _
            "VALUES(" & Data(1) & "," & Chr(34) & Data(2) & Chr(34) & ", " & Chr(34) & Data(3) & Chr(34) & ")"
            
            CurrentDb.Execute sqlString, dbFailOnError
        Case Else
            submitData = "Failed Selector, check code"
            Exit Function
    End Select

    submitData = "Success"

End Function

Function getQuestions(Data() As Variant) As Boolean
    'Desc: Fills an empty data array with questions and a space for the answers
    Dim db As Database
    
    On Error GoTo errorHandler
    Set db = CurrentDb
    Set rs = db.OpenRecordset("readSurveyOne", dbOpenSnapshot)
    
    rs.MoveLast
    rs.MoveFirst
    
    ReDim Data(1 To rs.recordcount, 1 To 5)
    'array indexes: 1. surveyType, 2. questionID, 3. questionText, 4. Question Type, 5. answer(blank)
    
    'loop through each record in the set and add the id and question text
    For i = 1 To rs.recordcount
        
        Data(i, 1) = rs.Fields(0)
        Data(i, 2) = rs.Fields(1)
        Data(i, 3) = rs.Fields(2)
        Data(i, 4) = rs.Fields(3)
        
        rs.MoveNext
    Next i
    
    'close the record, close the database
    rs.Close
    db.Close
    
    'return true for successful pull
    getQuestions = True
    Exit Function

errorHandler:
    'imedately close the set and database in the case of an error
    rs.Close
    db.Close
    'return false for failed pull
    getQuestions = False
    'print the debug message with error number and description to the console
    Debug.Print ("error#" & Err.Number & ": " & Err.Description)
    
End Function

Sub exportData()
    Dim db As Database
    Set db = CurrentDb
    
    nDate = jxNow()
    cSiteID = currentSiteID()
    
    fPath = CurrentProject.Path & "\" & nDate & " export\"
    
    'If there is no Directory, make one
    If Len(Dir(fPath, vbDirectory)) = 0 Then
        MkDir fPath
    End If
    
    'set the location and file name for each file that will be exported
    cstlocation1 = fPath & "Site " & cSiteID & "_" & nDate & " DirectCareEvents.csv"
    cstlocation2 = fPath & "Site " & cSiteID & "_" & nDate & " Monthly Narative.csv"
    cstlocation3 = fPath & "Site " & cSiteID & "_" & nDate & " nonpaitentCareEvents.csv"
    cstlocation4 = fPath & "Site " & cSiteID & "_" & nDate & " Survey Answers.csv"
    
    On Error GoTo errorHandler
    '1 - Export as CSV
    DoCmd.TransferText acExportDelim, , "tbl2_DirectCareEvents", cstlocation1, True
    DoCmd.TransferText acExportDelim, , "tbl2_monthlyNaratives", cstlocation2, True
    DoCmd.TransferText acExportDelim, , "tbl2_nonpatientCareEvents", cstlocation3, True
    DoCmd.TransferText acExportDelim, , "tbl2_answerLog", cstlocation4, True
    
    '2 - Append to Archive
    CurrentDb.Execute "appendDirectCare"
    CurrentDb.Execute "appendMonthlyNaratives"
    CurrentDb.Execute "appendNonpatientCare"
    CurrentDb.Execute "appendAnswerLog"
    
    '3 - Delete from Event
    CurrentDb.Execute "delDirectCare"
    CurrentDb.Execute "delMonthlyNaratives"
    CurrentDb.Execute "delNonpatientCare"
    CurrentDb.Execute "delAnswerLog"
    
    MsgBox ("Successful Export!")
    
    Exit Sub
    
errorHandler:

    MsgBox ("Unsuccessful Export...")
    Debug.Print ("error " & Err.Number & ": " & Err.Description)
    
End Sub

Sub importData()
    'Description: Import data from CSV File
    Set db = CurrentDb
    
    On Error GoTo errorHandler
    
    '1 - Import CSV to importEvents Table
    DoCmd.TransferText acImportDelim, , "importEvents", selectFile(), True
    '2 - Append to allEvents
    db.Execute ("1_appendtoAllEvents")
    '3 - Delete from importEvents Table
    db.Execute ("2_deleteImportEvents")
    
    MsgBox ("Successful Import!")
    Exit Sub
    
errorHandler:
    
    MsgBox ("unsuccessful Import...")
    Debug.Print ("error " & Err.Number & ": " & Err.Description)
    
End Sub

Function selectFile()
    Dim fd As FileDialog
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .AllowMultiSelect = False
        If .show Then
            selectFile = .SelectedItems(1)
        Else
            End
        End If
    End With
    
    Debug.Print selectFile
    Set fd = Nothing
    
End Function


