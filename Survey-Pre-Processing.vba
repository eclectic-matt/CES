'--------------------------------------
' FOR SPLITTING AND SANITISING THE DATA
'--------------------------------------

'----------------------------
' @name getResponseThreshold
' @param CohortSize = the number of eligible students
' @descr A standard function to return the threshold for the disclaimer message
' @usage Run on each cohort (e.g. year) to check if responseCount > threshold
'----------------------------
Function getResponseThreshold(ByVal cohortSize)
    getResponseThreshold = Application.Max(4, Round(((cohortSize + 0.5) / 2), 0))
    If gblnDebugging Then
        Debug.Print "Cohort = " & cohortSize & " so THR = " & getResponseThreshold
    End If
End Function

'----------------------------
' @name splitCourseSheets
' @descr Splits the reponses sheet into separate sheets for each course (by code) and copies student responses
' @usage Run on the RESPONSES sheet to split into course sheets for reporting
'----------------------------
Sub splitCourseSheets()

Call startTimer

If gblnDebugging Then
    Debug.Print "----"
    Debug.Print "---- START Split Course Sheets"
    Debug.Print "----"
Else
    Application.ScreenUpdating = False
End If

Dim studyYearWb As Workbook
Dim studyYearWs As Worksheet
'NEEDS COURSE YEAR VLOOKUP'ED INTO THE SHEET
fileBoo = IsWorkBookOpen(gstrStudyYearFile)
If fileBoo = True Then
    Set studyYearWb = Workbooks(gstrStudyYearName)     'File is open
Else
    Set studyYearWb = Workbooks.Open(gstrStudyYearFile) 'File is Closed
End If
Windows("Responses - FINAL data including partial responses 22.6.16.xlsx"). _
        Activate
Range("CE3").Select
ActiveCell.FormulaR1C1 = _
    "=VLOOKUP(RC[-72],'[EDITED - FORMATTED STUDENT SURVEY DATA WITH YEAR OF PROGRAMME.xlsx]List_Frame'!C3:C4,2,FALSE)"
Range("CE3").Select
Selection.AutoFill Destination:=Range("CE3:CE2004")
Range("CE3:CE2004").Select
Columns("CE:CE").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

' Setup ORIG workbook/sheet - duplicated for reports
Dim origWb As Workbook
'Dim gstrOrigWbFile As String

Dim courseWb As Workbook
Dim courseWs As Worksheet

Dim EA As Excel.Application
Set EA = New Excel.Application

' Check ORIG workbook open (responses) and set to origWb/origWs
fileBoo = IsWorkBookOpen(gstrOrigWbFile)
If fileBoo = True Then
    Set origWb = Workbooks(gstrOrigWbName)     'File is open
Else
    Set origWb = Workbooks.Open(gstrOrigWbFile) 'File is Closed
End If
Dim origWs As Worksheet
Set origWs = origWb.Sheets(gstrOrigWsName)

' Reference Workbook contains cohort sizes etc
fileBoo = IsWorkBookOpen(gstrRefWbFile)
If fileBoo = True Then
    Set refWb = Workbooks(gstrRefWbName)
Else
    Set refWb = Workbooks.Open(gstrRefWbFile)
End If
Set refWs = refWb.Sheets("COURSES")

' Now set up copied courseWb/courseWs
origWs.Copy
Set courseWb = ActiveWorkbook
Set courseWs = courseWb.ActiveSheet

' Adding summary data sheet
courseWb.Sheets.Add Before:=Worksheets(Worksheets.count)
courseWb.Sheets(Worksheets.count - 1).Name = "Summary Data"
Dim sumWs As Worksheet
Set sumWs = courseWb.Sheets("Summary Data")
With sumWs
    .Cells(1, 1) = "Course Code"
    .Cells(1, 2) = "Course Title"
    .Cells(1, 3) = "Study Year"
    .Cells(1, 4) = "Cohort Size"
    .Cells(1, 5) = "Response Rate (%)"
    .Cells(1, 6) = "Average Satisfaction"
    .Cells(1, 7) = "Median Satisfaction"
    .Cells(1, 8) = "Valid Responses"
    .Cells(1, 9) = "Published Flag"
    .Cells(1, 10) = "Department"
    .Cells(1, 11) = "School"
End With

' Now sort this sheet by COURSE ready to sanitise
courseWs.Name = "Course Reports"
fullResultsRange = "A2:CE" & (gintTotalRecords + 1)
courseWs.Range(fullResultsRange).Sort key1:=courseWs.Range("$T:$T"), key2:=courseWs.Range("$CE:$CE"), Header:=xlYes, MatchCase:=False, Orientation:=xlTopToBottom, SortMethod:=xlPinYin

finalRow = courseWs.Range("A" & Rows.count).End(xlUp).Row
starterRow = 3                      ' starts at row 3, rows 1+2 are headers
nameFound = ""                      ' set to the name of the "found" course
courseCount = 1                     ' increases when courses found in sorted list
currentRow = starterRow - 1
courseColName = "$T$" & starterRow & ":$T$" & finalRow
coursesCol = courseWs.Range(courseColName)

studyYearName = "$CE$" & starterRow & ":$CE$" & finalRow
studyYearCol = courseWs.Range(studyYearName)

lastRow = gintSheetDataRows + 1
courseWs.Select
' Loop through the courses column separating out found courses
For Each Cell In coursesCol
    courseCodeFound = Trim(Left(Cell, InStr(1, Cell, "-") - 1))
    currentRow = currentRow + 1
    If courseCodeFound = nameFound Then
        ' Add to course count
        courseCount = courseCount + 1
    Else
        If courseCodeFound <> nameFound And nameFound <> "" Then
            ' Process previous course data
            courseTitleFound = Trim(Mid(cellContents, Len(courseCodeFound) + 3, 100))
            '''' EDIT HERE - Check/Change cohortSize and responseThreshold
            ' --> Find threshold/cohort size
            'Set rngFindValue = refWs.Range(gstrCourseLookupRng).Find(What:=courseCodeFound, After:=refWs.Range(Left(gstrCourseLookupRng, InStr(1, gstrCourseLookupRng, ":") - 1)), LookIn:=xlValues)
            'rngFindRow = rngFindValue.Row
            'cohortSize = refWs.Cells(rngFindRow, gintCourseLookupReturn).Value2
            'responseThreshold = getResponseThreshold(cohortSize)
            ' --> Add a new course responses sheet
            courseWb.Sheets.Add Before:=Worksheets(Worksheets.count)
            Sheets(Worksheets.count - 1).Name = nameFound
            nameFound = courseCodeFound
            Set newCourseWs = ActiveSheet
            newCourseWs.Range("A1") = courseTitleFound
            'newCourseWs.Range("B1") = cohortSize
            'newCourseWs.Range("C1") = responseThreshold
            courseWs.Range("A" & starterRow & ":CE" & (currentRow - 1)).Copy Destination:=newCourseWs.Range("A" & lastRow)
            ' Then get ready for new course data
            starterRow = currentRow
            courseCount = 1
        Else
            ' First time course found
            courseCount = courseCount + 1
            nameFound = courseCodeFound
            courseTitleFound = Trim(Mid(cellContents, Len(courseCodeFound) + 3, 100))
        End If
        cellContents = Cell
    End If
Next
 
'''' START FINAL COURSE (with cohort size/threshold checks)!
courseRangeText = "A" & starterRow & ":CE" & (currentRow + 1)
courseCodeFound = Trim(Left(cellContents, InStr(1, cellContents, "-") - 1))
courseTitleFound = Trim(Mid(cellContents, Len(courseCodeFound) + 3, 100))

' --> Find threshold/cohort size
'Set rngFindValue = refWs.Range(gstrCourseLookupRng).Find(What:=courseCodeFound, After:=refWs.Range(Left(gstrCourseLookupRng, InStr(1, gstrCourseLookupRng, ":") - 1)), LookIn:=xlValues)
'rngFindRow = rngFindValue.Row
'cohortSize = refWs.Cells(rngFindRow, gintCourseLookupReturn).Value2
'responseThreshold = getResponseThreshold(cohortSize)
ActiveWorkbook.Sheets.Add Before:=Worksheets(Worksheets.count)
Sheets(Worksheets.count - 1).Name = courseCodeFound
Set newCourseWs = ActiveSheet
newCourseWs.Range("A1") = courseTitleFound
'newCourseWs.Range("B1") = cohortSize
'newCourseWs.Range("C1") = responseThreshold
courseWs.Range("A" & starterRow & ":CE" & (currentRow + 1)).Copy Destination:=newCourseWs.Range("A" & lastRow)
'''' END FINAL COURSE

' Save and end routine
courseWb.SaveAs FileName:=gstrReportsFilePath & "Split Course Report Sheets (" & Format(Date, "dd-mm-yy") & " " & Format(Time, "hh.mm.ss") & ").xlsx", FileFormat:=xlOpenXMLWorkbook

If Not gblnDebugging Then
    Application.ScreenUpdating = True
End If
Call endTimer

End Sub

'----------------------------
' @name splitModuleSheets
' @descr Splits the responses sheet into separate sheets for each module (by code) and copies student responses
' @usage Run on the RESPONSES sheet to split into module sheets for reporting
'----------------------------
Sub splitModuleSheets()

Call startTimer

If gblnDebugging Then
    Debug.Print "----"
    Debug.Print "---- START Split MODULE Sheets by CODE"
    Debug.Print "----"
End If
Application.ScreenUpdating = False

' Setup ORIG workbook/sheet - duplicated for reports
Dim origWb As Workbook
Dim gstrOrigWbFile As String

Dim moduleWb As Workbook
Dim moduleWs As Worksheet

Dim EA As Excel.Application
Set EA = New Excel.Application

' Check ORIG workbook open and set to origWb/origWs
'If gblnDebugging Then
'    Debug.Print "OrigWB Open Test - " & gstrOrigWbFile
'End If
'fileBoo = IsWorkBookOpen(gstrOrigWbFile)
'If fileBoo = True Then
    Set origWb = Workbooks(gstrOrigWbName)     'File is open
'Else
'    Set origWb = Workbooks.Open(gstrOrigWbFile) 'File is Closed
'End If
'Dim origWs As Worksheet
Set origWs = origWb.Sheets(gstrOrigWsName)

' Reference Workbook contains cohort sizes etc
If gblnDebugging Then
    Debug.Print "RefWB Open Test - " & gstrRefWbFile
End If
fileBoo = IsWorkBookOpen(gstrRefWbFile)
If fileBoo = True Then
    Set refWb = Workbooks(gstrRefWbName)
Else
    Set refWb = Workbooks.Open(gstrRefWbFile)
End If
' In the MODULES pre-processing macro
Set refWs = refWb.Sheets("MODULES")

' Now set up copied courseWb/courseWs
origWs.Copy
Set moduleWb = ActiveWorkbook
Set moduleWs = moduleWb.ActiveSheet
moduleWs.Name = "Module Reports"
' Adding summary data sheet
moduleWb.Sheets.Add Before:=Worksheets(Worksheets.count)
moduleWb.Sheets(Worksheets.count - 1).Name = "Summary Data"
With moduleWb.Sheets("Summary Data")
    .Cells(1, 1) = "Module Code"
    .Cells(1, 2) = "Module Title"
    .Cells(1, 3) = "Cohort Size"
    .Cells(1, 4) = "Response Rate (%)"
    .Cells(1, 5) = "Average Satisfaction"
    .Cells(1, 6) = "Median Satisfaction"
    .Cells(1, 7) = "Valid Responses"
    .Cells(1, 8) = "FHEQ Level"
    .Cells(1, 9) = "Published Flag"
    .Cells(1, 10) = "Department"
    .Cells(1, 11) = "School"
End With

finalRow = moduleWs.Range("A" & Rows.count).End(xlUp).Row
totalStudentsToProcess = finalRow - gintHeaderRows
modulesRangeText = "L3:S" & finalRow & ",AA3:AA" & finalRow
Set modulesRange = moduleWs.Range(modulesRangeText)

starterRow = 3                      ' starts at row 3, rows 1+2 are headers
nameFound = ""                      ' set to the name of the "found" course
moduleCount = 1                     ' increases when courses found in sorted list
currentRow = starterRow - 1
NewSheet = True

For Each Cell In modulesRange
    thisRow = Cell.Row
    moduleFound = Cell.Text
    
    If Not moduleFound = "" Then
        moduleCodeFound = Trim(Left(moduleFound, InStr(1, moduleFound, "-") - 2))
        moduleTitleFound = Trim(Mid(moduleFound, Len(moduleCodeFound) + 3, 100))
        NewSheet = True
        For Each Sheet In moduleWb.Sheets
            If Sheet.Name = moduleCodeFound Then
                 NewSheet = False
            End If
        Next
        If NewSheet Then
            moduleWb.Sheets.Add Before:=Worksheets(Worksheets.count)
            moduleWb.Sheets(Worksheets.count - 1).Name = moduleCodeFound
            If gblnDebugging Then
                Debug.Print "--> Creating new module sheet for " & moduleCodeFound
            End If
            Set newModuleWs = moduleWb.Sheets(moduleCodeFound)
            lastRow = gintSheetDataRows + 1
             ' Find threshold/cohort size
            Set rngFindValue = refWs.Range(gstrModuleLookupRng).Find(What:=moduleCodeFound, After:=refWs.Range("B8"), LookIn:=xlValues)
            rngFindRow = rngFindValue.Row
            cohortSize = refWs.Cells(rngFindRow, gintModuleLookupReturn).Value
            'responseThreshold = Application.Max(4, Round(cohortSize / 2, 0))
            responseThreshold = getResponseThreshold(cohortSize)
            newModuleWs.Range("A1") = moduleTitleFound         ' NOTE: not sanitised
            newModuleWs.Range("B1") = cohortSize
            newModuleWs.Range("C1") = responseThreshold
            
        Else
            Set newModuleWs = moduleWb.Sheets(moduleCodeFound)
            lastRow = newModuleWs.Cells(Rows.count, "A").End(xlUp).Row + 1
        End If
        moduleWs.Range("A" & thisRow & ":CD" & thisRow).Copy Destination:=newModuleWs.Range("A" & lastRow)
        If gblnDebugging Then
            'Debug.Print "Added student (#" & (thisRow - 2) & "/" & totalStudentsToProcess & ") responses for " & moduleCodeFound
        End If
        
    End If
Next

If gblnDebugging Then
    Debug.Print "----"
    Debug.Print "---- END Split MODULE Sheets by CODE"
    Debug.Print "----"
End If

Application.ScreenUpdating = True
moduleWb.SaveAs FileName:=gstrReportsFilePath & "\Split Module Report Sheets (" & Format(Date, "dd-mm-yy") & " " & Format(Time, "hh.mm.ss") & ").xlsx", FileFormat:=xlOpenXMLWorkbook

refWb.Close False
origWb.Close False
Set EA = Nothing

Call endTimer

End Sub

'----------------------------
' @name processStudentModules
' @descr Puts the student data (from Chris Anderson) into Qualtrics format (modules as columns for piped text)
' @usage Run on the student data sheet to turn rows of modules per student into columns
'----------------------------
Sub processStudentModules()

Dim fullRange As Range
Dim modulesRange As Range
Application.ScreenUpdating = False

maxModules = 11
starterRow = 2
nameFound = ""

endCell = Range("E" & Rows.count).End(xlUp).Address
finalRow = Range("E" & Rows.count).End(xlUp).Row
startCell = Range("A1").Address
Set fullRange = Range(startCell & ":" & endCell)

Columns("F:P").Select
Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

For Index = 1 To maxModules
    Range("F1:P1").Cells(Index) = "Module " & Index
Next
headerRow = starterRow

For thisRow = starterRow To finalRow
    nameCheck = Range("A" & thisRow).Value
    If nameCheck <> "" Then
        ' Name Row Found!
        headerRow = thisRow
        For Index = 1 To maxModules
            lowerCell = Range("A" & thisRow + Index)
            If lowerCell <> "" Then
                footerRow = thisRow + Index - 1
                Exit For
            End If
        Next
        Set modulesRange = Range("E" & headerRow & ":E" & footerRow)
        modulesRange.Select
        RowCount = footerRow - headerRow
        Selection.Copy
        Range("F" & headerRow).Select
        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=True
    Else
        ' Module Row Found!
    End If
Next

End Sub

'----------------------------
' @name tidyStudentModules
' @descr Removes empty rows from student data
' @usage Run AFTER processStudentModules to clean up sheet
'----------------------------
Sub tidyStudentModules()

Application.ScreenUpdating = False
Dim fullRange As Range
Set fullRange = Range("A1:S100")
Sheets(ActiveSheet.Name).Select
For Each Rng In fullRange
    If (WorksheetFunction.CountA(Cells(Rng.Row, 1)) = 0) Then
        Sheets(ActiveSheet.Name).Rows(Rng.Row).Delete
    End If
Next
Application.ScreenUpdating = True

End Sub
