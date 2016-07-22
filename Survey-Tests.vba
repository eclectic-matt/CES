' -------------------------
' NOT BEING USED AT PRESENT
' NOT "PRODUCTION" VERSIONS
' -------------------------
Dim reportingWorkbook As Workbook
Dim courseTitleToProcess As String

Sub sanitiseAllModuleSheets()

Call startTimer

Dim reportingWorkbook As Workbook
Set reportingWorkbook = ActiveWorkbook
Dim moduleCodeToProcess As String

For Each Sheet In ActiveWorkbook.Sheets
    moduleCodeToProcess = Sheet.Name
    If (Not moduleCodeToProcess = "Module Reports") And (Not moduleCodeToProcess = "Summary Data") Then
        If gblnDebugging Then
            Debug.Print "---------------------------"
            Debug.Print "Sanitising - " & moduleCodeToProcess
            Debug.Print "---------------------------"
        End If
        Call sanitiseForChecking(moduleCodeToProcess)
    End If
Next

Call endTimer

End Sub

Sub sanitiseForChecking(sheetName As String)

With Application.ActiveWorkbook.Sheets(sheetName)
    endRow = .Range("K" & Rows.count).End(xlUp).Row
    Debug.Print endRow
    .Range("B2:CE" & (endRow + 1)).Delete
    .Rows((endRow + 1) & ":" & (2 * endRow) + 2).EntireRow.Delete
    .Cells(1, 1).Value = "RespondentID"
    .Cells(1, 2).Value = "Best Comments"
    .Cells(1, 3).Value = "Worst Comments"
    .Cells(1, 4).Value = "Action Taken"
    .Columns("A:D").AutoFit
    .Columns("B:C").ColumnWidth = 60
    .Range("A1:D1").Font.Bold = True
    With .Range("A2:D" & endRow)
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
End With
End Sub

Sub showCountsMsgbox()
countRange = Application.Selection
counts = Array(0, 0, 0, 0)
responseCount = Selection.count
If responseCount = 1 Then
    MsgBox ("You have a single cell containing " & countRange)
Else
    For Each Cell In countRange
        Select Case Cell
            Case 1:
                counts(0) = counts(0) + 1
            Case 2:
                counts(1) = counts(1) + 1
            Case 3:
                counts(2) = counts(2) + 1
            Case 4:
                counts(3) = counts(3) + 1
            Case Else:
                elseCount = elseCount + 1
        End Select
    Next

    ValidResponses = responseCount - elseCount
    If ValidResponses = 0 Then
        ' Don't calculate percentages!
        zeroCnt = 0
        oneCnt = 0
        twoCnt = 0
        thrCnt = 0
    Else
        zeroCnt = Format(counts(0) / ValidResponses, "0.0%")
        oneCnt = Format(counts(1) / ValidResponses, "0.0%")
        twoCnt = Format(counts(2) / ValidResponses, "0.0%")
        thrCnt = Format(counts(3) / ValidResponses, "0.0%")
    End If
    
    MsgBox ("Showing counts in the selection (" & responseCount & " total)" & vbNewLine & _
        "Counted 1s = " & counts(0) & " (" & zeroCnt & ")" & vbNewLine & _
        "Counted 2s = " & counts(1) & " (" & oneCnt & ")" & vbNewLine & _
        "Counted 3s = " & counts(2) & " (" & twoCnt & ")" & vbNewLine & _
        "Counted 4s = " & counts(3) & " (" & thrCnt & ")")
End If

End Sub

Sub CopyOutJointCourses()

Dim Rng As Range
Dim WorkRng As Range
Dim off As Integer

maxRowsToCheck = 5

On Error Resume Next

xTitleId = "Fill Down Cells Tool"

Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Select the range to fill down" & vbNewLine & "i.e. if a cell has a value, copy this to the cell below to allow filters to work correctly", xTitleId, WorkRng.Address, Type:=8)

For Each Rng In WorkRng
    Select Case (Rng)
        Case "MAJOR"
            Debug.Print "Major in Cell " & Rng.Address
            ' Nowt - skip
            
        Case "JOINT"
            ' Count the number of joints below (to the next School in Col D)
            Debug.Print "Joint in Cell " & Rng.Address & " with owning School = " & Rng.Offset(0, -1)
            For off = 1 To maxRowsToCheck
                SchoolBelow = Rng.Offset(off, -1).Text
                Debug.Print SchoolBelow
                startRow = Rng.Row
                Debug.Print "Ready to copy from Row " & startRow
                If Len(SchoolBelow) = 0 Then
                    'Copy this row out
                        
                Else
                    'Stop at the row above!
                    endRow = Rng.Offset(off, -1).Row
                    Debug.Print "To Row " & endRow
                End If
            Next
            'ActiveSheet.Range("E" & startRow & ":F" & endRow).Select
            'Selection.Copy
            'ActiveSheet.Range("F" & startRow).PasteSpecial Transpose:=True
        Case Else
            Debug.Print "ELSE in Cell " & Rng.Address
            ' Nowt - skip

    End Select
Next Rng
End Sub

Sub restofcopy()
    
    If Rng.Text = "JOINT" Then
        Debug.Print "JOINT found on row " & Rng.Row
        Set copyRng = ActiveSheet.Range(Rng.Offset(1, 0), Rng.Offset(1, 1))
        copyRng.Copy
        'Cells(Rng.Row, Rng.Column + 2).Select
        ActiveSheet.Paste Destination:=ActiveSheet.Range(Cells(Rng.Row, Rng.Column + 2), Cells(Rng.Row, Rng.Column + 2))
        'Selection.PasteSpecial
        'Rows(Rng.Row + 1).Delete
    End If
Next Rng

MsgBox ("END - Filled Down")

End Sub


Sub COPYgenerateAllSplitYearCourseReports()

Call startTimer

current = 0
total = ActiveWorkbook.Sheets.count - 1     'Note: Course Reports sheet ignored
Call ShowProgressForm
Dim EA As Excel.Application
Set EA = New Excel.Application

Set reportingWorkbook = ActiveWorkbook

For Each Sheet In ActiveWorkbook.Sheets
    If Not Sheet.Name = "Course Reports" Then
        courseTitleToProcess = Sheet.Name
        If gblnDebugging Then
            Debug.Print "---------------------------"
            Debug.Print "Generating - " & Sheet.Name
            Debug.Print "---------------------------"
        End If
        current = current + 1
        Call UpdateProgressBar(current, total)
        'Worksheets(Sheet.Name).Activate
        Call generateAllSplitYearCourseReports(reportingWorkbook, courseTitleToProcess)
    End If
Next

Call Unload(ProgressForm)
Call endTimer
Set EA = Nothing
' Then save the Split Module Sheet (processed within sheets for record)

End Sub

Sub COPYtestReport()
    Set reportingWorkbook = ActiveWorkbook
    courseTitleToProcess = ActiveWorkbook.Sheets(ActiveSheet.Name)
    Call generateAllSplitYearCourseReports(reportingWorkbook, courseTitleToProcess)
End Sub

Sub TESTgenerateSplitYearCourseReport(reportWb As Workbook, courseCodeToProcess As String)

If gblnDebugging Then
    Debug.Print "----"
    Debug.Print "---- START Generate Split Year Course Reports"
    Debug.Print "----"
    Debug.Print "1) Setting Up"
End If

'Dim EA As Excel.Application
'Set EA = New Excel.Application

Dim repWb As Workbook
Dim repWs As Worksheet

Dim repWsName As String
Set repWb = reportWb
repWsName = courseCodeToProcess
'Debug.Print repWsName
Set repWs = repWb.Worksheets(repWsName)
'Set repWs = repWb.ActiveSheet

Dim refWb As Workbook
Dim refWs As Worksheet

responseCount = Range("E" & Rows.count).End(xlUp).Row - 1
courseCode = repWs.Name         'ActiveSheet.Name
courseTitle = courseCode & " - " & repWs.Range("A1").Text
If gblnDebugging Then
    Debug.Print "       TOTAL (all years) responses: " & responseCount
    Debug.Print "       Reporting for " & courseTitle & " with " & responseCount & " responses"
End If

' USING refWb AS REFERENCE FOR COHORT SIZES
fileBoo = False                                 ' Testing this as issues with open files
fileBoo = IsWorkBookOpen(gstrRefWbFile)
If fileBoo = True Then
    Set refWb = Application.Workbooks(gstrRefWbName)         ' File is open
Else
    'Set refWb = Workbooks.Open(gstrRefWbFile)     ' File is Closed
    Set refWb = Workbooks.Open(FileName:=gstrRefWbFile, IgnoreReadOnlyRecommended:=True, UpdateLinks:=0, ReadOnly:=True)
End If
Set refWs = refWb.Sheets("COURSES")

' USING origWb AS REFERENCE FOR QUESTION TEXT
fileBoo = False                                 ' Testing this as issues with open files
fileBoo = IsWorkBookOpen(gstrOrigWbFile)
If fileBoo = True Then
    Set origWb = Application.Workbooks(gstrOrigWbName)       ' File is open
Else
    Set origWb = Workbooks.Open(gstrOrigWbFile)     ' File is Closed
    'Set origWb = Workbooks.Open(FileName:=gstrOrigWbFile, IgnoreReadOnlyRecommended:=True, UpdateLinks:=0, ReadOnly:=True)      'Causing errors?
End If

If gblnDebugging Then
    Debug.Print "2) Check cohort size and threshold disclaimer"
End If

cohortSize = repWs.Range("B1").Text
responseRate = Round((responseCount / cohortSize) * 100, 2) & "%"
responseThreshold = repWs.Range("C1").Text
repWs.Range("A2:CE" & responseCount + 1).Sort key1:=repWs.Range("$CE:$CE"), Order2:=xlAscending ', Orientation:=xlTopToBottom, SortMethod:=xlPinYin, Header:=xlYes, MatchCase:=False,

If gblnDebugging Then
    Debug.Print "3) Getting course statistical data"
End If

studyYearCol = "$CE" & 2 & ":$CE" & (responseCount + 2)
studyYears = repWs.Range(studyYearCol)

YearFound = -100
Dim StudyYearsToProcess(0 To 4) As Integer      ' an array from 0 - 4 (the main UG study years) whose values are the # of respondents for that year
yearCounter = 1                                 ' Number of students in each year
thisRow = 1
'PWDcount = 0
For Each Cell In studyYears
    thisRow = thisRow + 1
    If Cell = YearFound Then
        ' Same year
        yearCounter = yearCounter + 1
    Else
        If YearFound = -100 Then
            ' Don't process, just start new
            YearFound = Cell
        Else
            ' Different year - Process OLD
            repWs.Rows(thisRow & ":" & (thisRow + (gintStatRows - 1))).EntireRow.Insert
            thisRow = thisRow + gintStatRows
            StudyYearsToProcess(YearFound) = yearCounter
            YearFound = Cell
            yearCounter = 1
        End If
    End If
Next

startRow = 2
Dim counts

For a = 0 To UBound(StudyYearsToProcess)
    If Not StudyYearsToProcess(a) = 0 Then
        ' Generate statistical data for this course
        responseCount = StudyYearsToProcess(a)
        counts = Array(0, 0, 0, 0)
        
        For Index = 0 To gintCourseDataColCount
            'If responseCount = 1 Then
            '    ' Ensure range in A1:A1 format (not just A1)
            '    cellA = Cells(startRow, gintCourseDataStartCol + Index).Address(False, False)
            '    cellB = Cells(startRow + responseCount - 1, gintCourseDataStartCol + Index).Address(False, False)
            '    Debug.Print "statRange = " & cellA & ":" & cellB
            '    Set statRange = repWs.Range(cellA & ":" & cellB)
            'Else
            '    ' Just use the usual stat range
            '    Set statRange = repWs.Range(Cells(startRow, gintCourseDataStartCol + Index), Cells(startRow + responseCount - 1, gintCourseDataStartCol + Index))
            '    'statRange = Range(Cells(startRow, gintCourseDataStartCol + Index), Cells(startRow + responseCount - 1, gintCourseDataStartCol + Index)).Address(False, False)
            'End If
            cellA = Cells(startRow, gintCourseDataStartCol + Index).Address(False, False)
            cellB = Cells(startRow + responseCount - 1, gintCourseDataStartCol + Index).Address(False, False)
            Debug.Print "statRange = " & cellA & ":" & cellB
            Set statRange = repWs.Range(cellA & ":" & cellB)
            
            ' MEDIAN AND AVERAGE *CANNOT* HANDLE SINGLE CELLS, SO SKIP
            If responseCount = 1 Then
                Median = statRange.Value2
                Average = statRange.Value2
            Else
                'CUSTOM - Median = getMedian(repWs.Range(statRange))
                'StatRange AS STRING - Median = Application.WorksheetFunction.Median(repWs.Range(statRange))
                Median = Application.WorksheetFunction.Median(statRange)
                'Average = Round(getAverage(repWs.Range(statRange)), 2)    'Average = Round(Application.WorksheetFunction.Average(repWs.Range(statRange)), 2)
                Average = Round(Application.WorksheetFunction.Average(statRange), 2)
                'StdDev = Round(getStdDevSam(ActiveSheet.Range(statRange)), 2)
            End If
            
            counts = Array(0, 0, 0, 0)
            elseCount = 0
            'For Each Cell In repWs.Range(statRange)
            For Each Cell In statRange
                Select Case Cell
                    Case 1:
                        counts(0) = counts(0) + 1
                    Case 2:
                        counts(1) = counts(1) + 1
                    Case 3:
                        counts(2) = counts(2) + 1
                    Case 4:
                        counts(3) = counts(3) + 1
                    Case Else:
                        elseCount = elseCount + 1
                End Select
            Next
            
            ' Using valid responses for calculations and using responseCount for row manipulation
            ValidResponses = responseCount - elseCount
            If ValidResponses = 0 Then
                ' Don't calculate percentages!
                zeroCnt = 0
                oneCnt = 0
                twoCnt = 0
                thrCnt = 0
            Else
                zeroCnt = Format(counts(0) / ValidResponses, "0.0%")
                oneCnt = Format(counts(1) / ValidResponses, "0.0%")
                twoCnt = Format(counts(2) / ValidResponses, "0.0%")
                thrCnt = Format(counts(3) / ValidResponses, "0.0%")
            End If
            ' If 6 responses (in Rows 2,3,4,5,6,7) then start printing stat data at Row 8
            repWs.Cells(startRow + responseCount, gintCourseDataStartCol + Index) = zeroCnt
            repWs.Cells(startRow + responseCount + 1, gintCourseDataStartCol + Index) = oneCnt
            repWs.Cells(startRow + responseCount + 2, gintCourseDataStartCol + Index) = twoCnt
            repWs.Cells(startRow + responseCount + 3, gintCourseDataStartCol + Index) = thrCnt
            repWs.Cells(startRow + responseCount + 4, gintCourseDataStartCol + Index) = ValidResponses
            repWs.Cells(startRow + responseCount + 5, gintCourseDataStartCol + Index) = Average
            repWs.Cells(startRow + responseCount + 6, gintCourseDataStartCol + Index) = Median
            'Cells(startRow + responseCount + 7, gintCourseDataStartCol + Index) = StdDev
            'repWs.Range(Cells(startRow + responseCount, gintCourseDataStartCol + Index), Cells(startRow + responseCount + gintStatRows - 1, gintCourseDataStartCol + Index)).Interior.Color = RGB(216, 215, 216)
            
        Next
        
        'Now ready to process next year
        firstCell = "A" & startRow
        lastCell = "CE" & startRow + StudyYearsToProcess(a) - 1
        'Debug.Print "For Year " & a & " there are " & StudyYearsToProcess(a) & " students who have responded in " & firstCell & ":" & lastCell
        startRow = startRow + StudyYearsToProcess(a) + gintStatRows
        
    End If
Next

' PART 2 - CREATING A COURSE REPORT
If gblnDebugging Then
    Debug.Print "4) Create Word Doc"
End If
Dim wrdApp As Word.Application
Dim wrdDoc As Word.document
Set wrdApp = CreateObject("Word.Application")
wrdApp.Visible = False

''''' WORKING HERE!!!
startRow = 2    'Then startRow = responseCount + startRow + gintStatRows
For StudyYear = 0 To 4
    'Debug.Print "Checking Year " & studyYear
    If Not StudyYearsToProcess(StudyYear) = 0 Then
        'Debug.Print "Processing Year " & studyYear
        Set rngFindValue = refWs.Range(gstrCourseLookupRng).Find(What:=repWs.Name, After:=refWs.Range(Left(gstrCourseLookupRng, InStr(1, gstrCourseLookupRng, ":") - 1)), LookIn:=xlValues)
        rngFindRow = rngFindValue.Row
        cohortSize = refWs.Cells(rngFindRow, 4 + StudyYear).Value2
        responseThreshold = getResponseThreshold(cohortSize)
        responseCount = StudyYearsToProcess(StudyYear)
        responseRate = Round((responseCount / cohortSize) * 100, 2) & "%"
        Debug.Print "RESPONSES = " & responseCount & " so RESPRATE = " & responseRate
        If responseCount < responseThreshold Then
            disclaimer = gstrThresholdDisclaimer
            disclaimer = Replace(disclaimer, "%RESP", responseCount)
            disclaimer = Replace(disclaimer, "$THRE", responseThreshold)
        Else
            disclaimer = ""
        End If
        
        Set wrdDoc = wrdApp.Documents.Add

        ' HEADERS AND FOOTERS AND PAGE SETUP
        With wrdDoc
            .Sections(1).Headers(wdHeaderFooterPrimary).Range.Text = gstrDocTitle & " " & gstrDocYear
            .Sections(1).Footers(wdHeaderFooterPrimary).Range.Text = "CES Report for Year " & StudyYear & ", " & courseTitle & " (generated: " & Format(Date, "dd-mm-yy") & " " & Format(Time, "hh.mm.ss") & ")"
            .Sections(1).Footers(wdHeaderFooterPrimary).PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberRight, FirstPage:=True
            .PageSetup.Orientation = wdOrientLandscape
        End With
        
        ' SET DOCUMENT STYLES
        With wrdDoc
                With .Styles(wdStyleHeading1).Font
                    .Name = "Arial"
                    .Size = 16
                    .Bold = True
                    .Color = wdColorBlack
                End With
                With .Styles(wdStyleHeading2).Font
                    .Name = "Arial"
                    .Size = 12
                    .Bold = True
                    .Color = wdColorBlack
                End With
                With .Styles(wdStyleHeading3).Font
                    .Name = "Arial"
                    .Size = 10
                    .Bold = True
                    .Color = wdColorBlack
                End With
                With .Styles(wdStyleHeading6).Font
                    .Name = "Arial"
                    .Size = 12
                    .Bold = True
                    .Color = vbRed
                End With
                With .Styles(wdStyleNormal).Font
                    .Name = "Arial"
                    .Size = 10
                    .Color = wdColorBlack
                End With
        End With
        
        If gblnDebugging Then
            Debug.Print "5) Add course summary data"
        End If
        
        With wrdDoc
            
        ' Titles and course details
            .Range(0).Style = .Styles(wdStyleHeading1)
            .Content.InsertAfter gstrDocTitle & " " & gstrDocYear
            .Content.InsertParagraphAfter
            
            .Range(.Characters.count - 1).Style = .Styles(wdStyleHeading2)
            .Content.InsertAfter "COURSE-LEVEL REPORT FOR " & UCase(courseTitle) & " (YEAR " & StudyYear & ")"
            .Content.InsertParagraphAfter
            
            .Range(.Characters.count - 1).Style = .Styles(wdStyleNormal)
            .Content.InsertParagraphAfter
            .Content.InsertAfter "Eligible Cohort Size: " & cohortSize
            .Content.InsertParagraphAfter
            .Content.InsertAfter "Number of responses: " & responseCount
            .Content.InsertParagraphAfter
            .Content.InsertAfter "Response Rate: " & responseRate
            .Content.InsertParagraphAfter
            
            If Not disclaimer = "" Then
                .Range(.Characters.count - 1).Style = .Styles(wdStyleHeading6)
                .Content.InsertAfter disclaimer
                .Content.InsertParagraphAfter
                .Content.InsertParagraphAfter
            End If
            
            .Range(.Characters.count - 1).InlineShapes.AddHorizontalLineStandard
            .Content.InsertParagraphAfter
        End With
        
        If gblnDebugging Then
            Debug.Print "6) Add course satisfaction tables"
        End If
            
        ' ------------------------------------
        '       COURSE SATISFACTION TABLES
        ' ------------------------------------
        
        ' ------- Overall Course Satisfaction START
        wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading2)
        wrdDoc.Content.InsertAfter "1) Overall Course Satisfaction"
        wrdDoc.Content.InsertParagraphAfter
        
        ' Table Headings
        'headerRow = Array("Question Text", "Not at all satisfied (1)", "Not very satisfied (2)", "Quite satisfied (3)", "Very satisfied (4)", "Total responses", "Mean", "Median", "Standard Deviation")
        headerRow = parseStrToArr(gstrCourseSatisfactionHeadings, ",")
        TableColumns = UBound(headerRow) - LBound(headerRow) + 1
        
        Set myRange = wrdDoc.Range(Start:=wrdDoc.Characters.count - 1, End:=wrdDoc.Characters.count)
        Set overallTable = wrdDoc.Tables.Add(Range:=myRange, NumRows:=2, NumColumns:= _
            TableColumns, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
            wdAutoFitFixed)
        overallTable.Style = "Grid Table 1 Light"
        overallTable.ApplyStyleHeadingRows = True
        
        ' (TableColumns - 1) as filling all cells and 0-index
        For Index = 0 To (TableColumns - 1)
            Set tableCell = overallTable.Cell(1, Index + 1)
            tableCell.Range.InsertAfter Text:=headerRow(Index)
        Next
        
        ' Question Title(s)
        'QTitle = Cells(1, gintCourseDataStartCol)
        QText = origWb.Worksheets("REFERENCE").Range("C" & gintCourseDataStartCol + 1).Text
        If gblnDebugging Then
            Debug.Print "       Question: " & QText
        End If
        overallTable.Cell(2, 1).Range.InsertAfter Text:=QText
        overallTable.Columns(1).SetWidth ColumnWidth:=gsngTotalPageWidthPoints / 2, RulerStyle:=wdAdjustNone
        
        ' Response Data
        ' (TableColumns - 2) as filling all but "Question Text" cell and 0-index
        For Index = 0 To (TableColumns - 2)
            Set tableCell = overallTable.Cell(2, Index + 2)
            overallTable.Columns(Index + 2).SetWidth ColumnWidth:=(gsngTotalPageWidthPoints / 2) / (TableColumns - 1), RulerStyle:=wdAdjustNone
            cellData = repWs.Cells(startRow + responseCount + Index, gintCourseDataStartCol).Text
            tableCell.Range.InsertAfter Text:=cellData
        Next
        ' ------- Overall Course Satisfaction END
        '
        wrdDoc.Range(wrdDoc.Characters.count - 1).InsertParagraphAfter
        '
        ' ------- Course Content table START
        wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading2)
        wrdDoc.Content.InsertAfter "2) Course Content"
        wrdDoc.Content.InsertParagraphAfter
        ' Table Headings
        headerRow = parseStrToArr(gstrCourseContentHeadings, ",")
        'headerRow = Array("Question Text", "Extremely (1)", "Moderately (2)", "Slightly (3)", "Not at all (4)", "Total responses", "Mean", "Median", "Standard Deviation")
        TableColumns = UBound(headerRow) - LBound(headerRow) + 1
        
        Set myRange = wrdDoc.Range(Start:=wrdDoc.Characters.count - 1, End:=wrdDoc.Characters.count)
        Set courseContentTable = wrdDoc.Tables.Add(Range:=myRange, NumRows:=3, NumColumns:= _
            TableColumns, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
            wdAutoFitFixed)
        courseContentTable.Style = "Grid Table 1 Light"
        courseContentTable.ApplyStyleHeadingRows = True
        
        For Index = 0 To (TableColumns - 1)
            Set tableCell = courseContentTable.Cell(1, Index + 1)
            tableCell.Range.InsertAfter Text:=headerRow(Index)
        Next
        
        ' ----- Getting Response Data
        For questionNo = 1 To 2
            ' Question Title(s)
            QText = origWb.Worksheets("REFERENCE").Range("C" & gintCourseDataStartCol + questionNo + 1).Text
            If gblnDebugging Then
                Debug.Print "       Question: " & QText
            End If
            courseContentTable.Cell(1 + questionNo, 1).Range.InsertAfter Text:=QText
            courseContentTable.Columns(1).SetWidth ColumnWidth:=(gsngTotalPageWidthPoints * 0.5), RulerStyle:=wdAdjustNone
            For Index = 0 To (TableColumns - 2)
                Set tableCell = courseContentTable.Cell(1 + questionNo, Index + 2)
                courseContentTable.Columns(Index + 2).SetWidth ColumnWidth:=((0.5 * gsngTotalPageWidthPoints) / 7), RulerStyle:=wdAdjustNone
                cellData = repWs.Cells(startRow + responseCount + Index, gintCourseDataStartCol + questionNo).Text
                ''' POTENTIAL FATAL ERROR (EXCEL CRASHES) BELOW!!!!
                If (Index = 4) And (CStr(cellData) <> CStr(responseCount)) Then
                    cellData = cellData & "**"
                End If
                tableCell.Range.InsertAfter Text:=cellData
            Next
        Next
        ' ------- Course Content table END
        '
        wrdDoc.Range(wrdDoc.Characters.count - 1).InsertBreak (wdPageBreak)
        '
        ' ------- Assessment table START
        wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading2)
        wrdDoc.Content.InsertAfter "3) Course Assessment"
        wrdDoc.Content.InsertParagraphAfter
        ' Table Headings
        'headerRow = Array("Question Text", "Excellent (1)", "Good (2)", "Fair (3)", "Poor (4)", "Total responses", "Mean", "Median")
        headerRow = parseStrToArr(gstrCourseAssessmentHeadings, ",")
        TableColumns = UBound(headerRow) - LBound(headerRow) + 1
        
        Set myRange = wrdDoc.Range(Start:=wrdDoc.Characters.count - 1, End:=wrdDoc.Characters.count)
        Set assessmentTable = wrdDoc.Tables.Add(Range:=myRange, NumRows:=4, NumColumns:= _
            TableColumns, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
            wdAutoFitFixed)
        assessmentTable.Style = "Grid Table 1 Light"
        assessmentTable.ApplyStyleHeadingRows = True
        
        For Index = 0 To (TableColumns - 1)
            Set tableCell = assessmentTable.Cell(1, Index + 1)
            tableCell.Range.InsertAfter Text:=headerRow(Index)
        Next
        
        ' ----- Getting Response Data
        FirstQuestion = 3
        RowOffset = 2 - FirstQuestion
        For questionNo = FirstQuestion To 5
        ' Question Title(s)
        QText = origWb.Worksheets("REFERENCE").Range("C" & gintCourseDataStartCol + questionNo + 1).Text
        If gblnDebugging Then
            Debug.Print "       Question: " & QText
        End If
        assessmentTable.Cell(RowOffset + questionNo, 1).Range.InsertAfter Text:=QText
        assessmentTable.Columns(1).SetWidth ColumnWidth:=gsngTotalPageWidthPoints / 2, RulerStyle:=wdAdjustNone
            For Index = 0 To (TableColumns - 2)
                Set tableCell = assessmentTable.Cell(RowOffset + questionNo, Index + 2)
                assessmentTable.Columns(Index + 2).SetWidth ColumnWidth:=(gsngTotalPageWidthPoints / 2) / (TableColumns - 1), RulerStyle:=wdAdjustNone
                cellData = repWs.Cells(startRow + responseCount + Index, gintCourseDataStartCol + questionNo).Text
                tableCell.Range.InsertAfter Text:=cellData
            Next
        Next
        ' ------- Assessment table END
        '
        wrdDoc.Range(wrdDoc.Characters.count - 1).InsertParagraphAfter
        '
        ' ------- Workload table START !!!NOTE - 3 response columns!!!
        wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading2)
        wrdDoc.Content.InsertAfter "4) Course Workload"
        wrdDoc.Content.InsertParagraphAfter
        ' Table Headings
        'headerRow = Array("Question Text", "Too much (1)", "About right (2)", "Too little (3)", "Total responses", "Mean", "Median")
        headerRow = parseStrToArr(gstrCourseWorkloadHeadings, ",")
        TableColumns = UBound(headerRow) - LBound(headerRow) + 1
        
        Set myRange = wrdDoc.Range(Start:=wrdDoc.Characters.count - 1, End:=wrdDoc.Characters.count)
        Set assessmentTable = wrdDoc.Tables.Add(Range:=myRange, NumRows:=3, NumColumns:=TableColumns, _
            DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
            wdAutoFitFixed)
        assessmentTable.Style = "Grid Table 1 Light"
        assessmentTable.ApplyStyleHeadingRows = True
        
        For Index = 0 To (TableColumns - 1)
            Set tableCell = assessmentTable.Cell(1, Index + 1)
            tableCell.Range.InsertAfter Text:=headerRow(Index)
        Next
        
        ' ----- Getting Response Data
        FirstQuestion = 6
        RowOffset = 2 - FirstQuestion
        For questionNo = FirstQuestion To 7
        ' Question Title(s)
        QText = origWb.Worksheets("REFERENCE").Range("C" & gintCourseDataStartCol + questionNo + 1).Text
        If gblnDebugging Then
            Debug.Print "       Question: " & QText
        End If
        assessmentTable.Cell(RowOffset + questionNo, 1).Range.InsertAfter Text:=QText
        assessmentTable.Columns(1).SetWidth ColumnWidth:=gsngTotalPageWidthPoints / 2, RulerStyle:=wdAdjustNone
            For Index = 0 To (TableColumns - 1)
                ' SKIP OVER INDEX 3 (AS NO FOURTH RESPONSE OPTION FOR THESE QUESTIONS)
                If Not Index = 3 Then
                    If Index > 3 Then
                        Ind = Index - 1
                    Else
                        Ind = Index
                    End If
                    Set tableCell = assessmentTable.Cell(RowOffset + questionNo, Ind + 2)
                    assessmentTable.Columns(Ind + 2).SetWidth ColumnWidth:=(gsngTotalPageWidthPoints / 2) / (TableColumns - 1), RulerStyle:=wdAdjustNone
                    cellData = repWs.Cells(startRow + responseCount + Index, gintCourseDataStartCol + questionNo).Text
                    tableCell.Range.InsertAfter Text:=cellData
                End If
            Next
        Next
        ' ------- Workload table END
        '
        wrdDoc.Range(wrdDoc.Characters.count - 1).InsertParagraphAfter
        '
        ' ------- Skills table START
        wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading2)
        wrdDoc.Content.InsertAfter "5) Skills Development"
        wrdDoc.Content.InsertParagraphAfter
        ' Table Headings
        'headerRow = Array("Question Text", "Excellent (1)", "Good (2)", "Fair (3)", "Poor (4)", "Total responses", "Mean", "Median")
        headerRow = parseStrToArr(gstrCourseSkillsHeadings, ",")
        TableColumns = UBound(headerRow) - LBound(headerRow) + 1
        
        Set myRange = wrdDoc.Range(Start:=wrdDoc.Characters.count - 1, End:=wrdDoc.Characters.count)
        Set skillsTable = wrdDoc.Tables.Add(Range:=myRange, NumRows:=7, NumColumns:=TableColumns, _
            DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
            wdAutoFitFixed)
        skillsTable.Style = "Grid Table 1 Light"
        skillsTable.ApplyStyleHeadingRows = True
        
        
        For Index = 0 To (TableColumns - 1)
            Set tableCell = skillsTable.Cell(1, Index + 1)
            tableCell.Range.InsertAfter Text:=headerRow(Index)
        Next
        
        ' ----- Getting Response Data
        FirstQuestion = 8
        RowOffset = 2 - FirstQuestion
        For questionNo = FirstQuestion To 13
        ' Question Title(s)
        QText = origWb.Worksheets("REFERENCE").Range("C" & gintCourseDataStartCol + questionNo + 1).Text
        If gblnDebugging Then
            Debug.Print "       Question: " & QText
        End If
        skillsTable.Cell(RowOffset + questionNo, 1).Range.InsertAfter Text:=QText
        skillsTable.Columns(1).SetWidth ColumnWidth:=gsngTotalPageWidthPoints / 2, RulerStyle:=wdAdjustNone
            For Index = 0 To (TableColumns - 2)
                Set tableCell = skillsTable.Cell(RowOffset + questionNo, Index + 2)
                skillsTable.Columns(Index + 2).SetWidth ColumnWidth:=(gsngTotalPageWidthPoints / 2) / (TableColumns - 1), RulerStyle:=wdAdjustNone
                cellData = repWs.Cells(startRow + responseCount + Index, gintCourseDataStartCol + questionNo).Text
                tableCell.Range.InsertAfter Text:=cellData
            Next
        Next
        ' ------- Skills table END
        '
        'wrdDoc.Range(wrdDoc.Characters.count - 1).InsertParagraphAfter
        wrdDoc.Range(wrdDoc.Characters.count - 1).InsertBreak (wdPageBreak)
        '
        ' ------- Prep table START
        wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading2)
        wrdDoc.Content.InsertAfter "6) Preparation for Life and Work"
        wrdDoc.Content.InsertParagraphAfter
        
        ' Table Headings
        'headerRow = Array("Question Text", "Extremely (1)", "Moderately (2)", "Slightly (3)", "Not at all (4)", "Total responses", "Mean", "Median")
        headerRow = parseStrToArr(gstrCoursePreparationHeadings, ",")
        TableColumns = UBound(headerRow) - LBound(headerRow) + 1
        
        Set myRange = wrdDoc.Range(Start:=wrdDoc.Characters.count - 1, End:=wrdDoc.Characters.count)
        Set prepTable = wrdDoc.Tables.Add(Range:=myRange, NumRows:=2, NumColumns:=TableColumns, _
            DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
            wdAutoFitFixed)
        prepTable.Style = "Grid Table 1 Light"
        prepTable.ApplyStyleHeadingRows = True
        
        For Index = 0 To (TableColumns - 1)
            Set tableCell = prepTable.Cell(1, Index + 1)
            tableCell.Range.InsertAfter Text:=headerRow(Index)
        Next
        
        ' ----- Getting Response Data
        FirstQuestion = 14
        RowOffset = 2 - FirstQuestion
        questionNo = 14     'For questionNo = FirstQuestion To 14
        ' Question Title(s)
        QText = origWb.Worksheets("REFERENCE").Range("C" & gintCourseDataStartCol + questionNo + 1).Text
        If gblnDebugging Then
            Debug.Print "       Question: " & QText
        End If
        prepTable.Cell(RowOffset + questionNo, 1).Range.InsertAfter Text:=QText
        prepTable.Columns(1).SetWidth ColumnWidth:=gsngTotalPageWidthPoints / 2, RulerStyle:=wdAdjustNone
            For Index = 0 To (TableColumns - 2)
                Set tableCell = prepTable.Cell(RowOffset + questionNo, Index + 2)
                prepTable.Columns(Index + 2).SetWidth ColumnWidth:=(gsngTotalPageWidthPoints / 2) / (TableColumns - 1), RulerStyle:=wdAdjustNone
                cellData = repWs.Cells(startRow + responseCount + Index, gintCourseDataStartCol + questionNo).Text
                tableCell.Range.InsertAfter Text:=cellData
            Next
        ' ------- Prep table END
        '
        wrdDoc.Range(wrdDoc.Characters.count - 1).InsertParagraphAfter
        '
        ' ------- Resources table START
        wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading2)
        wrdDoc.Content.InsertAfter "7) Course Resources/Equipment"
        wrdDoc.Content.InsertParagraphAfter
        ' Table Headings
        'headerRow = Array("Question Text", "Excellent (1)", "Good (2)", "Fair (3)", "Poor (4)", "Total responses", "Mean", "Median")
        headerRow = parseStrToArr(gstrCourseResourcesHeadings, ",")
        TableColumns = UBound(headerRow) - LBound(headerRow) + 1
        
        Set myRange = wrdDoc.Range(Start:=wrdDoc.Characters.count - 1, End:=wrdDoc.Characters.count)
        Set resourcesTable = wrdDoc.Tables.Add(Range:=myRange, NumRows:=5, NumColumns:=TableColumns, _
            DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
            wdAutoFitFixed)
        resourcesTable.Style = "Grid Table 1 Light"
        resourcesTable.ApplyStyleHeadingRows = True
        
        For Index = 0 To (TableColumns - 1)
            Set tableCell = resourcesTable.Cell(1, Index + 1)
            tableCell.Range.InsertAfter Text:=headerRow(Index)
        Next
        ' ----- Getting Response Data
        FirstQuestion = 15
        RowOffset = 2 - FirstQuestion
        For questionNo = FirstQuestion To 18
        ' Question Title(s)
        QText = origWb.Worksheets("REFERENCE").Range("C" & gintCourseDataStartCol + questionNo + 1).Text
        If gblnDebugging Then
            Debug.Print "       Question: " & QText
        End If
        resourcesTable.Cell(RowOffset + questionNo, 1).Range.InsertAfter Text:=QText
        resourcesTable.Columns(1).SetWidth ColumnWidth:=gsngTotalPageWidthPoints / 2, RulerStyle:=wdAdjustNone
            For Index = 0 To (TableColumns - 2)
                Set tableCell = resourcesTable.Cell(RowOffset + questionNo, Index + 2)
                resourcesTable.Columns(Index + 2).SetWidth ColumnWidth:=(gsngTotalPageWidthPoints / 2) / (TableColumns - 1), RulerStyle:=wdAdjustNone
                cellData = repWs.Cells(startRow + responseCount + Index, gintCourseDataStartCol + questionNo).Text
                tableCell.Range.InsertAfter Text:=cellData
            Next
        Next
        ' ------- Resources table END
        '
        wrdDoc.Range(wrdDoc.Characters.count - 1).InsertParagraphAfter
        '
        ' ------- Organisation table START
        wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading2)
        wrdDoc.Content.InsertAfter "8) Course Organisation"
        wrdDoc.Content.InsertParagraphAfter
        ' Table Headings
        'headerRow = Array("Question Text", "Strongly agree (1)", "Slightly agree (2)", "Slightly disagree (3)", "Strongly disagree (4)", "Total responses", "Mean", "Median")
        headerRow = parseStrToArr(gstrCourseOrganisationHeadings, ",")
        TableColumns = UBound(headerRow) - LBound(headerRow) + 1
        
        Set myRange = wrdDoc.Range(Start:=wrdDoc.Characters.count - 1, End:=wrdDoc.Characters.count)
        Set organisationTable = wrdDoc.Tables.Add(Range:=myRange, NumRows:=4, NumColumns:=TableColumns, _
            DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
            wdAutoFitFixed)
        organisationTable.Style = "Grid Table 1 Light"
        organisationTable.ApplyStyleHeadingRows = True
        
        For Index = 0 To (TableColumns - 1)
            Set tableCell = organisationTable.Cell(1, Index + 1)
            tableCell.Range.InsertAfter Text:=headerRow(Index)
        Next
        ' ----- Getting Response Data
        FirstQuestion = 19
        RowOffset = 2 - FirstQuestion
        For questionNo = FirstQuestion To 21
        ' Question Title(s)
        QText = origWb.Worksheets("REFERENCE").Range("C" & gintCourseDataStartCol + questionNo + 1).Text
        If gblnDebugging Then
            Debug.Print "       Question: " & QText
        End If
        organisationTable.Cell(RowOffset + questionNo, 1).Range.InsertAfter Text:=QText
        organisationTable.Columns(1).SetWidth ColumnWidth:=gsngTotalPageWidthPoints / 2, RulerStyle:=wdAdjustNone
            For Index = 0 To (TableColumns - 2)
                Set tableCell = organisationTable.Cell(RowOffset + questionNo, Index + 2)
                organisationTable.Columns(Index + 2).SetWidth ColumnWidth:=(gsngTotalPageWidthPoints / 2) / (TableColumns - 1), RulerStyle:=wdAdjustNone
                cellData = repWs.Cells(startRow + responseCount + Index, gintCourseDataStartCol + questionNo).Text
                tableCell.Range.InsertAfter Text:=cellData
            Next
        Next
        ' ------- Organisation table END
        '
        'wrdDoc.Range(wrdDoc.Characters.count - 1).InsertParagraphAfter
        wrdDoc.Range(wrdDoc.Characters.count - 1).InsertBreak (wdPageBreak)
        '
        ' ------- Comments table START
        If gblnDebugging Then
            Debug.Print "       Free Text Comments"
        End If
        wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading2)
        wrdDoc.Content.InsertAfter "9) Free Text Comments"
        wrdDoc.Content.InsertParagraphAfter
        
        Debug.Print "--- COMMENTS IN Range(CA" & startRow & ":CA" & startRow + responseCount - 1 & ")"
        Set commentsRange = repWs.Range("CA" & startRow & ":CA" & startRow + responseCount - 1)
        If commentsRange.count > 1 Then
            commentsRange.RemoveDuplicates
        End If
        
        commentsCount = 0
        For Each Cell In commentsRange
            If Not Trim(Cell) = "" Then
                commentsCount = commentsCount + 1
                If gblnDebugging Then
                    Debug.Print commentsCount & " = " & Cell
                End If
            End If
        Next
        Set myRange = wrdDoc.Range(Start:=wrdDoc.Characters.count - 1, End:=wrdDoc.Characters.count)
        Set commentsTable = wrdDoc.Tables.Add(Range:=myRange, NumRows:=commentsCount + 1, NumColumns:=1, _
            DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
            wdAutoFitFixed)
        'commentsTable.Style = "List Table 2"
        Set tableCell = commentsTable.Cell(1, 1)
        tableCell.Range.InsertAfter Text:="Are there any additional comments you would like to include?" & vbNewLine
        commentsTable.Borders.Enable = wdLineStyleSingle
        
        commentNo = 1
        For Each Cell In commentsRange
            If Not Trim(Cell) = "" Then
                Set tableCell = commentsTable.Cell(commentNo + 1, 1)
                tableCell.Range.InsertAfter Text:=Cell
                commentNo = commentNo + 1
            End If
        Next
        commentsTable.Style = "Grid Table 1 Light"
        commentsTable.ApplyStyleHeadingRows = True
        ' ------- Comments table END
        
        ' Shade all header rows to light grey
        For Index = 1 To wrdDoc.Tables.count
            'wrdDoc.Tables(Index).Rows(1).Shading.Texture = wdTexture12Pt5Percent
            With wrdDoc.Tables(Index)
                rCount = .Rows.count
                For rNum = 1 To rCount
                    If rNum Mod 2 = 1 Then
                        .Rows(rNum).Shading.Texture = wdTexture12Pt5Percent
                    Else
                        .Rows(rNum).Shading.Texture = wdTextureNone
                    End If
                Next
            End With
        Next
        
        ' Output report as saved PDF document
        wrdDoc.Activate
        sanitisedCourseTitle = Replace(Replace(Replace(Replace(Left(courseTitle, 30), ":", " "), "?", ""), "(", ""), ")", "")
        oName = gstrReportsFilePath & "COURSE REPORTS\" & sanitisedCourseTitle & " YEAR " & StudyYear & " [" & Format(Date, "dd-mm-yy") & " " & Format(Time, "hh.mm.ss") & "].pdf"
        wrdDoc.SaveAs2 FileName:=oName, FileFormat:=wdFormatPDF
        wrdDoc.Close (False)
        If gblnDebugging Then
            Debug.Print "Saved Report as PDF - " & oName
        End If
        
        ' ALL COURSE REPORTING HERE
        'NOTE: Update CellData formulas to include startRow (per year)
        
        
        ' AFTER REPORTING - Ready for next course year
        startRow = startRow + responseCount + gintStatRows
    End If
    
Next

wrdApp.Quit (False)
Set wrdApp = Nothing
'Set EA = Nothing
Debug.Print "-------------------------------"
Debug.Print "COMPLETE - " & courseTitle
Debug.Print "-------------------------------"

End Sub

Sub cutThisIntoLoopAbove()
Set wrdDoc = wrdApp.Documents.Add

' HEADERS AND FOOTERS AND PAGE SETUP
With wrdDoc
    .Sections(1).Headers(wdHeaderFooterPrimary).Range.Text = gstrDocTitle & " " & gstrDocYear
    .Sections(1).Footers(wdHeaderFooterPrimary).Range.Text = "CES Report for Year " & StudyYear & ", " & courseTitle & " (generated: " & Format(Date, "dd-mm-yy") & " " & Format(Time, "hh.mm.ss") & ")"
    .PageSetup.Orientation = wdOrientLandscape
End With

' SET DOCUMENT STYLES
With wrdDoc
        With .Styles(wdStyleHeading1).Font
            .Name = "Arial"
            .Size = 16
            .Bold = True
            .Color = wdColorBlack
        End With
        With .Styles(wdStyleHeading2).Font
            .Name = "Arial"
            .Size = 12
            .Bold = True
            .Color = wdColorBlack
        End With
        With .Styles(wdStyleHeading3).Font
            .Name = "Arial"
            .Size = 10
            .Bold = True
            .Color = wdColorBlack
        End With
        With .Styles(wdStyleHeading6).Font
            .Name = "Arial"
            .Size = 12
            .Bold = True
            .Color = vbRed
        End With
        With .Styles(wdStyleNormal).Font
            .Name = "Arial"
            .Size = 10
            .Color = wdColorBlack
        End With
End With

If gblnDebugging Then
    Debug.Print "5) Add course summary data"
End If

With wrdDoc
    
' Titles and course details
    .Range(0).Style = .Styles(wdStyleHeading1)
    .Content.InsertAfter gstrDocTitle & " " & gstrDocYear
    .Content.InsertParagraphAfter
    
    .Range(.Characters.count - 1).Style = .Styles(wdStyleHeading2)
    .Content.InsertAfter "COURSE-LEVEL REPORT FOR " & UCase(courseTitle)
    .Content.InsertParagraphAfter
    
    .Range(.Characters.count - 1).Style = .Styles(wdStyleNormal)
    .Content.InsertParagraphAfter
    .Content.InsertAfter "Eligible Cohort Size: " & cohortSize
    .Content.InsertParagraphAfter
    .Content.InsertAfter "Number of responses: " & responseCount
    .Content.InsertParagraphAfter
    .Content.InsertAfter "Response Rate: " & responseRate
    .Content.InsertParagraphAfter
    
    If Not disclaimer = "" Then
        .Range(.Characters.count - 1).Style = .Styles(wdStyleHeading6)
        .Content.InsertAfter disclaimer
        .Content.InsertParagraphAfter
        .Content.InsertParagraphAfter
    End If
    
    .Range(.Characters.count - 1).InlineShapes.AddHorizontalLineStandard
    .Content.InsertParagraphAfter
End With

If gblnDebugging Then
    Debug.Print "6) Add course satisfaction tables"
End If
    
' ------------------------------------
'       COURSE SATISFACTION TABLES
' ------------------------------------

' ------- Overall Course Satisfaction START
wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading2)
wrdDoc.Content.InsertAfter "Overall Course Satisfaction"
wrdDoc.Content.InsertParagraphAfter

' Table Headings
'headerRow = Array("Question Text", "Not at all satisfied (1)", "Not very satisfied (2)", "Quite satisfied (3)", "Very satisfied (4)", "Total responses", "Mean", "Median", "Standard Deviation")
headerRow = parseStrToArr(gstrCourseSatisfactionHeadings, ",")
TableColumns = UBound(headerRow) - LBound(headerRow) + 1

Set myRange = wrdDoc.Range(Start:=wrdDoc.Characters.count - 1, End:=wrdDoc.Characters.count)
Set overallTable = wrdDoc.Tables.Add(Range:=myRange, NumRows:=2, NumColumns:= _
    TableColumns, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
    wdAutoFitFixed)
overallTable.Style = "Grid Table 1 Light"
overallTable.ApplyStyleHeadingRows = True

' (TableColumns - 1) as filling all cells and 0-index
For Index = 0 To (TableColumns - 1)
    Set tableCell = overallTable.Cell(1, Index + 1)
    tableCell.Range.InsertAfter Text:=headerRow(Index)
Next

' Question Title(s)
'QTitle = Cells(1, gintCourseDataStartCol)
QText = origWb.Worksheets("REFERENCE").Range("C" & gintCourseDataStartCol + 1).Text
If gblnDebugging Then
    Debug.Print "       Question: " & QText
End If
overallTable.Cell(2, 1).Range.InsertAfter Text:=QText
overallTable.Columns(1).SetWidth ColumnWidth:=gsngTotalPageWidthPoints / 2, RulerStyle:=wdAdjustNone

' Response Data
' (TableColumns - 2) as filling all but "Question Text" cell and 0-index
For Index = 0 To (TableColumns - 2)
    Set tableCell = overallTable.Cell(2, Index + 2)
    overallTable.Columns(Index + 2).SetWidth ColumnWidth:=(gsngTotalPageWidthPoints / 2) / (TableColumns - 1), RulerStyle:=wdAdjustNone
    cellData = repWs.Cells(startRow + responseCount + Index, gintCourseDataStartCol).Text
    tableCell.Range.InsertAfter Text:=cellData
Next
' ------- Overall Course Satisfaction END
'
wrdDoc.Range(wrdDoc.Characters.count - 1).InsertParagraphAfter
'
' ------- Course Content table START
wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading2)
wrdDoc.Content.InsertAfter "Course Content"
wrdDoc.Content.InsertParagraphAfter
' Table Headings
headerRow = parseStrToArr(gstrCourseContentHeadings, ",")
'headerRow = Array("Question Text", "Extremely (1)", "Moderately (2)", "Slightly (3)", "Not at all (4)", "Total responses", "Mean", "Median", "Standard Deviation")
TableColumns = UBound(headerRow) - LBound(headerRow) + 1

Set myRange = wrdDoc.Range(Start:=wrdDoc.Characters.count - 1, End:=wrdDoc.Characters.count)
Set courseContentTable = wrdDoc.Tables.Add(Range:=myRange, NumRows:=3, NumColumns:= _
    TableColumns, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
    wdAutoFitFixed)
courseContentTable.Style = "Grid Table 1 Light"
courseContentTable.ApplyStyleHeadingRows = True

For Index = 0 To (TableColumns - 1)
    Set tableCell = courseContentTable.Cell(1, Index + 1)
    tableCell.Range.InsertAfter Text:=headerRow(Index)
Next

' ----- Getting Response Data
For questionNo = 1 To 2
    ' Question Title(s)
    QText = origWb.Worksheets("REFERENCE").Range("C" & gintCourseDataStartCol + questionNo + 1).Text
    If gblnDebugging Then
        Debug.Print "       Question: " & QText
    End If
    courseContentTable.Cell(1 + questionNo, 1).Range.InsertAfter Text:=QText
    courseContentTable.Columns(1).SetWidth ColumnWidth:=gsngTotalPageWidthPoints * 0.36, RulerStyle:=wdAdjustNone
    
    For Index = 0 To (TableColumns - 2)
        Set tableCell = courseContentTable.Cell(1 + questionNo, Index + 2)
        courseContentTable.Columns(Index + 2).SetWidth ColumnWidth:=(0.08 * gsngTotalPageWidthPoints), RulerStyle:=wdAdjustNone
        cellData = repWs.Cells(startRow + responseCount + Index, gintCourseDataStartCol + questionNo).Text
        tableCell.Range.InsertAfter Text:=cellData
    Next
Next
' ------- Course Content table END
'
wrdDoc.Range(wrdDoc.Characters.count - 1).InsertBreak (wdPageBreak)
'
' ------- Assessment table START
wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading2)
wrdDoc.Content.InsertAfter "Course Assessment"
wrdDoc.Content.InsertParagraphAfter
' Table Headings
'headerRow = Array("Question Text", "Excellent (1)", "Good (2)", "Fair (3)", "Poor (4)", "Total responses", "Mean", "Median")
headerRow = parseStrToArr(gstrCourseAssessmentHeadings, ",")
TableColumns = UBound(headerRow) - LBound(headerRow) + 1

Set myRange = wrdDoc.Range(Start:=wrdDoc.Characters.count - 1, End:=wrdDoc.Characters.count)
Set assessmentTable = wrdDoc.Tables.Add(Range:=myRange, NumRows:=4, NumColumns:= _
    TableColumns, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
    wdAutoFitFixed)
assessmentTable.Style = "Grid Table 1 Light"
assessmentTable.ApplyStyleHeadingRows = True

For Index = 0 To (TableColumns - 1)
    Set tableCell = assessmentTable.Cell(1, Index + 1)
    tableCell.Range.InsertAfter Text:=headerRow(Index)
Next

' ----- Getting Response Data
FirstQuestion = 3
RowOffset = 2 - FirstQuestion
For questionNo = FirstQuestion To 5
' Question Title(s)
QText = origWb.Worksheets("REFERENCE").Range("C" & gintCourseDataStartCol + questionNo + 1).Text
If gblnDebugging Then
    Debug.Print "       Question: " & QText
End If
assessmentTable.Cell(RowOffset + questionNo, 1).Range.InsertAfter Text:=QText
assessmentTable.Columns(1).SetWidth ColumnWidth:=gsngTotalPageWidthPoints / 2, RulerStyle:=wdAdjustNone
    For Index = 0 To (TableColumns - 2)
        Set tableCell = assessmentTable.Cell(RowOffset + questionNo, Index + 2)
        assessmentTable.Columns(Index + 2).SetWidth ColumnWidth:=(gsngTotalPageWidthPoints / 2) / (TableColumns - 1), RulerStyle:=wdAdjustNone
        cellData = repWs.Cells(startRow + responseCount + Index, gintCourseDataStartCol + questionNo).Text
        tableCell.Range.InsertAfter Text:=cellData
    Next
Next
' ------- Assessment table END
'
wrdDoc.Range(wrdDoc.Characters.count - 1).InsertParagraphAfter
'
' ------- Workload table START !!!NOTE - 3 response columns!!!
wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading2)
wrdDoc.Content.InsertAfter "Course Workload"
wrdDoc.Content.InsertParagraphAfter
' Table Headings
'headerRow = Array("Question Text", "Too much (1)", "About right (2)", "Too little (3)", "Total responses", "Mean", "Median")
headerRow = parseStrToArr(gstrCourseWorkloadHeadings, ",")
TableColumns = UBound(headerRow) - LBound(headerRow) + 1

Set myRange = wrdDoc.Range(Start:=wrdDoc.Characters.count - 1, End:=wrdDoc.Characters.count)
Set assessmentTable = wrdDoc.Tables.Add(Range:=myRange, NumRows:=3, NumColumns:=TableColumns, _
    DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
    wdAutoFitFixed)
assessmentTable.Style = "Grid Table 1 Light"
assessmentTable.ApplyStyleHeadingRows = True

For Index = 0 To (TableColumns - 1)
    Set tableCell = assessmentTable.Cell(1, Index + 1)
    tableCell.Range.InsertAfter Text:=headerRow(Index)
Next

' ----- Getting Response Data
FirstQuestion = 6
RowOffset = 2 - FirstQuestion
For questionNo = FirstQuestion To 7
' Question Title(s)
QText = origWb.Worksheets("REFERENCE").Range("C" & gintCourseDataStartCol + questionNo + 1).Text
If gblnDebugging Then
    Debug.Print "       Question: " & QText
End If
assessmentTable.Cell(RowOffset + questionNo, 1).Range.InsertAfter Text:=QText
assessmentTable.Columns(1).SetWidth ColumnWidth:=gsngTotalPageWidthPoints / 2, RulerStyle:=wdAdjustNone
    For Index = 0 To (TableColumns - 1)
        ' SKIP OVER INDEX 3 (AS NO FOURTH RESPONSE OPTION FOR THESE QUESTIONS)
        If Not Index = 3 Then
            If Index > 3 Then
                Ind = Index - 1
            Else
                Ind = Index
            End If
            Set tableCell = assessmentTable.Cell(RowOffset + questionNo, Ind + 2)
            assessmentTable.Columns(Ind + 2).SetWidth ColumnWidth:=(gsngTotalPageWidthPoints / 2) / (TableColumns - 1), RulerStyle:=wdAdjustNone
            cellData = repWs.Cells(startRow + responseCount + Index, gintCourseDataStartCol + questionNo).Text
            tableCell.Range.InsertAfter Text:=cellData
        End If
    Next
Next
' ------- Workload table END
'
wrdDoc.Range(wrdDoc.Characters.count - 1).InsertParagraphAfter
'
' ------- Skills table START
wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading2)
wrdDoc.Content.InsertAfter "Skills Development"
wrdDoc.Content.InsertParagraphAfter
' Table Headings
'headerRow = Array("Question Text", "Excellent (1)", "Good (2)", "Fair (3)", "Poor (4)", "Total responses", "Mean", "Median")
headerRow = parseStrToArr(gstrCourseSkillsHeadings, ",")
TableColumns = UBound(headerRow) - LBound(headerRow) + 1

Set myRange = wrdDoc.Range(Start:=wrdDoc.Characters.count - 1, End:=wrdDoc.Characters.count)
Set skillsTable = wrdDoc.Tables.Add(Range:=myRange, NumRows:=7, NumColumns:=TableColumns, _
    DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
    wdAutoFitFixed)
skillsTable.Style = "Grid Table 1 Light"
skillsTable.ApplyStyleHeadingRows = True


For Index = 0 To (TableColumns - 1)
    Set tableCell = skillsTable.Cell(1, Index + 1)
    tableCell.Range.InsertAfter Text:=headerRow(Index)
Next

' ----- Getting Response Data
FirstQuestion = 8
RowOffset = 2 - FirstQuestion
For questionNo = FirstQuestion To 13
' Question Title(s)
QText = origWb.Worksheets("REFERENCE").Range("C" & gintCourseDataStartCol + questionNo + 1).Text
If gblnDebugging Then
    Debug.Print "       Question: " & QText
End If
skillsTable.Cell(RowOffset + questionNo, 1).Range.InsertAfter Text:=QText
skillsTable.Columns(1).SetWidth ColumnWidth:=gsngTotalPageWidthPoints / 2, RulerStyle:=wdAdjustNone
    For Index = 0 To (TableColumns - 2)
        Set tableCell = skillsTable.Cell(RowOffset + questionNo, Index + 2)
        skillsTable.Columns(Index + 2).SetWidth ColumnWidth:=(gsngTotalPageWidthPoints / 2) / (TableColumns - 1), RulerStyle:=wdAdjustNone
        cellData = repWs.Cells(startRow + responseCount + Index, gintCourseDataStartCol + questionNo).Text
        tableCell.Range.InsertAfter Text:=cellData
    Next
Next
' ------- Skills table END
'
'wrdDoc.Range(wrdDoc.Characters.count - 1).InsertParagraphAfter
wrdDoc.Range(wrdDoc.Characters.count - 1).InsertBreak (wdPageBreak)
'
' ------- Prep table START
wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading2)
wrdDoc.Content.InsertAfter "Preparation for Life and Work"
wrdDoc.Content.InsertParagraphAfter

' Table Headings
'headerRow = Array("Question Text", "Extremely (1)", "Moderately (2)", "Slightly (3)", "Not at all (4)", "Total responses", "Mean", "Median")
headerRow = parseStrToArr(gstrCoursePreparationHeadings, ",")
TableColumns = UBound(headerRow) - LBound(headerRow) + 1

Set myRange = wrdDoc.Range(Start:=wrdDoc.Characters.count - 1, End:=wrdDoc.Characters.count)
Set prepTable = wrdDoc.Tables.Add(Range:=myRange, NumRows:=2, NumColumns:=TableColumns, _
    DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
    wdAutoFitFixed)
prepTable.Style = "Grid Table 1 Light"
prepTable.ApplyStyleHeadingRows = True

For Index = 0 To (TableColumns - 1)
    Set tableCell = prepTable.Cell(1, Index + 1)
    tableCell.Range.InsertAfter Text:=headerRow(Index)
Next

' ----- Getting Response Data
FirstQuestion = 14
RowOffset = 2 - FirstQuestion
questionNo = 14     'For questionNo = FirstQuestion To 14
' Question Title(s)
QText = origWb.Worksheets("REFERENCE").Range("C" & gintCourseDataStartCol + questionNo + 1).Text
If gblnDebugging Then
    Debug.Print "       Question: " & QText
End If
prepTable.Cell(RowOffset + questionNo, 1).Range.InsertAfter Text:=QText
prepTable.Columns(1).SetWidth ColumnWidth:=gsngTotalPageWidthPoints / 2, RulerStyle:=wdAdjustNone
    For Index = 0 To (TableColumns - 2)
        Set tableCell = prepTable.Cell(RowOffset + questionNo, Index + 2)
        prepTable.Columns(Index + 2).SetWidth ColumnWidth:=(gsngTotalPageWidthPoints / 2) / (TableColumns - 1), RulerStyle:=wdAdjustNone
        cellData = repWs.Cells(startRow + responseCount + Index, gintCourseDataStartCol + questionNo).Text
        tableCell.Range.InsertAfter Text:=cellData
    Next
' ------- Prep table END
'
wrdDoc.Range(wrdDoc.Characters.count - 1).InsertParagraphAfter
'
' ------- Resources table START
wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading2)
wrdDoc.Content.InsertAfter "Course Resources/Equipment"
wrdDoc.Content.InsertParagraphAfter
' Table Headings
'headerRow = Array("Question Text", "Excellent (1)", "Good (2)", "Fair (3)", "Poor (4)", "Total responses", "Mean", "Median")
headerRow = parseStrToArr(gstrCourseResourcesHeadings, ",")
TableColumns = UBound(headerRow) - LBound(headerRow) + 1

Set myRange = wrdDoc.Range(Start:=wrdDoc.Characters.count - 1, End:=wrdDoc.Characters.count)
Set resourcesTable = wrdDoc.Tables.Add(Range:=myRange, NumRows:=5, NumColumns:=TableColumns, _
    DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
    wdAutoFitFixed)
resourcesTable.Style = "Grid Table 1 Light"
resourcesTable.ApplyStyleHeadingRows = True

For Index = 0 To (TableColumns - 1)
    Set tableCell = resourcesTable.Cell(1, Index + 1)
    tableCell.Range.InsertAfter Text:=headerRow(Index)
Next
' ----- Getting Response Data
FirstQuestion = 15
RowOffset = 2 - FirstQuestion
For questionNo = FirstQuestion To 18
' Question Title(s)
QText = origWb.Worksheets("REFERENCE").Range("C" & gintCourseDataStartCol + questionNo + 1).Text
If gblnDebugging Then
    Debug.Print "       Question: " & QText
End If
resourcesTable.Cell(RowOffset + questionNo, 1).Range.InsertAfter Text:=QText
resourcesTable.Columns(1).SetWidth ColumnWidth:=gsngTotalPageWidthPoints / 2, RulerStyle:=wdAdjustNone
    For Index = 0 To (TableColumns - 2)
        Set tableCell = resourcesTable.Cell(RowOffset + questionNo, Index + 2)
        resourcesTable.Columns(Index + 2).SetWidth ColumnWidth:=(gsngTotalPageWidthPoints / 2) / (TableColumns - 1), RulerStyle:=wdAdjustNone
        cellData = repWs.Cells(startRow + responseCount + Index, gintCourseDataStartCol + questionNo).Text
        tableCell.Range.InsertAfter Text:=cellData
    Next
Next
' ------- Resources table END
'
wrdDoc.Range(wrdDoc.Characters.count - 1).InsertParagraphAfter
'
' ------- Organisation table START
wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading2)
wrdDoc.Content.InsertAfter "Course Organisation"
wrdDoc.Content.InsertParagraphAfter
' Table Headings
'headerRow = Array("Question Text", "Strongly agree (1)", "Slightly agree (2)", "Slightly disagree (3)", "Strongly disagree (4)", "Total responses", "Mean", "Median")
headerRow = parseStrToArr(gstrCourseOrganisationHeadings, ",")
TableColumns = UBound(headerRow) - LBound(headerRow) + 1

Set myRange = wrdDoc.Range(Start:=wrdDoc.Characters.count - 1, End:=wrdDoc.Characters.count)
Set organisationTable = wrdDoc.Tables.Add(Range:=myRange, NumRows:=4, NumColumns:=TableColumns, _
    DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
    wdAutoFitFixed)
organisationTable.Style = "Grid Table 1 Light"
organisationTable.ApplyStyleHeadingRows = True

For Index = 0 To (TableColumns - 1)
    Set tableCell = organisationTable.Cell(1, Index + 1)
    tableCell.Range.InsertAfter Text:=headerRow(Index)
Next
' ----- Getting Response Data
FirstQuestion = 19
RowOffset = 2 - FirstQuestion
For questionNo = FirstQuestion To 21
' Question Title(s)
QText = origWb.Worksheets("REFERENCE").Range("C" & gintCourseDataStartCol + questionNo + 1).Text
If gblnDebugging Then
    Debug.Print "       Question: " & QText
End If
organisationTable.Cell(RowOffset + questionNo, 1).Range.InsertAfter Text:=QText
organisationTable.Columns(1).SetWidth ColumnWidth:=gsngTotalPageWidthPoints / 2, RulerStyle:=wdAdjustNone
    For Index = 0 To (TableColumns - 2)
        Set tableCell = organisationTable.Cell(RowOffset + questionNo, Index + 2)
        organisationTable.Columns(Index + 2).SetWidth ColumnWidth:=(gsngTotalPageWidthPoints / 2) / (TableColumns - 1), RulerStyle:=wdAdjustNone
        cellData = repWs.Cells(startRow + responseCount + Index, gintCourseDataStartCol + questionNo).Text
        tableCell.Range.InsertAfter Text:=cellData
    Next
Next
' ------- Organisation table END
'
'wrdDoc.Range(wrdDoc.Characters.count - 1).InsertParagraphAfter
wrdDoc.Range(wrdDoc.Characters.count - 1).InsertBreak (wdPageBreak)
'
' ------- Comments table START
If gblnDebugging Then
    Debug.Print "       Free Text Comments"
End If
wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading2)
wrdDoc.Content.InsertAfter "Free Text Comments"
wrdDoc.Content.InsertParagraphAfter

Set commentsRange = Range("CA1:CA" & responseCount + 1)
If commentsRange.count > 1 Then
    commentsRange.RemoveDuplicates
End If

commentsCount = 0
For Each Cell In commentsRange
    If Not Trim(Cell) = "" Then
        commentsCount = commentsCount + 1
        If gblnDebugging Then
            Debug.Print commentsCount & " = " & Cell
        End If
    End If
Next
Set myRange = wrdDoc.Range(Start:=wrdDoc.Characters.count - 1, End:=wrdDoc.Characters.count)
Set commentsTable = wrdDoc.Tables.Add(Range:=myRange, NumRows:=commentsCount + 1, NumColumns:=1, _
    DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
    wdAutoFitFixed)
'commentsTable.Style = "List Table 2"
Set tableCell = commentsTable.Cell(1, 1)
tableCell.Range.InsertAfter Text:="Are there any additional comments you would like to include?" & vbNewLine
commentsTable.Borders.Enable = wdLineStyleSingle

commentNo = 1
For Each Cell In commentsRange
    If Not Trim(Cell) = "" Then
        Set tableCell = commentsTable.Cell(commentNo + 1, 1)
        tableCell.Range.InsertAfter Text:=Cell
        commentNo = commentNo + 1
    End If
Next
commentsTable.Style = "Grid Table 1 Light"
commentsTable.ApplyStyleHeadingRows = True
' ------- Comments table END

' Shade all header rows to light grey
For Index = 1 To wrdDoc.Tables.count
    'wrdDoc.Tables(Index).Rows(1).Shading.Texture = wdTexture12Pt5Percent
    With wrdDoc.Tables(Index)
        rCount = .Rows.count
        For rNum = 1 To rCount
            If rNum Mod 2 = 1 Then
                .Rows(rNum).Shading.Texture = wdTexture12Pt5Percent
            Else
                .Rows(rNum).Shading.Texture = wdTextureNone
            End If
        Next
    End With
Next

' Output report as saved PDF document
wrdDoc.Activate
sanitisedCourseTitle = Replace(Replace(Replace(Replace(Left(courseTitle, 30), ":", " "), "?", ""), "(", ""), ")", "")
oName = gstrReportsFilePath & "COURSE REPORTS/" & sanitisedCourseTitle & " [" & Format(Date, "dd-mm-yy") & " " & Format(Time, "hh.mm.ss") & "].pdf"
wrdDoc.SaveAs2 FileName:=oName, FileFormat:=wdFormatPDF
wrdDoc.Close (False)
If gblnDebugging Then
    Debug.Print "Saved Report as PDF - " & oName
End If
wrdApp.Quit (False)
Set wrdApp = Nothing
Set EA = Nothing

' NOTE: No error handling at present (want errors to break)
errorDisp:
If Err Then
    Debug.Print "An error was found!"
    Debug.Print "Description: " & Err.Description
    Debug.Print "Source: " & Err.Source
    Debug.Print "Number: " & Err.Number
    Debug.Print "DLL Error: " & Err.LastDllError
    Debug.Print "Help Context: " & Err.HelpContext
    Debug.Print ""
End If

'MsgBox ("Course Report Generated for " & courseTitle)

End Sub


Sub testingMultipleCourseYears()

'Call startTimer
Application.ScreenUpdating = True

If gblnDebugging Then
    Debug.Print "----"
    Debug.Print "---- START Generate Course Reports"
    Debug.Print "----"
    Debug.Print "1) Setting Up"
End If

Dim EA As Excel.Application
Set EA = New Excel.Application
'Dim origWb As Workbook
Dim repWs As Worksheet
Set repWs = ActiveSheet

responseCount = repWs.Range("E" & Rows.count).End(xlUp).Row - 1
courseCode = ActiveSheet.Name
courseTitle = courseCode & " - " & ActiveSheet.Range("A1").Text
repWs.Range("A2:CE" & responseCount + 1).Sort key1:=repWs.Range("$CE:$CE"), Order2:=xlAscending, Header:=xlNo, MatchCase:=False, Orientation:=xlTopToBottom, SortMethod:=xlPinYin

If gblnDebugging Then
    Debug.Print "       Number of responses: " & responseCount
    Debug.Print "       Reporting for " & courseTitle & " with " & responseCount & " responses"
End If

'' USING origWb AS REFERENCE FOR QUESTION TEXT
'fileBoo = False                                 ' Testing this as issues with open files
'fileBoo = IsWorkBookOpen(gstrOrigWbFile)
'If fileBoo = True Then
'    Set origWb = Workbooks(gstrOrigWbName)         ' File is open
'Else
'    Set origWb = Workbooks.Open(gstrOrigWbFile)     ' File is Closed
'End If

If gblnDebugging Then
    Debug.Print "2) Check cohort size and threshold disclaimer"
End If

studyYearCol = "$CE" & 2 & ":$CE" & (responseCount + 2)
studyYears = repWs.Range(studyYearCol)

YearFound = -100
Dim StudyYearsToProcess(0 To 4) As Integer      ' an array from 0 - 4 (the main UG study years) whose values are the # of respondents for that year
yearCounter = 1                                 ' Number of students in each year
thisRow = 1
'PWDcount = 0
For Each Cell In studyYears
thisRow = thisRow + 1
    If Cell = YearFound Then
        ' Same year
        yearCounter = yearCounter + 1
    Else
        If YearFound = -100 Then
            ' Don't process, just start new
            YearFound = Cell
        Else
            ' Different year - Process OLD
                'Debug.Print "Splitting on Row " & thisRow & ":" & thisRow + (gintStatRows - 1)
                repWs.Rows(thisRow & ":" & (thisRow + (gintStatRows - 1))).EntireRow.Insert
                thisRow = thisRow + gintStatRows
                'Range(Cells(thisRow), Cells(thisRow)).EntireRow.Resize(6).Insert Shift:=xlDown
                'Rows(thisRow & ":" & thisRow + 10).Insert
                'Range(Cells(thisRow), Cells(thisRow)).EntireRow.Insert
            'If Cell = "PWD" Then
            '    PWDcount = PWDcount + 1
            'Else
                StudyYearsToProcess(YearFound) = yearCounter
                YearFound = Cell
                yearCounter = 1
            'End If
        End If
    End If
Next
' ^^^ NOTHING ABOVE HERE BREAKS THE ROWS!!!! ^^^

startRow = 2
Dim counts

For a = 0 To UBound(StudyYearsToProcess)
    If Not StudyYearsToProcess(a) = 0 Then
        ' Generate statistical data for this course
        responseCount = StudyYearsToProcess(a)
        counts = Array(0, 0, 0, 0)
        
        For Index = 0 To gintCourseDataColCount
            
            statRange = Range(Cells(startRow, gintCourseDataStartCol + Index), Cells(startRow + responseCount - 1, gintCourseDataStartCol + Index)).Address(False, False)
            
            'Median = getMedian(repWs.Range(statRange))                '
            Median = EA.WorksheetFunction.Median(repWs.Range(statRange))
            'Average = Round(getAverage(repWs.Range(statRange)), 2)    '
            Average = Round(EA.WorksheetFunction.Average(repWs.Range(statRange)), 2)
            'StdDev = Round(getStdDevSam(ActiveSheet.Range(statRange)), 2)
            
            counts = Array(0, 0, 0, 0)
            elseCount = 0
            For Each Cell In repWs.Range(statRange)
                Select Case Cell
                    Case 1:
                        counts(0) = counts(0) + 1
                    Case 2:
                        counts(1) = counts(1) + 1
                    Case 3:
                        counts(2) = counts(2) + 1
                    Case 4:
                        counts(3) = counts(3) + 1
                    Case Else:
                        elseCount = elseCount + 1
                End Select
            Next
            
            ' Using valid responses for calculations and using responseCount for row manipulation
            ValidResponses = responseCount - elseCount
            If ValidResponses = 0 Then
                ' Don't calculate percentages!
                zeroCnt = 0
                oneCnt = 0
                twoCnt = 0
                thrCnt = 0
            Else
                zeroCnt = Format(counts(0) / ValidResponses, "0.0%")
                oneCnt = Format(counts(1) / ValidResponses, "0.0%")
                twoCnt = Format(counts(2) / ValidResponses, "0.0%")
                thrCnt = Format(counts(3) / ValidResponses, "0.0%")
            End If
            ' If 6 responses (in Rows 2,3,4,5,6,7) then start printing stat data at Row 8
            repWs.Cells(startRow + responseCount, gintCourseDataStartCol + Index) = zeroCnt
            repWs.Cells(startRow + responseCount + 1, gintCourseDataStartCol + Index) = oneCnt
            repWs.Cells(startRow + responseCount + 2, gintCourseDataStartCol + Index) = twoCnt
            repWs.Cells(startRow + responseCount + 3, gintCourseDataStartCol + Index) = thrCnt
            repWs.Cells(startRow + responseCount + 4, gintCourseDataStartCol + Index) = ValidResponses
            repWs.Cells(startRow + responseCount + 5, gintCourseDataStartCol + Index) = Average
            repWs.Cells(startRow + responseCount + 6, gintCourseDataStartCol + Index) = Median
            'Cells(startRow + responseCount + 7, gintCourseDataStartCol + Index) = StdDev
            repWs.Range(Cells(startRow + responseCount, gintCourseDataStartCol + Index), Cells(startRow + responseCount + 6, gintCourseDataStartCol + Index)).Interior.Color = RGB(216, 215, 216)
        Next
        
        'Now ready to process next year
        firstCell = "A" & startRow
        lastCell = "CE" & startRow + StudyYearsToProcess(a) - 1
        'Debug.Print "For Year " & a & " there are " & StudyYearsToProcess(a) & " students who have responded in " & firstCell & ":" & lastCell
        startRow = startRow + StudyYearsToProcess(a) + gintStatRows
        
    End If
Next

'Call endTimer
End Sub

Sub restOfStudyYears()

cohortSize = repWs.Range("B1").Text
responseRate = Round((responseCount / cohortSize) * 100, 2) & "%"
responseThreshold = repWs.Range("C1").Text
If responseCount < responseThreshold Then
    disclaimer = gstrThresholdDisclaimer
    disclaimer = Replace(disclaimer, "%RESP", responseCount)
    disclaimer = Replace(disclaimer, "$THRE", responseThreshold)
Else
    disclaimer = ""
End If

If gblnDebugging Then
    Debug.Print "3) Getting course statistical data"
End If

Dim counts
counts = Array(0, 0, 0, 0)

For Index = 0 To gintCourseDataColCount
    
    statRange = ActiveSheet.Range(Cells(2, gintCourseDataStartCol + Index), Cells(responseCount + 1, gintCourseDataStartCol + Index)).Address(False, False)
    
    Median = getMedian(ActiveSheet.Range(statRange))                ' Median = EA.WorksheetFunction.Median(ActiveSheet.Range(statRange))
    Average = Round(getAverage(ActiveSheet.Range(statRange)), 2)    ' Average = Round(WorksheetFunction.Average(ActiveSheet.Range(statRange)), 2)
    StdDev = Round(getStdDevSam(ActiveSheet.Range(statRange)), 2)
    
    counts = Array(0, 0, 0, 0)
    elseCount = 0
    For Each Cell In ActiveSheet.Range(statRange)
        Select Case Cell
            Case 1:
                counts(0) = counts(0) + 1
            Case 2:
                counts(1) = counts(1) + 1
            Case 3:
                counts(2) = counts(2) + 1
            Case 4:
                counts(3) = counts(3) + 1
            Case Else:
                elseCount = elseCount + 1
        End Select
    Next
    
    ValidResponses = responseCount - elseCount
    ' If 6 responses (in Rows 2,3,4,5,6,7) then start printing stat data at Row 8
    Cells(responseCount + 2, gintCourseDataStartCol + Index) = Format(counts(0) / ValidResponses, "0.0%")
    Cells(responseCount + 3, gintCourseDataStartCol + Index) = Format(counts(1) / ValidResponses, "0.0%")
    Cells(responseCount + 4, gintCourseDataStartCol + Index) = Format(counts(2) / ValidResponses, "0.0%")
    Cells(responseCount + 5, gintCourseDataStartCol + Index) = Format(counts(3) / ValidResponses, "0.0%")
    Cells(responseCount + 6, gintCourseDataStartCol + Index) = ValidResponses
    Cells(responseCount + 7, gintCourseDataStartCol + Index) = Average
    Cells(responseCount + 8, gintCourseDataStartCol + Index) = Median
    Cells(responseCount + 9, gintCourseDataStartCol + Index) = StdDev
Next


End Sub

Sub testingSplit()

Debug.Print gstrCourseSatisfactionHeadings
newArr = parseStrToArr(gstrCourseSatisfactionHeadings, ",")
For a = 0 To UBound(newArr)
    Debug.Print a & "=" & newArr(a)
Next

End Sub

Sub splitCourseWithYear()

' Setup ORIG workbook/sheet - duplicated for reports
Dim origWb As Workbook
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

' Now set up copied courseWb/courseWs
origWs.Copy
Set courseWb = ActiveWorkbook
Set courseWs = courseWb.ActiveSheet

' Now sort this sheet by COURSE ready to sanitise
courseWs.Name = "Course Reports"
fullResultsRange = "A2:CE" & (gintTotalRecords + 1)
courseWs.Range(fullResultsRange).Sort key1:=courseWs.Range("$T:$T"), key2:=courseWs.Range("$CE:$CE"), Header:=xlYes, MatchCase:=False, Orientation:=xlTopToBottom, SortMethod:=xlPinYin


End Sub


Sub errorCourseSplits()
'''''''------- REWRITE THIS AS SOME COURSES HAVE (SLIGHTLY) CHANGED TITLE!
' Loop through the courses column separating out found courses
For Each Cell In coursesCol
    'thisRow = Cell.Row
    currentRow = currentRow + 1
    If Cell = nameFound Then
        ' Add to course count
        courseCount = courseCount + 1
    Else
        If Cell <> nameFound And nameFound <> "" Then
            
            ' Process previous course data
            courseCodeFound = Trim(Left(nameFound, InStr(1, nameFound, "-") - 1))
            courseTitleFound = Trim(Mid(nameFound, Len(courseCodeFound) + 3, 100))
             ' Find threshold/cohort size
            Set rngFindValue = refWs.Range(gstrCourseLookupRng).Find(What:=courseCodeFound, After:=refWs.Range(Left(gstrCourseLookupRng, InStr(1, gstrCourseLookupRng, ":") - 1)), LookIn:=xlValues)
            rngFindRow = rngFindValue.Row
            cohortSize = refWs.Cells(rngFindRow, gintCourseLookupReturn).Value
            ''''' IF BELOW THRESHOLD -> NOTE THIS ON SHEET TITLE?
            responseThreshold = Application.Max(4, Round(cohortSize / 2, 0))
            If gblnDebugging Then
                Debug.Print "       Orig Name: " & nameFound
                Debug.Print "       Edit Name: " & courseCodeFound
                Debug.Print "       Cohort Found: " & cohortSize
            End If
            courseWb.Sheets.Add Before:=Worksheets(Worksheets.count)
            Sheets(Worksheets.count - 1).Name = courseCodeFound
            Set newCourseWs = ActiveSheet
            newCourseWs.Range("A1") = courseTitleFound
            newCourseWs.Range("B1") = cohortSize
            newCourseWs.Range("C1") = responseThreshold
            courseWs.Range("A" & starterRow & ":CD" & (currentRow - 1)).Copy Destination:=newCourseWs.Range("A" & lastRow)
            nameFound = Cell
            'ActiveSheet.Paste
            
            ' Ready for new course data
            starterRow = currentRow
            courseCount = 1
        Else
            ' First course found
            courseCount = courseCount + 1
            nameFound = Cell
        End If
    End If
Next
'''''''------- REWRITE THIS AS SOME COURSES HAVE (SLIGHTLY) CHANGED TITLE!
End Sub

Sub OLDsplitModuleSheetsbyTitle()

If gblnDebugging Then
    Debug.Print "----"
    Debug.Print "---- START Split MODULE Sheets"
    Debug.Print "----"
End If
'Application.ScreenUpdating = False

' Setup ORIG workbook/sheet - duplicated for reports
Dim origWb As Workbook
Dim gstrOrigWbFile As String

Dim moduleWb As Workbook
Dim moduleWs As Worksheet

Dim EA As Excel.Application
Set EA = New Excel.Application

' Check ORIG workbook open and set to origWb/origWs
fileBoo = IsWorkBookOpen(gstrOrigWbFile)
If fileBoo = True Then
    Set origWb = Workbooks(gstrOrigWbName)     'File is open
Else
    Set origWb = Workbooks.Open(gstrOrigWbFile) 'File is Closed
End If
Dim origWs As Worksheet
Set origWs = origWb.Sheets(gstrOrigWsName)

' Now set up copied courseWb/courseWs
origWs.Copy
Set moduleWb = ActiveWorkbook
Set moduleWs = moduleWb.ActiveSheet
moduleWs.Name = "Module Reports"

finalRow = Range("A" & Rows.count).End(xlUp).Row
modulesRangeText = "L3:S" & finalRow & ",AA3:AA" & finalRow
Set modulesRange = moduleWs.Range(modulesRangeText)
If gblnDebugging Then
    'Debug.Print modulesRangeText
End If

starterRow = 3                      ' starts at row 3, rows 1+2 are headers
nameFound = ""                      ' set to the name of the "found" course
moduleCount = 1                     ' increases when courses found in sorted list
currentRow = starterRow - 1
NewSheet = True
Dim sheetNames As Object
Set sheetNames = CreateObject("Scripting.Dictionary")

For Each Cell In modulesRange
    thisRow = Cell.Row
    moduleFound = Cell.Text
    sanitisedModuleTitle = Replace(Replace(Replace(Replace(Left(moduleFound, 30), ":", " "), "?", ""), "(", ""), ")", "")
    moduleTitle = Replace(Replace(Replace(Replace(Left(moduleFound, 30), ":", " "), "?", ""), "(", ""), ")", "")
    
    ''''' TESTING - CHANGE TO NAMING/LOOKUP BY MODULE CODE!!!
    moduleCodeFound = Left(moduleFound, InStr(1, moduleFound, "-") - 1)
    Debug.Print "Module code found = " & moduleCodeFound
    
    ' Code thanks to Rory at http://stackoverflow.com/questions/6688131/test-or-check-if-sheet-exists
    ' http://stackoverflow.com/a/28473714
    ' http://stackoverflow.com/questions/34995962/check-if-sheet-exists?lq=1
    ' https://support.microsoft.com/en-us/kb/211601
        
    'wsExists = Evaluate("ISREF('" & sanitisedModuleTitle & "'!A1)")
    If Not sanitisedModuleTitle = "" Then
        NewSheet = True
        
        For Each Sheet In moduleWb.Sheets
            If Sheet.Name = sanitisedModuleTitle Then
                 NewSheet = False
            End If
        Next
        If NewSheet Then
            moduleWb.Sheets.Add Before:=Worksheets(Worksheets.count)
            moduleWb.Sheets(Worksheets.count - 1).Name = sanitisedModuleTitle
            If gblnDebugging Then
                Debug.Print "Creating new module sheet for " & sanitisedModuleTitle
            End If
            Set newModuleWs = moduleWb.Sheets(sanitisedModuleTitle)
            lastRow = gintSheetDataRows + 1
            
            ' Find threshold/cohort size
            Set rngFindValue = origWb.Sheets("MODULES").Range(gstrModuleLookupRng).Find(What:=moduleCodeFound, After:=origWb.Sheets("MODULES").Range("B8"), LookIn:=xlValues)
            rngFindRow = rngFindValue.Row
            cohortSize = origWb.Sheets("MODULES").Cells(rngFindRow, gintModuleLookupReturn).Value
            ''''' IF BELOW THRESHOLD -> NOTE THIS ON SHEET TITLE?
            responseThreshold = Application.Max(4, Round(cohortSize / 2, 0))
            
            newModuleWs.Range("A1") = moduleTitle          ' NOTE: not sanitised
            newModuleWs.Range("B1") = cohortSize
            newModuleWs.Range("C1") = responseThreshold
            
        Else
            Set newModuleWs = moduleWb.Sheets(sanitisedModuleTitle)
            lastRow = newModuleWs.Cells(Rows.count, "A").End(xlUp).Row + 1
        End If
        moduleWs.Range("A" & thisRow & ":CD" & thisRow).Copy Destination:=newModuleWs.Range("A" & lastRow)
        If gblnDebugging Then
            Debug.Print "Added student responses for " & sanitisedModuleTitle
        End If
    End If
Next
If gblnDebugging Then
    Debug.Print "----"
    Debug.Print "---- END Split MODULE Sheets"
    Debug.Print "----"
End If

End Sub




Sub generateModuleReports()

Dim colNum As Long

With wrdDoc

    For i = 1 To responseCount
        
        .Range(.Characters.count - 1).Style = .Styles(wdStyleHeading2)
        .Content.InsertAfter "Module-Level Responses for Respondent ID: " & Range("A" & i)
        .Content.InsertParagraphAfter
                
        ' Getting module-specific questions
        For colNum = 12 To 19 Step 1
            
            titleCol = Col_Letter(colNum)
            mTitle = Range(titleCol & i)
            If mTitle = "" Then
                GoTo skip
            End If
            
            satCol = Col_Letter(colNum + 18)
            mSat = Range(satCol & i)
            If mSat = "" Then
                mSat = "<<BLANK>>"
            End If
            
            bestCol = Col_Letter(colNum + 27)
            mBest = Range(bestCol & i)
            If mBest = "" Then
                mBest = "<<BLANK>>"
            End If
            
            worstCol = Col_Letter(colNum + 28)
            mWorst = Range(worstCol & i)
            If mWorst = "" Then
                mWorst = "<<BLANK>>"
            End If
            
            .Range(.Characters.count - 1).Style = .Styles(wdStyleHeading3)
            .Content.InsertAfter "Module: " & mTitle
            .Content.InsertParagraphAfter
            .Content.InsertAfter "Overall Satisfaction: " & mSat
            .Content.InsertParagraphAfter
            .Content.InsertAfter "Best Thing (free text): " & mBest
            .Content.InsertParagraphAfter
            .Content.InsertAfter "Worst Thing (free text): " & mWorst
            .Content.InsertParagraphAfter
            .Content.InsertParagraphAfter
            
        Next
        
        ' MANUALLY ADD MODULE 9 AS OUT OF SEQUENCE!
        titleCol = Col_Letter(27)
        mTitle = Range(titleCol & i)
        If mTitle = "" Then
            GoTo skip
        End If
        
        satCol = Col_Letter(38)
        mSat = Range(satCol & i)
        If mSat = "" Then
            mSat = "<<BLANK>>"
        End If
        
        bestCol = Col_Letter(55)
        mBest = Range(bestCol & i)
        If mBest = "" Then
            mBest = "<<BLANK>>"
        End If
        
        worstCol = Col_Letter(56)
        mWorst = Range(worstCol & i)
        If mWorst = "" Then
            mWorst = "<<BLANK>>"
        End If
        .Range(.Characters.count - 1).Style = .Styles(wdStyleHeading3)
        .Content.InsertAfter "Module: " & mTitle
        .Content.InsertParagraphAfter
        .Content.InsertAfter "Overall Satisfaction: " & mSat
        .Content.InsertParagraphAfter
        .Content.InsertAfter "Free Text - Best Aspects: " & mBest
        .Content.InsertParagraphAfter
        .Content.InsertAfter "Free Text - Worst Aspects: " & mWorst
        .Content.InsertParagraphAfter
        .Content.InsertParagraphAfter
            
        '' OLD MODULE DETAILS CUT FROM HERE!
skip:
    
        .Range(.Characters.count - 1).InlineShapes.AddHorizontalLineStandard
        If i <> responseCount Then
            .Range(.Characters.count - 1).InsertBreak
        End If
    
    Next i
    'If dir("C:\Foldername\MyNewWordDoc.doc") <> "" Then
    '    Kill "C:\Foldername\MyNewWordDoc.doc"
    'End If
    '.SaveAs ("C:\Foldername\MyNewWordDoc.doc")
    '.Close ' close the document
End With
    
'wrdApp.Quit ' close the Word application
Set wrdDoc = Nothing
Set wrdApp = Nothing
'MsgBox ("Course Report Generated")

End Sub

Sub manipulateWordTables()

ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=4, NumColumns:= _
        4, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
    With Selection.Tables(1)
        If .Style <> "Table Grid" Then
            .Style = "Table Grid"
        End If
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        .ApplyStyleRowBands = True
        .ApplyStyleColumnBands = False
    End With
    Selection.TypeText Text:="fdsa"
    Selection.MoveRight Unit:=wdCell
    Selection.TypeText Text:="sdef"
    Selection.MoveRight Unit:=wdCell
    Selection.TypeText Text:="ere"
    Selection.MoveRight Unit:=wdCell
    Selection.TypeText Text:="qwe"
    Selection.MoveRight Unit:=wdCell
    Selection.PasteAsNestedTable

End Sub

Sub oldModuleDetails()
If Range("L" & i) = "" Then
            mTitle = "<<DATA NOT FOUND>>"
        Else
            mTitle = Range("L" & i)
        End If
        .Range(.Characters.count - 1).Style = .Styles(wdStyleHeading3)
        .Content.InsertAfter "Module: " & mTitle
        .Content.InsertParagraphAfter
        
        If Range("AD" & i) = "" Then
            mSat = "<<DATA NOT FOUND>>"
        Else
            mSat = Range("AD" & i)
        End If
        .Content.InsertAfter "Overall Satisfaction: " & mSat
        .Content.InsertParagraphAfter
        
        If Range("AM" & i) = "" Then
            mBest = "<<DATA NOT FOUND>>"
        Else
            mBest = Range("AM" & i)
        End If
        .Content.InsertAfter "Free Text - Best Aspects: " & mBest
        .Content.InsertParagraphAfter
        
        If Range("AN" & i) = "" Then
            mWorst = "<<DATA NOT FOUND>>"
        Else
            mWorst = Range("AN" & i)
        End If
        .Content.InsertAfter "Free Text - Worst Aspects: " & mWorst
        .Content.InsertParagraphAfter
End Sub


