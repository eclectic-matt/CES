Sub generateOneSplitYearCourseReport()
    Dim reportingWorkbook As Workbook
    Dim courseTitleToProcess As String
    Set reportingWorkbook = ActiveWorkbook
    courseTitleToProcess = ActiveSheet.Name
    'courseTitleToProcess = ActiveWorkbook.Sheets(ActiveSheet.Name)
    Call generateSplitYearCourseReport(reportingWorkbook, courseTitleToProcess)
End Sub


Sub generateAllSplitYearCourseReports()
    
    Call startTimer
    
    current = 0
    total = ActiveWorkbook.Sheets.count - 2     'Note: "Course Reports" and "Summary Data" sheets ignored
    
    Call ShowProgressForm

    Dim EA As Excel.Application
    Set EA = New Excel.Application
    
    Dim reportingWorkbook As Workbook
    Set reportingWorkbook = ActiveWorkbook
    Dim courseTitleToProcess As String
    
    For Each Sheet In ActiveWorkbook.Sheets
        If (Not Sheet.Name = "Course Reports") And (Not Sheet.Name = "Summary Data") Then
            courseTitleToProcess = Sheet.Name
            If gblnDebugging Then
                Debug.Print "---------------------------"
                Debug.Print "Generating - " & Sheet.Name
                Debug.Print "---------------------------"
            End If
            current = current + 1
            Call UpdateProgressBar(current, total)
            'Worksheets(Sheet.Name).Activate
            Call generateSplitYearCourseReport(reportingWorkbook, courseTitleToProcess)
        End If
    Next

    Call Unload(ProgressForm)
    Call endTimer
    Set EA = Nothing

End Sub

Sub generateSplitYearCourseReport(reportWb As Workbook, courseCodeToProcess As String)

    If gblnDebugging Then
        Debug.Print "----"
        Debug.Print "---- START Generate Split Year Course Report"
        Debug.Print "----"
        Debug.Print "1) Setting Up"
    End If

    ' repWb = the Workbook/Sheet to be reported (Split Course Sheets)
    Dim repWb As Workbook
    Dim repWs As Worksheet
    Dim repWsName As String
    Set repWb = reportWb
    repWsName = courseCodeToProcess             'Debug.Print repWsName
    Set repWs = repWb.Worksheets(repWsName)     'Set repWs = repWb.ActiveSheet

    ' refWb = the REFERENCE Workbook/Sheet containing cohort sizes and other reporting data
    Dim refWb As Workbook
    Dim refWs As Worksheet
    
    ' sumWs = the SUMMARY Worksheet (part of repWb) to populate with summary data
    Dim sumWs As Worksheet
    Set sumWs = repWb.Worksheets("Summary Data")

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
        Set refWb = Workbooks(gstrRefWbName)
    Else
        Set refWb = Workbooks.Open(gstrRefWbFile)
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

' OKAY.... here is where updates are needed - cohortSize and responseThreshold no longer saved in sheet
' UPDATES HERE!!!!

'cohortSize = repWs.Range("B1").Text
responseRate = Round((responseCount / cohortSize) * 100, 2) & "%"
responseThreshold = repWs.Range("C1").Text
repWs.Range("A2:CE" & responseCount + 1).Sort key1:=repWs.Range("$CE:$CE"), Order2:=xlAscending ', Orientation:=xlTopToBottom, SortMethod:=xlPinYin, Header:=xlYes, MatchCase:=False,

If gblnDebugging Then
    Debug.Print "3) Getting course statistical data"
End If

studyYearCol = "$CE" & 2 & ":$CE" & (responseCount + 2)
studyYears = repWs.Range(studyYearCol)

yearFound = -100
Dim StudyYearsToProcess(0 To 4) As Integer      ' an array from 0 - 4 (the main UG study years) whose values are the # of respondents for that year
yearCounter = 1                                 ' Number of students in each year
thisRow = 1
'PWDcount = 0
For Each Cell In studyYears
    thisRow = thisRow + 1
    If Cell = yearFound Then
        ' Same year
        yearCounter = yearCounter + 1
    Else
        If yearFound = -100 Then
            ' Don't process, just start new
            yearFound = Cell
        Else
            ' Different year - Process OLD
            repWs.Rows(thisRow & ":" & (thisRow + (gintStatRows - 1))).EntireRow.Insert
            thisRow = thisRow + gintStatRows
            StudyYearsToProcess(yearFound) = yearCounter
            yearFound = Cell
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
        Debug.Print "Resp Count for Year " & a & " = " & responseCount
        ' cohortSize = lookup in reference sheet, the column based on "a" (StudyYear)
        Set cohortRowFound = refWs.Range(gstrCourseLookupRng).Find(What:=repWs.Name, After:=refWs.Range(Left(gstrCourseLookupRng, InStr(1, gstrCourseLookupRng, ":") - 1)), LookIn:=xlValues)
        Debug.Print(cohortRowFound)
        cohortSize = refWs.Cells(cohortRowFound,4+a).Value2
        Debug.Print(cohortSize)
        responseRate = Round((responseCount / cohortSize) * 100, 2) & "%"
        
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

With sumWs
    endRow = .Range("A" & Rows.count).End(xlUp).Row + 1
    If gblnDebugging Then
        Debug.Print "Summary End Row = " & endRow
    End If
    .Cells(endRow, 1) = courseCode
    .Cells(endRow, 2) = courseTitle
    .Cells(endRow, 3) = cohortSize
    .Cells(endRow, 4) = responseRate
    .Cells(endRow, 5) = Average
    .Cells(endRow, 6) = Median
    .Cells(endRow, 7) = ValidResponses
    .Cells(endRow, 8) = StudyYear
End With
    
' PART 2 - CREATING A COURSE REPORT
If gblnDebugging Then
    Debug.Print "4) Create Word Doc"
End If
Dim wrdApp As Word.Application
Dim wrdDoc As Word.document
Set wrdApp = CreateObject("Word.Application")
wrdApp.Visible = False

If cohortSize < gintPublicationThreshold Then
    oName = gstrReportsFilePath & "COURSE REPORTS\" & "DO NOT PUBLISH - " & sanitisedCourseTitle & " YEAR " & StudyYear & " [" & Format(Date, "dd-mm-yy") & " " & Format(Time, "hh.mm.ss") & "].pdf"
    wrdDoc.SaveAs2 FileName:=oName, FileFormat:=wdFormatPDF
    wrdDoc.Close (False)
    sumWs.Cells(endRow, 8) = "Not Published"
    GoTo doNotPublish
End If

''''' WORKING HERE!!!
startRow = 2    'Then startRow = responseCount + startRow + gintStatRows
For StudyYear = 0 To 4
    Debug.Print "Checking Year " & StudyYear & " (" & StudyYearsToProcess(StudyYear) & ")"
    If Not StudyYearsToProcess(StudyYear) = 0 Then
        'Debug.Print "Processing Year " & studyYear
        Set rngFindValue = refWs.Range(gstrCourseLookupRng).Find(What:=repWs.Name, After:=refWs.Range(Left(gstrCourseLookupRng, InStr(1, gstrCourseLookupRng, ":") - 1)), LookIn:=xlValues)
        rngFindRow = rngFindValue.Row
        cohortSize = refWs.Cells(rngFindRow, 4 + StudyYear).Value2
        responseThreshold = getResponseThreshold(cohortSize)
        responseCount = StudyYearsToProcess(StudyYear)
        responseRate = Round((responseCount / cohortSize) * 100, 2) & "%"
        'Debug.Print "RESPONSES = " & responseCount & " so RESPRATE = " & responseRate
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

doNotPublish:
    wrdApp.Quit (False)
    Set wrdApp = Nothing
    'Set EA = Nothing
    Debug.Print "-------------------------------"
    Debug.Print "COMPLETE - " & courseTitle
    Debug.Print "-------------------------------"

End Sub


