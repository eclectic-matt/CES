Sub createAllModuleSummaries()

Dim repWb As Workbook
Dim repWs As Worksheet
Dim schoolRng As Range
Dim deptRng As Range
Dim copyRng As Range
Dim endRow, startRow, finalRow As Integer
Dim School As String
Dim counter As Integer

Set repWb = ActiveWorkbook
Set repWs = repWb.Worksheets("Summary Data")
endRow = repWs.Range("A" & repWs.Rows.count).End(xlUp).Row
Debug.Print endRow
Set schoolRng = repWs.Range("$K1:$K" & endRow)
'NO! Need to copy specific rows! Set copyRng = repWs.Range("$A:$J")
'N.B. For course summaries, key3:="B:B"
repWs.Range("$A:$L").Sort key1:=repWs.Range("$K:$K"), key2:=repWs.Range("$J:$J"), key3:=repWs.Range("$A:$A"), Header:=xlYes, MatchCase:=False, Orientation:=xlTopToBottom, SortMethod:=xlPinYin

School = "Err"
startRow = 2
thisRow = startRow
For Each Cell In schoolRng
    thisRow = thisRow + 1
    If Not Cell = "School" Then
        If Not Cell = schoolFound And schoolFound <> "" Then
            finalRow = thisRow - 3
            startRow = Application.WorksheetFunction.Max(startRow - 2, 2)
            Debug.Print schoolFound & " - copy from Row " & startRow & " to Row " & finalRow & " (" & (finalRow - startRow + 1) & " in total)"
            Call createModuleSummaryReport(schoolFound, startRow, finalRow)
            startRow = thisRow
        End If
        schoolFound = Cell
    End If
Next
startRow = startRow - 2
finalRow = thisRow - 2
Debug.Print schoolFound & " - copy from Row " & startRow & " to Row " & finalRow & " (" & (finalRow - startRow + 1) & " in total)"
Call createModuleSummaryReport(schoolFound, startRow, finalRow)
Debug.Print ("All School Module Summary Reports done!")
End Sub

Sub createOneModuleSummaryReport()
Dim startRow, finalRow As Integer
startRow = 668
finalRow = 732
schoolFound = "SCLS"
Call createModuleSummaryReport(schoolFound, startRow, finalRow)
End Sub

Sub createModuleSummaryReport(ByVal School As String, ByRef firstRow, lastRow As Integer)

Dim deptRng As Range
Dim BooFirst As Boolean
Set deptRng = ActiveSheet.Range("$J" & firstRow & ":$J" & lastRow)

Dim wrdApp As Word.Application
Dim wrdDoc As Word.document
Set wrdApp = CreateObject("Word.Application")
'wrdApp.Visible = False
Set wrdDoc = wrdApp.Documents.Add


With wrdDoc
    ' SET DOCUMENT STYLES
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
    With .Styles(wdStyleHeading4)
        .ParagraphFormat.Alignment = wdAlignParagraphRight
        With .Font
            .Name = "Arial"
            .Size = 10
            .Bold = True
            .Italic = True
            .Color = wdColorGray80
        End With
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
    ' HEADERS AND FOOTERS AND PAGE SETUP
    .PageSetup.Orientation = wdOrientLandscape
    With .Sections(1).Headers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary)
        '.Range.Style = .Styles(wdStyleHeading1)
        .Range.Text = "COURSE EVALUATION SURVEY REPORT 2015/16" & vbCr & _
            "SCHOOL-LEVEL REPORT FOR " & School & vbCr & _
            "MODULE SUMMARY REPORT" & vbCr
        .Range.ParagraphFormat.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
    End With
    With .Sections(1).Footers(wdHeaderFooterPrimary)
        .Range.Text = "CES Module Summary Report for " & School & " (generated: " & Format(Date, "dd-mm-yy") & " " & Format(Time, "hh.mm.ss") & ")"
        .PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberRight, FirstPage:=True
    End With
End With

startRow = 0
thisRow = startRow
For Each Cell In deptRng
    thisRow = thisRow + 1
    If Not Cell = "Department" Then
        If Not Cell = deptFound And deptFound <> "" Then
            ''' PROCESS DEPT DATA
            finalRow = thisRow - 2
            If Not startRow = 0 Then
                startRow = startRow - 1
            End If
            beginAt = firstRow + startRow
            endAt = firstRow + finalRow
            Debug.Print "   --> " & deptFound & " - copy from "; beginAt & " to " & endAt
            wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading2)
            With wrdDoc
                ' Titles and course details
                .Range(wrdDoc.Characters.count - 1).Style = .Styles(wdStyleHeading1)
                .Range(.Characters.count - 1).Select
                .Content.InsertAfter "DEPARTMENT: " & deptFound
                .Content.InsertParagraphAfter
            End With
            
            intNoOfRows = endAt - beginAt + 2
            intNoOfColumns = 9
            Set myRange = wrdDoc.Range(Start:=wrdDoc.Characters.count - 1, End:=wrdDoc.Characters.count)
            Set overallTable = wrdDoc.Tables.Add(Range:=myRange, NumRows:=intNoOfRows, NumColumns:= _
                intNoOfColumns, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
                wdAutoFitFixed)
            overallTable.Style = "Grid Table 1 Light"
            overallTable.ApplyStyleHeadingRows = True
            
            ' Table Headings
            headerRow = Array("Module Code", "Module Title", "Cohort Size", "Average Satisfaction", "Median Satisfaction", "Valid Responses", "Valid Response Rate (%)", "FHEQ Level", "Published Flag")
            For Index = 0 To (intNoOfColumns - 1)
                Set tableCell = overallTable.Cell(1, Index + 1)
                tableCell.Range.InsertAfter Text:=headerRow(Index)
            Next
                        
            ' Response Data
            For RowIndex = 0 To (intNoOfRows - 2)
                For ColIndex = 1 To (intNoOfColumns)
                    Set tableCell = overallTable.Cell(2 + RowIndex, ColIndex)
                    If ColIndex = 2 Then
                        overallTable.Columns(ColIndex).SetWidth ColumnWidth:=(gsngTotalPageWidthPoints / 4), RulerStyle:=wdAdjustNone
                    Else
                        overallTable.Columns(ColIndex).SetWidth ColumnWidth:=(3 * (gsngTotalPageWidthPoints / 32)), RulerStyle:=wdAdjustNone
                    End If
                    'overallTable.Columns(Index + 2).SetWidth ColumnWidth:=(gsngTotalPageWidthPoints / 2) / 7, RulerStyle:=wdAdjustNone
                    cellData = Excel.Application.Cells(beginAt + RowIndex, ColIndex).Text
                    'Debug.Print "Processing --> " & cellData
                    tableCell.Range.InsertAfter Text:=cellData
                Next
            Next
            
            wrdDoc.Range(wrdDoc.Characters.count - 1).InsertBreak (wdPageBreak)
            ''' END PROCESS DEPT DATA
            startRow = thisRow
        End If
        deptFound = Cell
    End If
Next
''' PROCESS FINAL DEPT DATA
finalRow = thisRow - 1
If Not startRow = 0 Then
    startRow = startRow - 1
End If
beginAt = firstRow + startRow
endAt = firstRow + finalRow
Debug.Print "   --> " & deptFound & " - copy from "; beginAt & " to " & endAt
wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading2)
With wrdDoc
    ' Titles and course details
    .Range(wrdDoc.Characters.count - 1).Style = .Styles(wdStyleHeading1)
    .Range(.Characters.count - 1).Select
    .Content.InsertAfter "DEPARTMENT: " & deptFound
    .Content.InsertParagraphAfter
End With

intNoOfRows = endAt - beginAt + 2
intNoOfColumns = 9
Set myRange = wrdDoc.Range(Start:=wrdDoc.Characters.count - 1, End:=wrdDoc.Characters.count)
Set overallTable = wrdDoc.Tables.Add(Range:=myRange, NumRows:=intNoOfRows, NumColumns:= _
    intNoOfColumns, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
    wdAutoFitFixed)
overallTable.Style = "Grid Table 1 Light"
overallTable.ApplyStyleHeadingRows = True

' Table Headings
headerRow = Array("Module Code", "Module Title", "Cohort Size", "Average Satisfaction", "Median Satisfaction", "Valid Responses", "Valid Response Rate (%)", "FHEQ Level", "Published Flag")
For Index = 0 To (intNoOfColumns - 1)
    Set tableCell = overallTable.Cell(1, Index + 1)
    tableCell.Range.InsertAfter Text:=headerRow(Index)
Next
            
' Response Data
For RowIndex = 0 To (intNoOfRows - 2)
    For ColIndex = 1 To (intNoOfColumns)
        Set tableCell = overallTable.Cell(2 + RowIndex, ColIndex)
        If ColIndex = 2 Then
            overallTable.Columns(ColIndex).SetWidth ColumnWidth:=(gsngTotalPageWidthPoints / 4), RulerStyle:=wdAdjustNone
        Else
            overallTable.Columns(ColIndex).SetWidth ColumnWidth:=(3 * (gsngTotalPageWidthPoints / 32)), RulerStyle:=wdAdjustNone
        End If
        'overallTable.Columns(Index + 2).SetWidth ColumnWidth:=(gsngTotalPageWidthPoints / 2) / 7, RulerStyle:=wdAdjustNone
        cellData = Excel.Application.Cells(beginAt + RowIndex, ColIndex).Text
        'Debug.Print "Processing --> " & cellData
        tableCell.Range.InsertAfter Text:=cellData
    Next
Next
''' END PROCESS FINAL DEPT DATA
oName = gstrReportsFilePath & "SCHOOL REPORTS/SUMMARY REPORTS/Module Summary - " & School & " [" & Format(Date, "dd-mm-yy") & " " & Format(Time, "hh.mm.ss") & "].pdf"
wrdDoc.SaveAs2 FileName:=oName, FileFormat:=wdFormatPDF
wrdDoc.Close (False)
If gblnDebugging Then
    Debug.Print "Saved Report as PDF - " & oName
End If
wrdApp.Quit (False)
Set wrdApp = Nothing
'Application.CutCopyMode = False

End Sub

Sub createAllCourseSummaries()

Dim repWb As Workbook
Dim repWs As Worksheet
Dim schoolRng As Range
Dim deptRng As Range
Dim copyRng As Range
Dim endRow, startRow, finalRow As Integer
Dim School As String
Dim counter As Integer

Set repWb = ActiveWorkbook
Set repWs = repWb.Worksheets("Summary Data")
endRow = repWs.Range("A" & repWs.Rows.count).End(xlUp).Row
Debug.Print endRow
Set schoolRng = repWs.Range("$K1:$K" & endRow)
'NO! Need to copy specific rows! Set copyRng = repWs.Range("$A:$J")
'N.B. For course summaries, key3:="B:B"
repWs.Range("$A:$K").Sort key1:=repWs.Range("$K:$K"), key2:=repWs.Range("$J:$J"), key3:=repWs.Range("$A:$A"), Header:=xlYes, MatchCase:=False, Orientation:=xlTopToBottom, SortMethod:=xlPinYin

School = "Err"
startRow = 2
thisRow = startRow
For Each Cell In schoolRng
    thisRow = thisRow + 1
    If Not Cell = "School" Then
        If Not Cell = schoolFound And schoolFound <> "" Then
            finalRow = thisRow - 3
            startRow = Application.WorksheetFunction.Max(startRow - 2, 2)
            Debug.Print schoolFound & " - copy from Row " & startRow & " to Row " & finalRow & " (" & (finalRow - startRow + 1) & " in total)"
            Call createCourseSummaryReport(schoolFound, startRow, finalRow)
            startRow = thisRow
        End If
        schoolFound = Cell
    End If
Next
startRow = startRow - 2
finalRow = thisRow - 2
Debug.Print schoolFound & " - copy from Row " & startRow & " to Row " & finalRow & " (" & (finalRow - startRow + 1) & " in total)"
Call createCourseSummaryReport(schoolFound, startRow, finalRow)
Debug.Print ("All School Course Summary Reports done!")
End Sub

Sub createCourseSummaryReport(ByVal School As String, ByRef firstRow, lastRow As Integer)

Dim deptRng As Range
Dim BooFirst As Boolean
Set deptRng = ActiveSheet.Range("$J" & firstRow & ":$J" & lastRow)

Dim wrdApp As Word.Application
Dim wrdDoc As Word.document
Set wrdApp = CreateObject("Word.Application")
'wrdApp.Visible = False
Set wrdDoc = wrdApp.Documents.Add


With wrdDoc
    ' SET DOCUMENT STYLES
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
    With .Styles(wdStyleHeading4)
        .ParagraphFormat.Alignment = wdAlignParagraphRight
        With .Font
            .Name = "Arial"
            .Size = 10
            .Bold = True
            .Italic = True
            .Color = wdColorGray80
        End With
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
    ' HEADERS AND FOOTERS AND PAGE SETUP
    .PageSetup.Orientation = wdOrientLandscape
    With .Sections(1).Headers(Word.WdHeaderFooterIndex.wdHeaderFooterPrimary)
        '.Range.Style = .Styles(wdStyleHeading1)
        .Range.Text = "COURSE EVALUATION SURVEY REPORT 2015/16" & vbCr & _
            "SCHOOL-LEVEL REPORT FOR " & School & vbCr & _
            "COURSE SUMMARY REPORT" & vbCr
        .Range.ParagraphFormat.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
    End With
    With .Sections(1).Footers(wdHeaderFooterPrimary)
        .Range.Text = "CES Course Summary Report for " & School & " (generated: " & Format(Date, "dd-mm-yy") & " " & Format(Time, "hh.mm.ss") & ")"
        .PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberRight, FirstPage:=True
    End With
End With

startRow = 0
thisRow = startRow
For Each Cell In deptRng
    thisRow = thisRow + 1
    If Not Cell = "Department" Then
        If Not Cell = deptFound And deptFound <> "" Then
            ''' PROCESS DEPT DATA
            finalRow = thisRow - 2
            If Not startRow = 0 Then
                startRow = startRow - 1
            End If
            beginAt = firstRow + startRow
            endAt = firstRow + finalRow
            Debug.Print "   --> " & deptFound & " - copy from "; beginAt & " to " & endAt
            wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading2)
            With wrdDoc
                ' Titles and course details
                .Range(wrdDoc.Characters.count - 1).Style = .Styles(wdStyleHeading1)
                .Range(.Characters.count - 1).Select
                .Content.InsertAfter "DEPARTMENT: " & deptFound
                .Content.InsertParagraphAfter
            End With
            
            intNoOfRows = endAt - beginAt + 2
            intNoOfColumns = 9
            Set myRange = wrdDoc.Range(Start:=wrdDoc.Characters.count - 1, End:=wrdDoc.Characters.count)
            Set overallTable = wrdDoc.Tables.Add(Range:=myRange, NumRows:=intNoOfRows, NumColumns:= _
                intNoOfColumns, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
                wdAutoFitFixed)
            overallTable.Style = "Grid Table 1 Light"
            overallTable.ApplyStyleHeadingRows = True
            
            ' Table Headings
            headerRow = Array("Course Code", "Course Title", "Study Year", "Cohort Size", "Average Satisfaction", "Median Satisfaction", "Valid Responses", "Valid Response Rate (%)", "Published Flag")
            For Index = 0 To (intNoOfColumns - 1)
                Set tableCell = overallTable.Cell(1, Index + 1)
                tableCell.Range.InsertAfter Text:=headerRow(Index)
            Next
                        
            ' Response Data
            For RowIndex = 0 To (intNoOfRows - 2)
                For ColIndex = 1 To (intNoOfColumns)
                    Set tableCell = overallTable.Cell(2 + RowIndex, ColIndex)
                    If ColIndex = 2 Then
                        overallTable.Columns(ColIndex).SetWidth ColumnWidth:=(gsngTotalPageWidthPoints / 4), RulerStyle:=wdAdjustNone
                    Else
                        overallTable.Columns(ColIndex).SetWidth ColumnWidth:=(3 * (gsngTotalPageWidthPoints / 32)), RulerStyle:=wdAdjustNone
                    End If
                    'overallTable.Columns(Index + 2).SetWidth ColumnWidth:=(gsngTotalPageWidthPoints / 2) / 7, RulerStyle:=wdAdjustNone
                    cellData = Excel.Application.Cells(beginAt + RowIndex, ColIndex).Text
                    'Debug.Print "Processing --> " & cellData
                    tableCell.Range.InsertAfter Text:=cellData
                Next
            Next
            
            wrdDoc.Range(wrdDoc.Characters.count - 1).InsertBreak (wdPageBreak)
            ''' END PROCESS DEPT DATA
            startRow = thisRow
        End If
        deptFound = Cell
    End If
Next
''' PROCESS FINAL DEPT DATA
finalRow = thisRow - 1
If Not startRow = 0 Then
    startRow = startRow - 1
End If
beginAt = firstRow + startRow
endAt = firstRow + finalRow
Debug.Print "   --> " & deptFound & " - copy from "; beginAt & " to " & endAt
wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading2)
With wrdDoc
    ' Titles and course details
    .Range(wrdDoc.Characters.count - 1).Style = .Styles(wdStyleHeading1)
    .Range(.Characters.count - 1).Select
    .Content.InsertAfter "DEPARTMENT: " & deptFound
    .Content.InsertParagraphAfter
End With

intNoOfRows = endAt - beginAt + 2
intNoOfColumns = 9
Set myRange = wrdDoc.Range(Start:=wrdDoc.Characters.count - 1, End:=wrdDoc.Characters.count)
Set overallTable = wrdDoc.Tables.Add(Range:=myRange, NumRows:=intNoOfRows, NumColumns:= _
    intNoOfColumns, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
    wdAutoFitFixed)
overallTable.Style = "Grid Table 1 Light"
overallTable.ApplyStyleHeadingRows = True

' Table Headings
headerRow = Array("Course Code", "Course Title", "Study Year", "Cohort Size", "Average Satisfaction", "Median Satisfaction", "Valid Responses", "Valid Response Rate (%)", "Published Flag")
For Index = 0 To (intNoOfColumns - 1)
    Set tableCell = overallTable.Cell(1, Index + 1)
    tableCell.Range.InsertAfter Text:=headerRow(Index)
Next
            
' Response Data
For RowIndex = 0 To (intNoOfRows - 2)
    For ColIndex = 1 To (intNoOfColumns)
        Set tableCell = overallTable.Cell(2 + RowIndex, ColIndex)
        If ColIndex = 2 Then
            overallTable.Columns(ColIndex).SetWidth ColumnWidth:=(gsngTotalPageWidthPoints / 4), RulerStyle:=wdAdjustNone
        Else
            overallTable.Columns(ColIndex).SetWidth ColumnWidth:=(3 * (gsngTotalPageWidthPoints / 32)), RulerStyle:=wdAdjustNone
        End If
        'overallTable.Columns(Index + 2).SetWidth ColumnWidth:=(gsngTotalPageWidthPoints / 2) / 7, RulerStyle:=wdAdjustNone
        cellData = Excel.Application.Cells(beginAt + RowIndex, ColIndex).Text
        'Debug.Print "Processing --> " & cellData
        tableCell.Range.InsertAfter Text:=cellData
    Next
Next
''' END PROCESS FINAL DEPT DATA
oName = gstrReportsFilePath & "SCHOOL REPORTS/SUMMARY REPORTS/Course Summary - " & School & " [" & Format(Date, "dd-mm-yy") & " " & Format(Time, "hh.mm.ss") & "].pdf"
wrdDoc.SaveAs2 FileName:=oName, FileFormat:=wdFormatPDF
wrdDoc.Close (False)
If gblnDebugging Then
    Debug.Print "Saved Report as PDF - " & oName
End If
wrdApp.Quit (False)
Set wrdApp = Nothing
'Application.CutCopyMode = False

End Sub

Sub createAllCourseChecksheets()
Application.DisplayAlerts = False
Dim reportingWorkbook As Workbook
Set reportingWorkbook = Workbooks("Split Course Report Sheets (08-07-16 11.48.22)")

current = 0
For Each Sheet In ActiveWorkbook.Sheets
    If (Not Sheet.Name = "Course Reports") And (Not Sheet.Name = "Summary Data") Then
        courseTitleToProcess = Sheet.Name
        If gblnDebugging Then
            Debug.Print "---------------------------"
            Debug.Print "Summarising - " & Sheet.Name
            Debug.Print "---------------------------"
        End If
        current = current + 1
        Call addToCourseChecksheet(reportingWorkbook, courseTitleToProcess)
    End If
Next
Application.DisplayAlerts = True
If gblnDebugging Then
    Debug.Print "---------------------------"
    Debug.Print "ENDED Summarising"
    Debug.Print "---------------------------"
End If
End Sub

Sub addToCourseChecksheet(ByRef repWb As Workbook, ByVal course As String)

Dim refWb As Workbook
Dim refWs As Worksheet
Dim repWs As Worksheet

Set repWs = repWb.Worksheets(course)

fileBoo = False                                 ' Testing this as issues with open files
fileBoo = IsWorkBookOpen(gstrRefWbFile)
If fileBoo = True Then
    Set refWb = Workbooks(gstrRefWbName)
    Debug.Print gstrRefWbFile & " -> Open file detected!"
Else
    Set refWb = Workbooks.Open(gstrRefWbFile)
    Debug.Print gstrRefWbFile & " -> Opening this file!"
End If
Set refWs = refWb.Sheets("COURSES")
    
School = Application.WorksheetFunction.VLookup(course, refWs.Range(gstrCourseLookupRng), 7, False)
Debug.Print course & " -> " & School
Dim schoolWb As Workbook
Dim DirFile As String

DirFile = gstrReportsFilePath & "/SCHOOL REPORTS/CHECK SHEETS/" & School & " - Course Checksheet.xlsx"
If Len(dir(DirFile)) = 0 Then
    Debug.Print DirFile & " does not exist"
    Set schoolWb = Workbooks.Add(xlWBATWorksheet)
    schoolWb.SaveAs gstrReportsFilePath & "/SCHOOL REPORTS/CHECK SHEETS/" & School & " - Course Checksheet.xlsx"
Else
    fileBoo = False                                 ' Testing this as issues with open files
    fileBoo = IsWorkBookOpen(DirFile)
    If fileBoo = True Then
        If schoolWb Is Nothing Then
            Set schoolWb = Workbooks(School & " - Course Checksheet.xlsx")
        End If
        Debug.Print DirFile & " -> Open file detected!"
    Else
        Set schoolWb = Workbooks.Open(DirFile)
        Debug.Print DirFile & " -> Opening this file!"
    End If
End If
Debug.Print "schoolWb has " & schoolWb.Sheets.count & " sheets"
repWb.Sheets(course).Copy After:=schoolWb.Sheets(schoolWb.Sheets.count)

Dim testWs As Worksheet
On Error Resume Next
Set testWs = schoolWb.Sheets("Sheet1")
On Error GoTo 0
If Not testWs Is Nothing Then
    testWs.Delete
End If
 
Set copiedWs = schoolWb.Sheets(schoolWb.Sheets.count)
endRow = copiedWs.Range("A" & Rows.count).End(xlUp).Row
Debug.Print course & " (endRow)=" & endRow
With copiedWs
    .Cells(1, 1).Value = "RespondentID"
    .Columns("B:BZ").Delete Shift:=xlToLeft
    .Columns("C:F").Delete Shift:=xlToLeft
    .Columns("A:B").AutoFit
    .Columns("B:B").ColumnWidth = 60
    .Range("A1:D1").Font.Bold = True
    With .Range("B2:D" & endRow)
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    .Cells(1, 2).Value = "Free-Text Comments"
End With

For thisRow = 1 To endRow
    If Len(copiedWs.Cells(thisRow, 1)) = 0 Then
        copiedWs.Rows(thisRow & ":" & (thisRow + gintStatRows - 1)).EntireRow.Delete Shift:=xlToUp
    End If
Next
copiedWs.Rows(endRow + 1 & ":" & endRow + gintStatRows).EntireRow.Delete
'' NOTE: Sheet(s) must be saved manually once summaries generated!
End Sub

Sub createAllModuleChecksheets()
Application.DisplayAlerts = False
Dim reportingWorkbook As Workbook
Set reportingWorkbook = Workbooks("School Checking Sheet - MODULES")

current = 0
For Each Sheet In ActiveWorkbook.Sheets
    If (Not Sheet.Name = "Module Reports") And (Not Sheet.Name = "Summary Data") Then
        courseTitleToProcess = Sheet.Name
        If gblnDebugging Then
            Debug.Print "---------------------------"
            Debug.Print "Summarising - " & Sheet.Name
            Debug.Print "---------------------------"
        End If
        current = current + 1
        Call addToModuleChecksheet(reportingWorkbook, courseTitleToProcess)
    End If
Next
Application.DisplayAlerts = True
If gblnDebugging Then
    Debug.Print "---------------------------"
    Debug.Print "ENDED Summarising"
    Debug.Print "---------------------------"
End If
End Sub

Sub addToModuleChecksheet(ByRef repWb As Workbook, ByVal course As String)

Dim refWb As Workbook
Dim refWs As Worksheet
Dim repWs As Worksheet

Set repWs = repWb.Worksheets(course)

fileBoo = False                                 ' Testing this as issues with open files
fileBoo = IsWorkBookOpen(gstrRefWbFile)
If fileBoo = True Then
    Set refWb = Workbooks(gstrRefWbName)
    Debug.Print gstrRefWbFile & " -> Open file detected!"
Else
    Set refWb = Workbooks.Open(gstrRefWbFile)
    Debug.Print gstrRefWbFile & " -> Opening this file!"
End If
Set refWs = refWb.Sheets("MODULES")
    
School = Application.WorksheetFunction.VLookup(course, refWs.Range(gstrModuleLookupRng), 5, False)
Debug.Print course & " -> " & School
Dim schoolWb As Workbook
Dim DirFile As String

DirFile = gstrReportsFilePath & "/SCHOOL REPORTS/CHECK SHEETS/" & School & " - Module Checksheet.xlsx"
If Len(dir(DirFile)) = 0 Then
    Debug.Print DirFile & " does not exist"
    Set schoolWb = Workbooks.Add(xlWBATWorksheet)
    schoolWb.SaveAs gstrReportsFilePath & "/SCHOOL REPORTS/CHECK SHEETS/" & School & " - Module Checksheet.xlsx"
Else
    fileBoo = False                                 ' Testing this as issues with open files
    fileBoo = IsWorkBookOpen(DirFile)
    If fileBoo = True Then
        If schoolWb Is Nothing Then
            Set schoolWb = Workbooks(School & " - Module Checksheet.xlsx")
        End If
        Debug.Print DirFile & " -> Open file detected!"
    Else
        Set schoolWb = Workbooks.Open(DirFile)
        Debug.Print DirFile & " -> Opening this file!"
    End If
End If
Debug.Print "schoolWb has " & schoolWb.Sheets.count & " sheets"
repWb.Sheets(course).Copy After:=schoolWb.Sheets(schoolWb.Sheets.count)

Dim testWs As Worksheet
On Error Resume Next
Set testWs = schoolWb.Sheets("Sheet1")
On Error GoTo 0
If Not testWs Is Nothing Then
    testWs.Delete
End If
 
'' NOTE: Sheet(s) must be saved manually once summaries generated!
End Sub
