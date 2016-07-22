Sub generateOneModuleReport()
    Dim reportingWorkbook As Workbook
    Dim moduleTitleToProcess As String
    Set reportingWorkbook = ActiveWorkbook
    moduleTitleToProcess = ActiveSheet.Name
    Call generateSingleModuleReport(reportingWorkbook, moduleTitleToProcess)
End Sub

Sub generateSpecificModuleReport()
    Dim reportingWorkbook As Workbook
    Dim moduleTitleToProcess As String
    Set reportingWorkbook = Workbooks("Split Module Report Sheets (08-07-16 12.09.28)")
    moduleTitleToProcess = "V1413"
    Call generateSingleModuleReport(reportingWorkbook, moduleTitleToProcess)
End Sub


Sub generateAllModuleReports()

Call startTimer

Dim reportingWorkbook As Workbook
Set reportingWorkbook = ActiveWorkbook
Dim moduleCodeToProcess As String

current = 0
total = ActiveWorkbook.Sheets.count - 1     'Note: Module Reports sheet ignored
Call ShowProgressForm

For Each Sheet In ActiveWorkbook.Sheets
    moduleCodeToProcess = Sheet.Name
    If (Not moduleCodeToProcess = "Module Reports") And (Not moduleCodeToProcess = "Summary Data") Then
        If gblnDebugging Then
            Debug.Print "---------------------------"
            Debug.Print "Generating - " & moduleCodeToProcess
            Debug.Print "---------------------------"
        End If
        current = current + 1
        Call UpdateProgressBar(current, total)
        'Sheet.Select
        Call generateSingleModuleReport(reportingWorkbook, moduleCodeToProcess)
    End If
Next

Call Unload(ProgressForm)
Call endTimer

End Sub

Sub generateSingleModuleReport(reportWb As Workbook, moduleCode As String)

Dim repWb As Workbook
Dim repWs As Worksheet

If gblnDebugging Then
    Debug.Print "----"
    Debug.Print "---- START Generate Module Report"
    Debug.Print "----"
    Debug.Print "1) Setting Up"
End If

Dim wrdApp As Word.Application
Dim wrdDoc As Word.document

Dim EA As Excel.Application
Set EA = New Excel.Application
Dim origWb As Workbook

Set repWb = reportWb
Set repWs = repWb.Worksheets(moduleCode)

Dim sumWs As Worksheet
Set sumWs = repWb.Sheets("Summary Data")

'moduleCode = repWs.Name
moduleTitle = repWs.Range("A1").Text
cohortSize = repWs.Range("B1").Text
fullModule = moduleCode & " - " & moduleTitle
responseThreshold = repWs.Range("C1").Text
responseCount = repWs.Range("A" & Rows.count).End(xlUp).Row - 1
responseRate = Round((responseCount / cohortSize) * 100, 2) & "%"
If responseCount < responseThreshold Then
    disclaimer = gstrThresholdDisclaimer
    disclaimer = Replace(disclaimer, "%RESP", responseCount)
    disclaimer = Replace(disclaimer, "$THRE", responseThreshold)
Else
    disclaimer = ""
End If


If gblnDebugging Then
    Debug.Print "       Number of responses: " & responseCount
    Debug.Print "       Reporting for " & moduleCode & " with " & responseCount & " responses"
End If

If gblnDebugging Then
    Debug.Print "2) Pre-process module data (" & moduleTitle & ")"
    ' Match up module data (in various columns) to the moduleCode listed
End If

moduleEndRow = responseCount + 1
With repWs
    Set module18DataRng = .Range(.Cells(gintModuleStartRow, gintModuleDataStartCol), .Cells(moduleEndRow, gintModuleDataEndCol))
    Set module9DataRng = .Range(.Cells(gintModuleStartRow, 27), .Cells(moduleEndRow, 27))
End With

' PROCESS - each MODULE sheet contains ALL responses from students
' Find "this" module (from Sheet.Name) and filter responses

reportingAtRow = responseCount + 2
' Reporting Columns =
    ' A = Stars
    ' B = Best Free
    ' C = Worst Free

' Modules 1 - 8
For Each Cell In module18DataRng
    If Not Cell = "" Then
        modCodeFound = Trim(Left(Cell.Text, InStr(1, Cell.Text, "-") - 1))
        If modCodeFound = moduleCode Then
            modRow = Cell.Row
            modNum = Cell.Column - 11
            modStars = repWs.Cells(modRow, modNum + 29).Value2
            If modStars = Empty Then
                modStars = "N/A"
            End If
            modCommBest = repWs.Cells(modRow, (2 * modNum) + 37).Value2
            Set modCommBestRng = repWs.Range(repWs.Cells(modRow, (2 * modNum) + 37), repWs.Cells(modRow, (2 * modNum) + 37))
            If modCommBestRng.HasFormula Then
                Debug.Print "----- COMMENT SANITISE"
                Debug.Print "WAS: " & modCommBestRng.Formula
                modCommBest = Replace(modCommBestRng.Formula, "=", "-")
                Debug.Print "NOW: " & modCommBest
            End If
            modCommWorst = repWs.Cells(modRow, (2 * modNum) + 38).Value2
            Set modCommWRng = repWs.Range(repWs.Cells(modRow, (2 * modNum) + 38), repWs.Cells(modRow, (2 * modNum) + 38))
            If modCommWRng.HasFormula Then
                Debug.Print "----- COMMENT SANITISE"
                Debug.Print "WAS: " & modCommWRng.Formula
                modCommWorst = Replace(modCommWRng.Formula, "=", "-")
                Debug.Print "NOW: " & modCommWorst
            End If
            
            ' Then report this (same sheet, printed lower down)
            repWs.Cells(reportingAtRow + (modRow - 1), 1) = modStars
            repWs.Cells(reportingAtRow + (modRow - 1), 2) = modCommBest
            repWs.Cells(reportingAtRow + (modRow - 1), 3) = modCommWorst
            
'            If gblnDebugging Then
'                Debug.Print "-------------------------------------"
'                Debug.Print "Student Number: " & (modRow - 1) & " of " & responseCount
'                Debug.Print "This cell contains: " & Cell
'                Debug.Print "Module Code Found: " & modCodeFound
'                Debug.Print "Nominally, this is module: " & modNum
'                Debug.Print "STAR RATING: " & modStars
'                Debug.Print "Best comments: " & modCommBest
'                Debug.Print "Worst comments: " & modCommWorst
'            End If
        End If
    End If
Next

' Module 9
For Each Cell In module9DataRng
    If Not Cell = "" Then
        modCodeFound = Trim(Left(Cell.Text, InStr(1, Cell.Text, "-") - 1))
        If modCodeFound = moduleCode Then
            modRow = Cell.Row
            modNum = 9
            
            modStars = repWs.Cells(modRow, modNum + 29).Value2
            modCommBest = repWs.Cells(modRow, (2 * modNum) + 37).Value2
            Set modCommBestRng = repWs.Range(repWs.Cells(modRow, (2 * modNum) + 37), repWs.Cells(modRow, (2 * modNum) + 37))
            If modCommBestRng.HasFormula Then
            'If Left(modCommBest, 1) = "=" Then
                Debug.Print "----- COMMENT SANITISE"
                Debug.Print "WAS: " & modCommBestRng.Formula
                modCommBest = Replace(modCommBestRng.Formula, "=", "-")
                Debug.Print "NOW: " & modCommBest
            End If
            modCommWorst = repWs.Cells(modRow, (2 * modNum) + 38).Value2
            Set modCommWRng = repWs.Range(repWs.Cells(modRow, (2 * modNum) + 38), repWs.Cells(modRow, (2 * modNum) + 38))
            If modCommWRng.HasFormula Then
            'If Left(modCommWorst, 1) = "=" Then
                Debug.Print "----- COMMENT SANITISE"
                Debug.Print "WAS: " & modCommWRng.Formula
                modCommWorst = Replace(modCommWRng.Formula, "=", "-")
                Debug.Print "NOW: " & modCommWorst
            End If
            
            ' Then report this (same sheet, lower down)
            repWs.Cells(reportingAtRow + (modRow - 1), 1) = modStars
            repWs.Cells(reportingAtRow + (modRow - 1), 2) = modCommBest
            repWs.Cells(reportingAtRow + (modRow - 1), 3) = modCommWorst
            
            If gblnDebugging Then
                Debug.Print "-------------------------------------"
                Debug.Print "Student Number: " & (modRow - 1) & " of " & responseCount
                Debug.Print "This cell contains: " & Cell
                Debug.Print "Module Code Found: " & modCodeFound
                Debug.Print "Nominally, this is module: " & modNum
                Debug.Print "STAR RATING: " & modStars
                Debug.Print "Best comments: " & modCommBest
                Debug.Print "Worst comments: " & modCommWorst
            End If
        End If
    End If
Next

If gblnDebugging Then
    Debug.Print "3) Getting module statistical data"
End If

' IF I'M RIGHT - ERRORS BELOW!
Dim statRange As Range
cellA = Cells(reportingAtRow + 1, 1).Address(False, False)
cellB = Cells(reportingAtRow + responseCount, 1).Address(False, False)
Debug.Print repWs.Range(cellA).Value2
Debug.Print repWs.Range(cellB).Value2
Set statRange = repWs.Range(cellA & ":" & cellB)
'statRange = repWs.Range(repWs.Cells(reportingAtRow + 1, 1), repWs.Cells(reportingAtRow + responseCount, 1)).Address(False, False)
NAcount = False
If cellA = cellB Then
    Median = statRange
    Average = statRange
    If statRange.Value2 = "N/A" Then
        NAcount = 1
    Else
        NAcount = 0
    End If
Else
    If repWs.Range(cellA).Value2 = "N/A" And repWs.Range(cellB).Value2 = "N/A" Then
        NAcount = 1
        Median = "N/A"
        Average = "N/A"
    Else
        Median = Application.WorksheetFunction.Median(statRange)
        Average = Round(Application.WorksheetFunction.Average(statRange), 2)
        NAcount = countIn(statRange, "N/A")
    End If
End If
'Median = getMedian(ActiveSheet.Range(statRange))                    ' CUSTOM FUNCTION AS OLD VERSION THREW ERRORS! Median = EA.WorksheetFunction.Median(ActiveSheet.Range(statRange))
'Average = Round(getAverage(ActiveSheet.Range(statRange)), 2)        ' CUSTOM FUNCTION AS OLD VERSION THREW ERRORS! Average = Round(WorksheetFunction.Average(ActiveSheet.Range(statRange)), 2)
'NAcount = Application.WorksheetFunction.CountIf(statRange, "=N/A")

If (NAcount <> False) Or (NAcount <> 0) Then
    NAnote = gstrNAresponseDisclaimer
    NAnote = Replace(NAnote, "%NAs", NAcount)
Else
    NAnote = ""
End If

Dim counts
counts = Array(0, 0, 0, 0)
elseCount = 0
'For Each Cell In ActiveSheet.Range(statRange)
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
repWs.Cells(reportingAtRow + responseCount + 1, 1) = zeroCnt
repWs.Cells(reportingAtRow + responseCount + 1, 2) = oneCnt
repWs.Cells(reportingAtRow + responseCount + 1, 3) = twoCnt
repWs.Cells(reportingAtRow + responseCount + 1, 4) = thrCnt
repWs.Cells(reportingAtRow + responseCount + 1, 5) = ValidResponses
repWs.Cells(reportingAtRow + responseCount + 1, 6) = Average
repWs.Cells(reportingAtRow + responseCount + 1, 7) = Median

Dim refWb As Workbook
Dim refWs As Worksheet
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

Department = EA.WorksheetFunction.VLookup(repWs.Name, refWs.Range(gstrModuleLookupRng), 4, False)
School = EA.WorksheetFunction.VLookup(repWs.Name, refWs.Range(gstrModuleLookupRng), 5, False)
FHEQlevel = EA.WorksheetFunction.VLookup(repWs.Name, refWs.Range(gstrModuleLookupRng), 6, False)

' PASTE SUMMARY DATA TO SHEET "Summary Data" sumWs
'CODE   TITLE   COHORT  RESP%   AVERAGE     MEDIAN     VALID    PUBLISHED?
With sumWs
    endRow = .Range("A" & sumWs.Rows.count).End(xlUp).Row + 1
    .Cells(endRow, 1) = moduleCode
    .Cells(endRow, 2) = moduleTitle
    .Cells(endRow, 3) = cohortSize
    .Cells(endRow, 4) = responseRate
    .Cells(endRow, 5) = Average
    .Cells(endRow, 6) = Median
    .Cells(endRow, 7) = ValidResponses
    .Cells(endRow, 8) = FHEQlevel
    .Cells(endRow, 10) = Department
    .Cells(endRow, 11) = School
End With
    
If cohortSize < gintPublicationThreshold Then
    sumWs.Cells(endRow, 9) = "Not Published"
    'GoTo doNotPublish
End If

' PART 2 - CREATING A MODULE REPORT
If gblnDebugging Then
    Debug.Print "4) Create Word Doc"
End If
'Dim wrdApp As Word.Application
'Dim wrdDoc As Word.document
Set wrdApp = CreateObject("Word.Application")
wrdApp.Visible = False
Set wrdDoc = wrdApp.Documents.Add

' HEADERS AND FOOTERS AND PAGE SETUP
With wrdDoc
    .Sections(1).Headers(wdHeaderFooterPrimary).Range.Text = gstrDocTitle & " " & gstrDocYear
    .Sections(1).Footers(wdHeaderFooterPrimary).Range.Text = "CES Report for " & fullModule & " (generated: " & Format(Date, "dd-mm-yy") & " " & Format(Time, "hh.mm.ss") & ")"
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
End With

If gblnDebugging Then
    Debug.Print "5) Add module summary data"
End If

With wrdDoc
    
' Titles and course details
    .Range(0).Style = .Styles(wdStyleHeading1)
    .Content.InsertAfter gstrDocTitle & " " & gstrDocYear
    .Content.InsertParagraphAfter
    
    .Range(.Characters.count - 1).Style = .Styles(wdStyleHeading2)
    .Content.InsertAfter "MODULE-LEVEL REPORT FOR " & UCase(fullModule)
    .Content.InsertParagraphAfter
    .Content.InsertAfter gstrMoreInfo
    .Content.InsertParagraphAfter
    
    .Range(.Characters.count - 1).Style = .Styles(wdStyleNormal)
    .Content.InsertParagraphAfter
    .Content.InsertAfter "Eligible Cohort Size: " & cohortSize
    .Content.InsertParagraphAfter
    .Content.InsertAfter "Number of responses: " & responseCount
    .Content.InsertParagraphAfter
    .Content.InsertAfter "Response Rate: " & responseRate
    .Content.InsertParagraphAfter
    
    '.Range(.Characters.count - 1).Style = wrdApp.ActiveDocument.Styles(Alert)
    .Range(.Characters.count - 1).Style = .Styles(wdStyleHeading6)
    .Content.InsertAfter disclaimer
    .Content.InsertParagraphAfter
    .Content.InsertParagraphAfter
    .Range(.Characters.count - 1).InlineShapes.AddHorizontalLineStandard
    .Content.InsertParagraphAfter
End With

If gblnDebugging Then
    Debug.Print "6) Add module satisfaction tables"
End If

' ------- Overall Module Satisfaction START
wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading2)
wrdDoc.Content.InsertAfter "Overall Module Satisfaction"
wrdDoc.Content.InsertParagraphAfter

Set myRange = wrdDoc.Range(Start:=wrdDoc.Characters.count - 1, End:=wrdDoc.Characters.count)
Set overallTable = wrdDoc.Tables.Add(Range:=myRange, NumRows:=2, NumColumns:= _
    8, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
    wdAutoFitFixed)
overallTable.Style = "Grid Table 1 Light"
overallTable.ApplyStyleHeadingRows = True

' Table Headings
headerRow = Array("Question Text", "Not at all satisfied (1)", "Not very satisfied (2)", "Quite satisfied (3)", "Very satisfied (4)", "Total responses", "Mean", "Median")
For Index = 0 To 7
    Set tableCell = overallTable.Cell(1, Index + 1)
    tableCell.Range.InsertAfter Text:=headerRow(Index)
Next

' Question Title(s)
overallTable.Cell(2, 1).Range.InsertAfter Text:="Overall, how satisfied were you with '" & fullModule & "' this year?"
overallTable.Columns(1).SetWidth ColumnWidth:=gsngTotalPageWidthPoints / 2, RulerStyle:=wdAdjustNone

' Response Data
For Index = 0 To 6
    Set tableCell = overallTable.Cell(2, Index + 2)
    overallTable.Columns(Index + 2).SetWidth ColumnWidth:=(gsngTotalPageWidthPoints / 2) / 7, RulerStyle:=wdAdjustNone
    cellData = repWs.Cells(reportingAtRow + responseCount + 1, Index + 1).Text
    tableCell.Range.InsertAfter Text:=cellData
Next
' ------- Overall Module Satisfaction END
'
wrdDoc.Range(wrdDoc.Characters.count - 1).Style = wrdDoc.Styles(wdStyleHeading4)
wrdDoc.Content.InsertAfter NAnote
wrdDoc.Range(wrdDoc.Characters.count - 1).InsertParagraphAfter
wrdDoc.Range(wrdDoc.Characters.count - 1).InsertBreak (wdPageBreak)
'
' ------- Free Text Comments - START
Set bestCommRng = repWs.Range(repWs.Cells(reportingAtRow, 2), repWs.Cells(reportingAtRow + responseCount, 2))
Set worstCommRng = repWs.Range(repWs.Cells(reportingAtRow, 3), repWs.Cells(reportingAtRow + responseCount, 3))

'PRESUMING THAT COMMENTS ARE MANDATORY
Set myRange = wrdDoc.Range(Start:=wrdDoc.Characters.count - 1, End:=wrdDoc.Characters.count)
Set bestCommentsTable = wrdDoc.Tables.Add(Range:=myRange, NumRows:=responseCount + 1, NumColumns:=1, _
    DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
    wdAutoFitFixed)
bestCommentsTable.Style = "List Table 1 Light"
Set tableCell = bestCommentsTable.Cell(1, 1)
tableCell.Range.InsertAfter Text:="Please tell us the BEST thing about " & fullModule & ":"

commentNo = 1
For Each Cell In bestCommRng
    If Left(Cell.Text, 1) = "=" Then
        CellText = Replace(Cell.Text, "=", "-")
    Else
        CellText = Cell
    End If
    If Not Trim(CellText) = "" Then
        Set tableCell = bestCommentsTable.Cell(commentNo + 1, 1)
        tableCell.Range.InsertAfter Text:=CellText
        commentNo = commentNo + 1
    Else
        tableCell.Select
        Selection.Rows.Delete
        'tableCell.Rows.Delete
    End If
Next
wrdDoc.Range(wrdDoc.Characters.count - 1).InsertParagraphAfter

Set myRange = wrdDoc.Range(Start:=wrdDoc.Characters.count - 1, End:=wrdDoc.Characters.count)
Set worstCommentsTable = wrdDoc.Tables.Add(Range:=myRange, NumRows:=responseCount + 1, NumColumns:=1, _
    DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
    wdAutoFitFixed)
worstCommentsTable.Style = "List Table 1 Light"
Set tableCell = worstCommentsTable.Cell(1, 1)
tableCell.Range.InsertAfter Text:="Please tell us the WORST thing about " & fullModule & ":"

commentNo = 1
For Each Cell In worstCommRng
    If Left(Cell.Text, 1) = "=" Then
        CellText = Replace(Cell.Text, "=", "-")
    Else
        CellText = Cell
    End If
    If Not Trim(CellText) = "" Then
        Set tableCell = worstCommentsTable.Cell(commentNo + 1, 1)
        tableCell.Range.InsertAfter Text:=CellText
        commentNo = commentNo + 1
    End If
Next

' ------- Free Text Comments - END

' Output report as saved PDF document
wrdDoc.Activate
If School = "" Or School = "NO SCHOOL" Then
    School = "OTHER"
End If

sanitisedModuleTitle = Replace(Replace(Replace(Replace(Left(fullModule, 17), ":", " "), "?", ""), "(", ""), ")", "")
If cohortSize < gintPublicationThreshold Then
    'sumWs.Cells(endRow, 9) = "Not Published"
    'GoTo doNotPublish
    oName = gstrReportsFilePath & "SCHOOL REPORTS/" & School & "/DO NOT PUBLISH/DO NOT PUBLISH - " & sanitisedModuleTitle & " [" & Format(Date, "dd-mm-yy") & " " & Format(Time, "hh.mm.ss") & "].pdf"
Else
    oName = gstrReportsFilePath & "SCHOOL REPORTS/" & School & "/MODULE REPORTS/" & sanitisedModuleTitle & " [" & Format(Date, "dd-mm-yy") & " " & Format(Time, "hh.mm.ss") & "].pdf"
End If
wrdDoc.SaveAs2 FileName:=oName, FileFormat:=wdFormatPDF
wrdDoc.Close (False)
If gblnDebugging Then
    Debug.Print "Saved Report as PDF - " & oName
End If
wrdApp.Quit (False)
Set wrdApp = Nothing
'Set EA = Nothing
'GoTo endOfSub

doNotPublish:
    
    'Set wrdApp = CreateObject("Word.Application")
    'wrdApp.Visible = False
    'Set wrdDoc = wrdApp.Documents.Add
    'If School = "" Then
    '    School = "OTHER"
    'End If
    'oName = gstrReportsFilePath & "SCHOOL REPORTS/" & School & "/MODULE REPORTS/" & "NOT PUBLISHED - " & modCodeFound & " [" & Format(Date, "dd-mm-yy") & " " & Format(Time, "hh.mm.ss") & "].pdf"
    'wrdDoc.SaveAs2 FileName:=oName, FileFormat:=wdFormatPDF
    'wrdDoc.Close (False)
    'If gblnDebugging Then
    '    Debug.Print "Saved Report as PDF - " & oName
    'End If
    'wrdApp.Quit (False)
    'Set wrdApp = Nothing
    

endOfSub:
End Sub


