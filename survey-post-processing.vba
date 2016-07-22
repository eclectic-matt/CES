Sub sanitiseAllCourseSheets()

Call startTimer

Dim reportingWorkbook As Workbook
Set reportingWorkbook = ActiveWorkbook
Dim courseCodeToProcess As String

For Each Sheet In ActiveWorkbook.Sheets
    courseCodeToProcess = Sheet.Name
    If (Not courseCodeToProcess = "Course Reports") And (Not courseCodeToProcess = "Summary Data") Then
        If gblnDebugging Then
            Debug.Print "---------------------------"
            Debug.Print "Sanitising - " & courseCodeToProcess
            Debug.Print "---------------------------"
        End If
        Call sanitiseCourseSheetForChecking(courseCodeToProcess)
    End If
Next

Call endTimer

End Sub

Sub sanitiseCourseSheetForChecking(sheetName As String)

With Application.ActiveWorkbook.Sheets(sheetName)
    firstBlockEndRow = .Range("A" & Rows.count).End(xlUp).Row
    .Range("B2:BZ" & firstBlockEndRow).Delete
    .Rows(firstBlockEndRow + 1 & ":" & firstBlockEndRow + 7).EntireRow.Delete
    secondBlockEndRow = .Range("A" & Rows.count).End(xlUp).Row
    .Range("B" & firstBlockEndRow + 1 & ":BZ" & secondBlockEndRow).Delete
    .Rows(secondBlockEndRow + 1 & ":" & secondBlockEndRow + 7).EntireRow.Delete
    .Columns("C:F").EntireColumn.Delete
    .Cells(1, 1).Value = "RespondentID"
    .Cells(1, 2).Value = "Free Text Comments"
    .Cells(1, 3).Value = "Action Taken"
    .Columns("A:C").AutoFit
    .Columns("B:B").ColumnWidth = 60
    .Range("A1:C1").Font.Bold = True
    With .Range("A2:D" & secondBlockEndRow)
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
End With
End Sub




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
        Call sanitiseModuleSheetForChecking(moduleCodeToProcess)
    End If
Next

Call endTimer

End Sub

Sub sanitiseModuleSheetForChecking(sheetName As String)

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

