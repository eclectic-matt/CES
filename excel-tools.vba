'----------------------------
' Useful Excel Functions!
'----------------------------

Sub CloseAllWorkbooks()
For Each wb In Application.Workbooks
    wb.Close False
Next
'Set Application = Nothing
End Sub

Function parseStrToArr(ByRef str As String, delimiter)
Dim Arr() As String
Arr = Split(str, delimiter, Compare:=vbTextCompare)
parseStrToArr = Arr
End Function

'----------------------------
' @name  getStdDevSam
' @descr An attempt to avoid the WORKSHEET FUNCTION version of StdDev.S
' @usage Takes a range and returns the SAMPLE standard deviation
' @form  Math.SQR( Sum(Value - Mean)^2 / (SampleSize - 1) )
'----------------------------
Function getStdDevSam(ByVal stdRng As Range)
    numItems = stdRng.count
    mean = getAverage(stdRng)
    If mean = "N/A" Then
        deviation = "N/A"
    Else
        devSum = 0
        For Each Cell In stdRng
            If Not Len(Cell) = 0 Then
                devSum = devSum + (Cell - mean) ^ 2
            Else
                numItems = numItems - 1
            End If
        Next
        deviation = Math.Sqr((devSum / (numItems - 1)))
    End If
    getStdDevSam = deviation
End Function

'----------------------------
' @name  getStdDevPop
' @descr An attempt to avoid the WORKSHEET FUNCTION version of StdDev.P
' @usage Takes a range and returns the POPULATION standard deviation
' @form  Math.SQR( Sum(Value - Mean)^2 / (PopulationSize) )
'----------------------------
Function getStdDevPop(ByVal stdRng As Range)
    numItems = stdRng.count
    mean = getAverage(stdRng)
    If mean = "N/A" Then
        deviation = "N/A"
    Else
        devSum = 0
        For Each Cell In stdRng
            If Not Len(Cell) = 0 Then
                devSum = devSum + (Cell - mean) ^ 2
            Else
                numItems = numItems - 1
            End If
        Next
        deviation = Math.Sqr((devSum / numItems))
    End If
    getStdDevPop = deviation
End Function

'----------------------------
' @name  getMedian
' @descr An attempt to avoid the WORKSHEET FUNCTION version of Median
' @usage Takes a range and returns the median value (sorted as range as no VBA Arr Sort)
'----------------------------
Function getMedian(ByRef medRng As Range)
    medRng.Sort key1:=medRng, order1:=xlAscending, MatchCase:=False, Orientation:=xlTopToBottom, SortMethod:=xlPinYin, Header:=xlNo, DataOption1:=xlSortTextAsNumbers
    Dim numItems As Integer
    'TESTING COUNT
    'Debug.Print "Rng Cnt: " & medRng.count
    'Debug.Print "EA.CntA: " & Application.WorksheetFunction.CountA(medRng)
    'Debug.Print "Blk Cnt: " & countIn(medRng, "")
    numItems = medRng.count - countIn(medRng, "")
    If numItems = 0 Then
        midAvg = "N/A"
    Else
        If numItems = 1 Then
            midAvg = medRng.Value2
        Else
            Dim Arr() As Variant
            Arr = medRng
            If numItems Mod 2 = 0 Then
                middle1 = Arr(Round(numItems / 2, 0), 1)
                middle2 = Arr(Round(numItems / 2, 0) + 1, 1)
                midAvg = (middle1 + middle2) / 2
            Else
                midAvg = Arr(Round(numItems / 2, 0), 1)
            End If
        End If
    End If
    getMedian = midAvg
End Function

'----------------------------
' @name  getAverage
' @descr An attempt to avoid the WORKSHEET FUNCTION version of Average
' @usage Takes a range and returns the arithmetic average (mean) value
'----------------------------
Function getAverage(ByRef avgRng As Range)
    numItems = avgRng.count
    avgSum = 0
    For Each Cell In avgRng
        If IsNumeric(Cell) And Not Len(Cell) = 0 Then
            avgSum = avgSum + Cell
        Else
            numItems = numItems - 1
        End If
    Next
    If numItems = 0 Then
        avg = "N/A"
    Else
        avg = avgSum / numItems
    End If
    getAverage = avg
End Function

'----------------------------
' @name  getMax
' @descr An attempt to avoid the WORKSHEET FUNCTION version of MAX
' @usage Takes a range and returns the largest number found (converts strings to DOUBLES)
'----------------------------
Function getMax(ByRef maxRng As Range)
Dim maxFound
maxFound = MinDouble
For Each Cell In maxRng
    If IsNumeric(Cell) And Len(Cell) <> 0 Then
        NumToCheck = CDbl(Cell)
        If NumToCheck > maxFound Then
            maxFound = NumToCheck
        End If
    End If
Next
If maxFound = MinDouble And countIn(maxRng, "0") = 0 Then
    getMax = "N/A"
Else
    getMax = maxFound
End If
End Function

' MaxDouble and MinDouble from: http://www.tushar-mehta.com/publish_train/xl_vba_cases/1003%20MinMaxVals.shtml
Function MinDouble() As Double
    MinDouble = -1.79769313486231E+308
End Function
Function MaxDouble() As Double
    MaxDouble = 1.79769313486231E+308
End Function


'----------------------------
' @name  getRange
' @descr An attempt to avoid the WORKSHEET FUNCTION version of Range
' @usage Takes a sheet-range and returns the range (diff between smallest and largest)
'----------------------------
Function getRange(ByRef rangeRng As Range)
minVal = getMin(rangeRng)
maxVal = getMax(rangeRng)
If minVal = "N/A" Or maxVal = "N/A" Then
    Rng = "N/A"
Else
    Rng = maxVal - minVal
End If
getRange = Rng
End Function


'----------------------------
' @name  getMin
' @descr An attempt to avoid the WORKSHEET FUNCTION version of MIN
' @usage Takes a range and returns the smallest number found (converts strings to DOUBLES)
'----------------------------
Function getMin(ByRef minRng As Range)
Dim minFound
minFound = MaxDouble
For Each Cell In minRng
    If IsNumeric(Cell) And Len(Cell) <> 0 Then
        NumToCheck = CDbl(Cell)
        If NumToCheck < minFound Then
            minFound = NumToCheck
        End If
    End If
Next
If minFound = MaxDouble And countIn(minRng, "0") = 0 Then
    getMin = "N/A"
Else
    getMin = minFound
End If
End Function



'----------------------------
' @name  countIn
' @descr An attempt to avoid the WORKSHEET FUNCTION version of CountIf
' @usage Takes a range and returns the number of cells matched by the condition (all converted to strings)
'----------------------------
Function countIn(ByRef countRng As Range, condition)
count = 0
If Not VarType(condition) = vbString Then
    condition = CStr(condition)
End If
For Each Cell In countRng
    If Not VarType(Cell) = vbString Then
        Cell = CStr(Cell)
    End If
    If Cell = condition Then
        count = count + 1
    End If
Next
countIn = count
End Function

'----------------------------
' @name  openNamedSheet
' @descr Opens the worksheet (in the active workbook) with the user entered sheet name
' @usage Run in Workbook and enter a sheet name in the InputBox
'----------------------------
Sub openNamedSheet()
m = InputBox("Please enter a sheet name", "Open Named Sheet")
If Not m = vbNo Or m = vbCancel Or Len(m) = 0 Then
    ActiveWorkbook.Sheets(m).Activate
Else
    MsgBox "Cancelled"
End If
End Sub

'----------------------------
' @name  startTimer
' @descr STARTS the timer to test macro run times
' @usage CALL at the beginning of a module and then CALL endTimer before "End Sub"
'----------------------------
Sub startTimer()
tStart = Timer
End Sub

'----------------------------
' @name  endTimer
' @descr ENDS the timer and displays macro run time
' @usage CALL before "End Sub" - make sure startTimer called at beginning!
'----------------------------
Sub endTimer()
timeTaken = Timer - tStart
'tMins = Round(timeTaken / 60, 0)
tMins = Int(timeTaken / 60)
tSecs = Format(timeTaken Mod 60, "00")

'tMins = timeTaken Mod 60
'tSecs = timeTaken - (60 * tMins)
MsgBox ("This process took " & timeTaken & " secs" & _
        vbNewLine & vbNewLine & _
        tMins & " minutes and " & tSecs & " seconds." & _
        vbCr & tMins & ":" & tSecs)
End Sub

'----------------------------
' @name  showLastRow
' @descr Brings up a MsgBox showing the "last" row in column A of the current sheet
' @usage Run on a worksheet to show the number of filled rows in column A
'----------------------------
Sub showLastRow()
    lastRow = ActiveSheet.Cells(Rows.count, "A").End(xlUp).Row
    MsgBox ("Last Row: " & lastRow)
End Sub

'----------------------------
' @name  complexAlphabetise
' @descr Sorts worksheets alphabetically by testing each one against each other
' @usage Run on a worksheet to sort - but not very efficient! Takes O(n)^2 - 1 to complete!!!!
' @src   https://www.extendoffice.com/documents/excel/629-excel-sort-sheets.html
'----------------------------
Sub complexAlphabetise()
For i = 1 To moduleWb.Sheets.count
    For j = 1 To moduleWb.Sheets.count - 1
        If UCase$(moduleWb.Sheets(j).Name) > UCase$(moduleWb.Sheets(j + 1).Name) Then
            moduleWb.Sheets(j).Move After:=moduleWb.Sheets(j + 1)
        End If
    Next
Next
End Sub

'----------------------------
' @name  sortSheets
' @descr Sorts worksheets alphabetically by listing sheet names in a new WS then using a sort function
' @usage Run on a worksheet to sort - more efficient for large # of worksheets!
' @src   http://www.excelforum.com/l/359252-sort-excel-sheets-into-alphabetical-order.html
'----------------------------
Sub SortSheets()
    Call startTimer
Dim sht As Worksheet
Dim mySht As Worksheet
Dim i As Integer
Dim endRow As Long
Dim shtNames As Range
Dim Cell As Range

Set mySht = Sheets.Add
mySht.Move Before:=Sheets(1)
For i = 2 To Sheets.count
    mySht.Cells(i - 1, 1).Value = Sheets(i).Name
Next i
endRow = mySht.Cells(Rows.count, 1).End(xlUp).Row
Set shtNames = mySht.Range(Cells(1, 1), Cells(endRow, 1))
shtNames.Sort key1:=Range("A1"), order1:=xlAscending, Header:= _
xlNo, OrderCustom:=1

i = 2
For Each Cell In shtNames
    Sheets(Cell.Value).Move Before:=Sheets(i)
    i = i + 1
Next Cell

Application.DisplayAlerts = False
mySht.Delete
Application.DisplayAlerts = True
    Call endTimer
    
End Sub

'----------------------------
' @name  IsWorkBookOpen
' @arg   FileName: Takes a filename (in the full form, C:/Documents/FileName.xlsx)
' @descr Tests if this filename is already open in Excel
' @usage Assign a boolean variable to IsWorkBookOpen, can then IF to either Set or Open this WB w/o errors
' @src   http://stackoverflow.com/questions/9373082/detect-whether-excel-workbook-is-already-open/9373914#9373914
'----------------------------
Function IsWorkBookOpen(FileToTest As String)
    Dim ff As Long, ErrNo As Long
    If Workbooks.CanCheckOut(FileName:=FileToTest) = False Then
        IsWorkBookOpen = False
        Debug.Print FileToTest & " -> Locked file detected!"
    End If
        
    On Error Resume Next
    ff = FreeFile()
    Open FileToTest For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False
    Case 70:   IsWorkBookOpen = True
    Case Else: Error ErrNo
    End Select
End Function

'----------------------------
' @name  Col_Letter
' @arg   lngCol: Takes a column number (long)
' @descr Returns the column letter from A (1) - XFD (16384)
' @usage Use in VBA, or as =PERSONAL.XLSB!Col_Letter(COLUMN())
' @src   http://stackoverflow.com/a/12797190
'----------------------------
Function Col_Letter(lngCol As Long) As String
    Col_Letter = Split(Cells(, lngCol).Address, "$")(1)
End Function


Sub createLoadsOfWorkbooks()

total = 150
For i = 1 To total Step 1
    ActiveWorkbook.Sheets.Add Before:=Worksheets(Worksheets.count)
    
    With Worksheets(Worksheets.count)
        .Range("B:B").ClearContents
        .Range(.Range("A2"), .Cells(.Rows.count, "A").End(xlUp)).SpecialCells( _
                xlCellTypeConstants, 23).Offset(0, 1).FormulaR1C1 = _
                "=RANDBETWEEN(1+50*(ROW()-2),50+50*(ROW()-2))"
        .Range("B:B").Value = .Range("B:B").Value
        .Range("B1").Value = "Random Number"
    End With

Next

End Sub

Sub InsertCheckBoxes()

Dim Rng As Range
Dim WorkRng As Range
Dim WS As Worksheet
' Not used currently:
Dim chk As CheckBox

On Error Resume Next

xTitleId = "Matt's Checkbox Tool"

Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Select the range to add checkboxes to:", xTitleId, WorkRng.Address, Type:=8)
Set WS = Application.ActiveSheet

Application.ScreenUpdating = False

For Each Rng In WorkRng
    With WS.CheckBoxes.Add(Rng.Left, Rng.Top, Rng.Width, Rng.Height)
        .Characters.Text = Rng.Value
    End With
Next

' To link checkboxes (unclear if necessary for this purpose)
For Each chk In ActiveSheet.CheckBoxes
   With chk
      .LinkedCell = .TopLeftCell.Offset(0, 0).Address
   End With
Next chk

WorkRng.ClearContents
WorkRng.Select
Application.ScreenUpdating = True

MsgBox ("Checkboxes Inserted")

End Sub


Sub DeleteBlankRows()

Dim Rng As Range
Dim WorkRng As Range

On Error Resume Next

xTitleId = "Delete Blank Rows Tool"

Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Select the range to check", xTitleId, WorkRng.Address, Type:=8)

Application.ScreenUpdating = False

For Each Rng In WorkRng

    If WorksheetFunction.CountA(Cells(Rng.Row, 1)) = 0 Then
        Rows(Rng.Row).Delete
    End If

Next Rng

Application.ScreenUpdating = True

MsgBox ("END - Blank Rows Deleted")

End Sub

Sub DeleteWSRows()

Dim Rng As Range
Dim WorkRng As Range

On Error Resume Next

xTitleId = "Delete Workshops Tool"
strToFind = "WORKSHOP"

Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Select the range to check", xTitleId, WorkRng.Address, Type:=8)

Application.ScreenUpdating = False

For Each Rng In WorkRng

    If Cells(Rng.Row, 8).Text = strToFind Then
        Rows(Rng.Row).Delete
    End If

Next Rng

Application.ScreenUpdating = True

MsgBox ("END - WS Deleted")

End Sub


Sub CopyDown()

Dim Rng As Range
Dim WorkRng As Range
Dim cellVal As String
Dim Record As Integer
'Dim total As Integer
Dim PctDone As Single

On Error Resume Next

Complete = False
tStart = Timer
xTitleId = "Fill Down Cells Tool"

Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Select the range to fill down" & vbNewLine & "i.e. if a cell has a value, copy this to the cell below to allow filters to work correctly", xTitleId, WorkRng.Address, Type:=8)

Application.ScreenUpdating = False
'Application.ScreenUpdating = True
Call ShowProgressForm

Record = 0

total = WorkRng.CountLarge
'Call ProgBar(total)
    
For Each Rng In WorkRng

    cellVal = Rng.Value
    If Len(cellVal) > 0 Then
        
       If Rng.Offset(RowOffset:=1, columnOffset:=0).Value = "" Then
       'Or Rng.Offset(rowOffset:=1, columnOffset:=0).Value Is Null
            Rng.Offset(RowOffset:=1, columnOffset:=0).Value = cellVal
            Record = Record + 2
            ' Update the percentage completed.
            PctDone = Record / total
            ' Call subroutine that updates the progress bar.
            Call UpdateProgressBar(PctDone, Record, total)
        End If
        
    End If

Next Rng

Rows("2:2").Select
Selection.Copy
Rows("3:500").Select
Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
Columns("F:F").Select
'Selection.NumberFormat = "0"
Selection.NumberFormat = "############"

Application.ScreenUpdating = True
'Completed:
t = Timer - tStart
Complete = True
With ProgressForm
    .FrameProgress.Caption = "COMPLETED."
    .LabelProgress.Caption = "TASK COMPLETE."
    .Label2.Caption = "SUCCESS - " & Denominator & " records updated."
    .Label1.Caption = "Time taken: approx. " & Round(t, 1) & " seconds."
End With
ProgressForm.Show
t = Timer - tStart
'MsgBox ("MACRO TOOK " & t & " secs")

Call Unload(ProgressForm)



End Sub

Sub ProgBar(total)

If Complete = True Then GoTo Completed:
tStart = Timer

' Local vars - passed to ProgressBar subs as "num" and "den"
Dim Numerator As Integer
Dim Denominator As Integer
Dim PctDone As Single

Application.ScreenUpdating = False

TimeLeft = ""

' Loop through cells.
For r = 1 To RowMax
    
    For c = 1 To ColMax
        Numerator = Numerator + 1
    Next c

    ' Update the percentage completed.
    PctDone = Numerator / Denominator

    ' Call subroutine that updates the progress bar.
    Call UpdateProgressBar(PctDone, Numerator, Denominator)

Next r

Completed:

Complete = True
With ProgressForm
    .FrameProgress.Caption = "COMPLETED."
    .LabelProgress.Caption = "TASK COMPLETE."
    .Label2.Caption = "SUCCESS - " & Denominator & " records updated."
    .Label1.Caption = "Time taken: approx. " & Round(t, 1) & " seconds."
End With
ProgressForm.Show
t = Timer - tStart
MsgBox ("Just Completed in " & t & " secs")

'tStart = Timer
't = Timer - tStart
'Do While t < timeOut
'    ProgressForm.LabelProgress.Caption = "CLOSING IN " & Round(timeOut - t, 1) & " SECONDS"
'    'DoEvents
'    t = Timer - tStart
'Loop

'ProgressForm.Hide
Call Unload(ProgressForm)

'MsgBox ("Finally Completed.")


End Sub

Sub HighlightBlanks()
Dim Rng As Range
Dim WorkRng As Range
On Error Resume Next
xTitleId = "Highlight Blanks Tool"
Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Select the range to check", xTitleId, WorkRng.Address, Type:=8)
Application.ScreenUpdating = False
For Each Rng In WorkRng
    If Len(Rng.Text) < 1 Then
        Rng.Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End If
Next Rng
Application.ScreenUpdating = True
MsgBox ("END - Highlighted")
End Sub

Sub FillDownTest()

Dim Rng As Range
Dim ToFill As Range
Dim WorkRng As Range
Dim i As Integer

On Error Resume Next

xTitleId = "Highlight Blanks Tool"

Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Select the range to check", xTitleId, WorkRng.Address, Type:=8)

'Application.ScreenUpdating = False
'Application.Worksheets(1).Select

For Each Rng In WorkRng

    If Len(Rng.Text) > 1 Then
    
        Set ToFill = Rng
        ToFill.Select
        Selection.Copy
        'MsgBox (Selection)
        
        For i = 0 To 25 Step 1
            
            If Len(Rng.Offset(i, 0).Text) < 1 Then
            
                MsgBox (Rng.Offset(i, 0).Address + " - Value: " + Rng.Offset(i, 0).Text)
                'Rng.Offset(i, 0).Select
                Range(Rng.Offset(i, 0).Address).PasteSpecial Paste:=xlPasteValues
                'ActiveSheet.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                '    :=False, Transpose:=False
                    
            Else
                
                GoTo skip:
                
            End If
        
        Next
    
    End If

skip:

Next Rng

Application.ScreenUpdating = True

MsgBox ("END - Filled")

End Sub

Sub AutoFillDown()
Dim LR As Long
LR = Range("A" & Rows.count).End(xlUp).Row
Range("A" & LR).AutoFill Destination:=Range("A" & LR).Resize(2)
End Sub


Sub ProtectAll()
Dim WS As Worksheet
Dim wb As Worksheets
Dim count As Integer
count = ActiveWorkbook.Worksheets.count
Dim i As Integer
For i = 1 To count
    Set WS = ActiveWorkbook.Worksheets(i)
    WS.Protect
Next
End Sub

Sub CheckBoxCull()
Dim chk As CheckBox
Dim i As Integer
i = 0
For Each chk In ActiveSheet.CheckBoxes
      chk.Delete
      i = i + 1
Next chk
MsgBox ("Checkbox Cull Complete. " & i & " boxes deleted.")
End Sub

Sub UncheckAll()
Dim chk As CheckBox
Dim i As Integer
i = 0
For Each chk In ActiveSheet.CheckBoxes
      chk.Value = False
      i = i + 1
Next chk
MsgBox ("UNChecking Complete. " & i & " boxes UNticked.")
End Sub

Sub CheckToggleRange()
Dim isect As Object
Dim Rng As Range
Dim WorkRng As Range
Dim WS As Worksheet
Dim chk As CheckBox

On Error Resume Next

xTitleId = "Matt's Check Toggler"

Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Select the toggle checkboxes:", xTitleId, WorkRng.Address, Type:=8)
Set WS = Application.ActiveSheet

Application.ScreenUpdating = False

For Each chk In ActiveSheet.CheckBoxes

    Set isect = Application.Intersect(WorkRng, Range(chk.LinkedCell))
    If isect Is Nothing Then
        'MsgBox ("Ranges do not intersect at " & chk.LinkedCell)
        GoTo skip:
    End If
    
chk.Delete
    
skip:
Set isect = Nothing
Next chk

MsgBox ("Toggled checkboxes")

End Sub

Sub CheckAll()
Dim chk As CheckBox
Dim i As Integer
i = 0
    For Each chk In ActiveSheet.CheckBoxes
          'If Left(chk.LinkedCell, 2) = "$A" Then
            chk.Value = True
            i = i + 1
          'End If
    Next chk
    MsgBox ("Checking Complete. " & i & " boxes ticked.")
End Sub

Sub ShowChecked()

Dim m As Variant
Dim Name As String
Dim Invitees(100) As String
Dim Idx As Integer
Dim total As Integer
Idx = 0
Name = "" & vbNewLine & ""

m = MsgBox("This will display each checked box in a list (this might take a while)." & vbNewLine & vbNewLine & "Continue at your own peril!!!", vbOKCancel)
returnCol = 7

If m = vbCancel Then
    Exit Sub
End If

For Each chk In ActiveSheet.CheckBoxes
    With chk
        If .Value = 1 Then
            Invitees(Idx) = Cells(Range(.LinkedCell).Row, returnCol).Value
            Idx = Idx + 1
        End If
    End With
Next
total = Idx
' List all invitees in MsgBox
For Idx = 0 To total
    Name = Name & vbNewLine & Invitees(Idx)
Next
MsgBox ("Send invites to " & Name)
End Sub


Sub clearBlanks()

Dim checkRng As Range
Dim xTitleId As String
Set checkRng = Application.Selection
xTitleId = "Matt's Blank Clearer"
Set checkRng = Application.InputBox("Select range to clear blanks:", xTitleId, checkRng.Address, Type:=8)
Dim testCell As Object
Dim testValue As String

For Each testCell In checkRng
    
    With testCell
        
        .Select
        testValue = Selection.Value
        'MsgBox (testValue)
        If (testValue = "") Or (Selection = "") Or (testValue = 0) Then
            Selection.ClearContents
        End If
    
    End With
    
Next

MsgBox ("Cleared")

End Sub

Sub FillDownForFilters()

Dim Rng As Range
Dim Below As Range
Dim WorkRng As Range
Dim WS As Worksheet
Dim off As Integer
Dim cellVal As String

debugging = True
maxRowsToFill = 10

On Error Resume Next

xTitleId = "Fill Down Cells Tool"

Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Select the range to fill down" & vbNewLine & "i.e. if a cell has a value, copy this to the cell below to allow filters to work correctly", xTitleId, WorkRng.Address, Type:=8)
Set WS = Application.ActiveSheet

'Application.ScreenUpdating = False

For Each Rng In WorkRng

    If Len(Rng.Value) > 0 Then
        
        cellVal = Rng.Value
        If debugging Then
            Debug.Print "---------------------------"
            Debug.Print "Value to check = " & cellVal
        End If
                
        For off = 1 To maxRowsToFill
        
            Set Below = Rng.Offset(off, 0)
            belowVal = Below.Text
            
            If debugging Then
                Debug.Print "Value below = " & belowVal
            End If
            
            If (belowVal = "") Or (belowVal Is Null) Then
                Below = cellVal
                If debugging Then
                    Debug.Print "Inserting " & cellVal
                End If
            Else
                GoTo nextOff
            End If
            
nextOff:
        Next off
        
    End If

Next Rng

'Application.ScreenUpdating = True

MsgBox ("END - Filled Down")

End Sub

' Usage as in "=PERSONAL.XLSB!findBlankCol(ROW(),TRUE)"

Function findBlankCol(Row As Integer, returnType As Boolean)

Dim thisWb As Workbook
Dim thisWs As Worksheet
Dim BlankCol As Integer
Dim BlankColName As String

Set thisWb = ActiveWorkbook
Set thisWs = thisWb.Sheets(ActiveSheet.Name)
thisWs.Select

With thisWs
    BlankCol = .Cells(Row, .Columns.count).End(xlToLeft).Column + 1
End With

Select Case (returnType)
    Case True
        'Number
        findBlankCol = BlankCol

    Case False
        'Text
        BlankColName = Chr(64 + BlankCol)
        findBlankCol = BlankColName
End Select

End Function
