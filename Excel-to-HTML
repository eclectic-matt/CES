Sub aXLread()

Dim WorkRng As Range
Dim Rng As Range
Dim WS As Worksheet
Dim htmlOut As String
Dim fontCol As Long
Dim fontRGB As String
Dim rngBorCol As Long
Dim rngBorRGB As String
Dim PhysLink As String
Dim MathLink As String
Dim LinkStart As String
Dim oFolder As String
Dim xTitleId As String

PhysLink = "http://www.sussex.ac.uk/mps/internal/departments/physicsandastronomy/ug/ugcourses/2014/"
MathLink = "http://www.sussex.ac.uk/mps/internal/departments/mathematics/ug/ugcourses/2014/"

'''' AMEND THESE WHEN SWITCHING DEPARTMENTS
'''
LinkStart = PhysLink
oFolder = "G:\mps\mps-admin\Physics and Astronomy\Curriculum\2015-16\UG COURSES\Syllabus HTML Tables\"
'''
''''

On Error Resume Next

xTitleId = "Matt's XLS to HTML Tool"
Set tableRng = Application.Selection
Set tableRng = Application.InputBox("Select table range:", xTitleId, tableRng.Address, Type:=8)
Set WS = Application.ActiveSheet

Application.ScreenUpdating = False

rngWid = tableRng.Columns.count
rngHgt = tableRng.Rows.count

'htmlOut = "<table style='" & _
        "border-collapse:collapse;'>" & _
        "<tbody>     <tr>   "


htmlOut = htmlOut & "<table style='border-collapse:collapse;'><tbody> <tr>"
htmlOut = htmlOut & "<td style='text-align: right; vertical-align: bottom; font-size: 10px; font-family: Arial; color: #000000; border-top: 1px  #000000; border-bottom: 1px  #000000; border-left: 1px  #000000; border-right: 1px  #000000; width: 59.25px; height: 15px; '></td>        <td style='text-align: left; vertical-align: top; font-size: 9px; font-family: Arial; color: #000000; border-top: 1px solid #000000; border-bottom: 1px solid #000000; border-left: 1px  #000000; border-right: 1px  #000000; width: 20.25px; height: 15px; '>Period</td>        <td style='text-align: left; vertical-align: top; font-size: 9px; font-family: Arial; color: #000000; border-top: 1px solid #000000; border-bottom: 1px solid #000000; border-left: 1px  #000000; border-right: 1px  #000000; width: 35.25px; height: 15px; '>Term</td> "
htmlOut = htmlOut & "<td style='text-align: left; vertical-align: top; font-size: 9px; font-family: Arial; color: #000000; border-top: 1px solid #000000; border-bottom: 1px solid #000000; border-left: 1px  #000000; border-right: 1px  #000000; width: 75px; height: 15px; '>Syllabus Rule </td>"
htmlOut = htmlOut & "<td style='text-align: center; vertical-align: top; font-size: 9px; font-family: Arial; color: #000000; border-top: 1px solid #000000; border-bottom: 1px solid #000000; border-left: 1px  #000000; border-right: 1px  #000000; width: 30px; height: 15px; '>Credits</td>        <td style='text-align: center; vertical-align: top; font-size: 9px; font-family: Arial; color: #000000; border-top: 1px solid #000000; border-bottom: 1px solid #000000; border-left: 1px  #000000; border-right: 1px  #000000; width: 39px; height: 15px; '>Level</td>        <td style='text-align: left; vertical-align: top; font-size: 9px; font-family: Arial; color: #000000; border-top: 1px solid #000000; border-bottom: 1px solid #000000; border-left: 1px  #000000; border-right: 1px  #000000; width: 309.75px; height: 15px; '>"
htmlOut = htmlOut & "Module Title (Link)</td>       <td style='text-align: center; vertical-align: top; font-size: 9px; font-family: Arial; color: #000000; border-top: 1px solid #000000; border-bottom: 1px solid #000000; border-left: 1px  #000000; border-right: 1px  #000000; width: 48px; height: 15px; '>Code</td>        <td style='text-align: left; vertical-align: top; font-size: 8px; font-family: Arial; color: #000000; border-top: 1px solid #000000; border-bottom: 1px solid #000000; border-left: 1px  #000000; border-right: 1px solid #000000; width: 51px; height: 15px; '>Pre-R</td>        <td style='text-align: left; vertical-align: top; font-size: 8px; font-family: Arial; color: #000000; border-top: 1px solid #000000; border-bottom: 1px solid #000000; border-left: 1px solid #000000; border-right: 1px  #000000; width: 51px; height: 15px; '>Co-R</td></tr><tr>"

oldRow = tableRng.Row
merged = False

For Each Rng In tableRng

    rngCol = Rng.Column
    rngRow = Rng.Row
    If rngRow <> oldRow Then
        htmlOut = htmlOut & "</tr><tr>      "
    End If
    oldRow = rngRow
    
    rngMergeAddr = Rng.MergeArea.Address
    rngAddr = Rng.Address
    rngMergeStart = ""
    rngMergeStart = Left(rngMergeAddr, (InStr(1, rngMergeAddr, ":", vbTextCompare) - 1))
    
    If (rngMergeStart <> "" And rngAddr <> rngMergeStart) Then
            ' Create no cell
            
    Else
        If rngAddr = rngMergeStart Then
            rngColspan = Rng.MergeArea.Columns.count
            rngRowspan = Rng.MergeArea.Rows.count
            ' span (merge) cell
            htmlOut = htmlOut & "<td rowspan='" & rngRowspan & _
                "' colspan='" & rngColspan & "' style ='"
        Else
            ' Normal Cell
            htmlOut = htmlOut & "   <td style='"
            
        End If
            '''''''''
            ''' ALIGNMENT
            '''''''''
            fontHor = Rng.HorizontalAlignment
            Select Case fontHor
                Case xlLeft
                    fontHor = "left"
                Case xlCenter
                    fontHor = "center"
                Case xlRight
                    fontHor = "right"
                Case Else
                    fontHor = "right"
            End Select
            htmlOut = htmlOut & "text-align: " & fontHor & "; "
    
            fontVer = Rng.VerticalAlignment
            Select Case fontVer
                Case xlTop
                    fontVer = "top"
                Case xlCenter
                    fontVer = "center"
                Case xlBottom
                    fontVer = "bottom"
                Case Else
                    fontVer = "bottom"
            End Select
            htmlOut = htmlOut & "vertical-align: " & fontVer & "; "
            
            '''''''''
            ''' FONT
            '''''''''
            fontSiz = Rng.Font.Size
            htmlOut = htmlOut & "font-size: " & fontSiz & "px; "
            
            fontFam = Rng.Font.Name
            htmlOut = htmlOut & "font-family: " & fontFam & "; "
            
            fontStk = Rng.Font.Strikethrough
            If fontStk = False Then
                fontStk = ""
            Else
                fontStk = "line-through"
            End If
            fontUnd = Rng.Font.Underline
            If fontUnd = xlUnderlineStyleNone Then
                fontUnd = ""
            Else
                fontUnd = "underline"
            End If
            If (fontUnd <> "" Or fontStk <> "") Then
                htmlOut = htmlOut & "text-decoration: " & fontUnd & " " & fontStk & "; "
            End If
            
            fontSty = Rng.Font.FontStyle
            If fontSty = "Regular" Or xlAutomatic Then
                fontSty = ""
            Else
                
                fontBld = Rng.Font.Bold
                If fontBld = False Then
                    fontBld = ""
                Else
                    fontBld = "bold"
                End If
                
                fontIta = Rng.Font.Italic
                If fontIta = False Then
                    fontIta = ""
                Else
                    fontIta = "italic"
                End If
                    
                fontSty = fontBld & " " & fontIta
                If fontSty <> "" Then
                    htmlOut = htmlOut & "font-style: " & fontSty & "; "
                End If
                
            End If
        
            fontCol = Rng.Font.Color
            fontRGB = Color_to_RGB(fontCol)
            
            htmlOut = htmlOut & "color: " & fontRGB & "; "
            
            bBorRGB = tBorRGB = rBorRGB = lBorRGB = ""
    
            tBorSty = lineConvert(Rng.Borders(xlEdgeTop).LineStyle)
            tBorWgt = weightConvert(Rng.Borders(xlEdgeTop).Weight)
            Dim tBorCol As Long
            tBorCol = Rng.Borders(xlEdgeTop).Color
            If tBorCol <> 0 Or xlAutomatic Then
                tBorRGB = Color_to_RGB(tBorCol)
            End If
            htmlOut = htmlOut & "border-top: " & tBorWgt & " " & tBorSty & " " & tBorRGB & "; "
            
            bBorSty = lineConvert(Rng.Borders(xlEdgeBottom).LineStyle)
            bBorWgt = weightConvert(Rng.Borders(xlEdgeBottom).Weight)
            Dim bBorCol As Long
            bBorCol = Rng.Borders(xlEdgeBottom).Color
            If bBorCol <> 0 Or xlAutomatic Then
                bBorRGB = Color_to_RGB(bBorCol)
            End If
            htmlOut = htmlOut & "border-bottom: " & bBorWgt & " " & bBorSty & " " & bBorRGB & "; "
           
            lBorSty = lineConvert(Rng.Borders(xlEdgeLeft).LineStyle)
            lBorWgt = weightConvert(Rng.Borders(xlEdgeLeft).Weight)
            Dim lBorCol As Long
            lBorCol = Rng.Borders(xlEdgeLeft).Color
            If lBorCol <> 0 Or xlAutomatic Then
                lBorRGB = Color_to_RGB(lBorCol)
            End If
            htmlOut = htmlOut & "border-left: " & lBorWgt & " " & lBorSty & " " & lBorRGB & "; "
            
            rBorSty = lineConvert(Rng.Borders(xlEdgeRight).LineStyle)
            rBorWgt = weightConvert(Rng.Borders(xlEdgeRight).Weight)
            Dim rBorCol As Long
            rBorCol = Rng.Borders(xlEdgeRight).Color
            If rBorCol <> 0 Or xlAutomatic Then
                rBorRGB = Color_to_RGB(rBorCol)
            End If
            htmlOut = htmlOut & "border-right: " & rBorWgt & " " & rBorSty & " " & rBorRGB & "; "
                
            Dim bgCol As Long
            bgCol = Rng.Interior.Color
            If bgCol <> 16777215 And bgCol <> xlAutomatic Then
                bgRGB = Color_to_RGB(bgCol)
                htmlOut = htmlOut & "background: " & bgRGB & "; "
            End If
            
            cellWid = Rng.Width
            cellHgt = Rng.Height
            htmlOut = htmlOut & "width: " & cellWid & _
                    "px; height: " & cellHgt & "px; "
            
            Text = Rng.Text
            txtDir = Rng.Orientation
            txtClass = textDirectionConvert(txtDir)
            htmlOut = htmlOut & txtClass
            
            'hypOn = rng.Hyperlinks.Count
            'formOn = rng.HasFormula
            
            'If (InStr(1, rng.Formula, "HYPERLINK") > 0) Then
            If (Rng.Column = 7) And (Rng.Row > tableRng.Row) Then
            
                Set sequenceTable = Worksheets("Sheet1").Range("$A$2:$B$181")
                courseCode = Range("H" & tableRng.Row).Value
                moduleCode = Range("H" & Rng.Row).Value
                
                sequence = Application.WorksheetFunction.VLookup(moduleCode, sequenceTable, 2, False)
                '''
                'MsgBox (sequence)
                '''
                hypAddr = LinkStart & courseCode & "/" & sequence
                
                htmlOut = htmlOut & "'><a href='" & _
                    hypAddr & "'>" & Text & "</a></td>    "
            Else
                htmlOut = htmlOut & "'>" & Text & "</td>     "
            End If
    End If
    
Next

htmlOut = htmlOut & "</tr></tbody></table>"

Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Dim oFile As Object
oName = oFolder & courseCode & " [" & Format(Date, "dd-mm-yy") & " " & Format(Time, "hh.mm.ss") & "].txt"
Set oFile = fso.CreateTextFile(oName)
oFile.WriteLine htmlOut
oFile.Close
Set fso = Nothing
Set oFile = Nothing
MsgBox ("HTML table output to:" & vbCrLf & oName)

WorkRng.ClearContents
WorkRng.Select
Application.ScreenUpdating = True

End Sub

Sub textDir()

xTitleId = "Text Direction Tool"
Set tableRng = Application.Selection
Set tableRng = Application.InputBox("Select table range:", xTitleId, tableRng.Address, Type:=8)
Set WS = Application.ActiveSheet
txtDir = tableRng.Orientation
Select Case (txtDir)
    Case xlVertical
        Direction = "vertical"
        
        textdirstyle = "-ms-writing-mode: lr-tb; -webkit-writing-mode: horizontal-tb; -moz-writing-mode: horizontal-tb; -ms-writing-mode: horizontal-tb; writing-mode: horizontal-tb;"
    Case xlUpward
        Direction = "upwards"
        textdirstyle = "-ms-writing-mode: tb-lr; -webkit-writing-mode: vertical-lr; -moz-writing-mode: vertical-lr; -ms-writing-mode: vertical-lr; writing-mode: vertical-lr;"
    Case xlDownward
        Direction = "downwards"
        textdirstyle = "-ms-writing-mode: tb-rl; -webkit-writing-mode: vertical-rl; -moz-writing-mode: vertical-rl; -ms-writing-mode: vertical-rl; writing-mode: vertical-rl;"
    Case (-90 < txtDir < 90)
        Direction = "angle of " & txtDir
    Case Else
        'Case xlAutomatic Or xlHorizontal
        Direction = ""
        
    End Select

MsgBox (Direction & vbNewLine & textdirstyle)

End Sub


Function textDirectionConvert(dir) As String

Select Case (dir)
    Case xlVertical
        'Direction = "vertical"
        textDirectionConvert = ""
        'textDirectionConvert = "word-wrap:break-word; word-break: break-all; text-overflow:clip; overflow:hidden; display:block; top:0; width:0.5em; height:auto;"
        
    Case xlUpward
        'Direction = "upwards"
        textDirectionConvert = "-webkit-transform:rotate(270deg); -moz-transform:rotate(270deg); -o-transform: rotate(270deg); -ms-transform:rotate(270deg); transform: rotate(270deg); white-space:nowrap; display:block; bottom:0; width:20px; height:20px;"
        
    Case xlDownward
        'Direction = "downwards"
        textDirectionConvert = "-webkit-transform:rotate(90deg); -moz-transform:rotate(90deg); -o-transform: rotate(90deg); -ms-transform:rotate(90deg); transform: rotate(90deg); white-space:nowrap; display:block; bottom:0; width:20px; height:20px;"
        
    Case Else
        'Case xlAutomatic Or xlHorizontal
        'Direction = ""
        textDirectionConvert = ""
    End Select
    Exit Function

'MsgBox (Direction & vbNewLine & textdirstyle)

End Function
Function lineConvert(Line) As String

    Select Case Line
        Case xlContinuous
            lineConvert = "solid"
        Case xlDouble
            lineConvert = "double"
        Case xlDash
            lineConvert = "dashed"
        Case xlDot
            lineConvert = "dotted"
        Case Else
            lineConvert = ""
    End Select
    Exit Function
End Function

Function weightConvert(Weight) As String

    Select Case Weight
        Case xlHairline
            weightConvert = "1px"
        Case xlThin
            weightConvert = "1px"
        Case xlMedium
            weightConvert = "5px"
        Case xlThick
            weightConvert = "10px"
        Case Else
            weightConvert = ""
    End Select
    Exit Function
End Function

Function Color_to_RGB(Color As Long) As String

    r = Application.WorksheetFunction.Dec2Hex(Color Mod 256)
    If r = 0 Then
        r = "00"
    End If
    
    G = Application.WorksheetFunction.Dec2Hex((Color \ 256) Mod 256)
    If G = 0 Then
          G = "00"
    End If

    B = Application.WorksheetFunction.Dec2Hex((Color \ 256 \ 256) Mod 256)
    If B = 0 Then
          B = "00"
    End If
  
    Color_to_RGB = "#" & r & G & B
    Exit Function
    
End Function
