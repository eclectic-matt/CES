'----------------------------
Public Const gblnDebugging As Boolean = True                                               ' Change this to "True" to get debug.print status messages
'----------------------------
' Fixed sheet names
'----------------------------
Public Const gstrOrigWbName As String = "Responses - FINAL data including partial responses 22.6.16"                        ' The name of the responses Workbook (origWb)
Public Const gstrOrigWsName As String = "Course_Evaluation_Survey_201516"                   ' The name of the responses Worksheet (origWs)
Public Const gstrRefWbName As String = "CES - Reference Sheet (DO NOT DELETE)"              ' The name of the REFERENCE Workbook (refWb)
Public Const gstrStudyYearName As String = "EDITED - FORMATTED STUDENT SURVEY DATA WITH YEAR OF PROGRAMME"
'----------------------------
' Fixed File/Folder Paths
'----------------------------
Public Const gstrStudyYearFile As String = "G:\ar\ar_adqe\Shared\Enhancement\Qualtrics\EDITED - FORMATTED STUDENT SURVEY DATA WITH YEAR OF PROGRAMME.xlsx"
Public Const gstrReportsFilePath As String = "G:\ar\ar_adqe\Shared\Enhancement\Qualtrics\GENERATED REPORTS\"         'THEN ADD "COURSE REPORTS\" or "MODULE REPORTS\"
Public Const gstrOrigWbFile As String = "G:\ar\ar_adqe\Shared\Enhancement\Qualtrics\GENERATED REPORTS\" & gstrOrigWbName & ".xlsx"
Public Const gstrRefWbFile As String = "G:\ar\ar_adqe\Shared\Enhancement\Qualtrics\" & gstrRefWbName & ".xlsx"
'----------------------------
' Fixed ranges
'----------------------------
Public Const gstrModuleLookupRng As String = "B8:G754"                                    ' The range of cells in the REF sheet for modules
Public Const gintModuleLookupReturn As Integer = 4                                        ' The column number (not VLOOKUP) to return cohort size for modules
Public Const gstrCourseLookupRng As String = "B8:H193"                                    ' The range of cells in the REF sheet for courses
'Public Const gintCourseLookupReturn As Integer = 4                                       ' NOT USED - NEED TO CHANGE BASED ON STUDY YEAR - The column number (not VLOOKUP) to return cohort size for courses
Public Const gstrModule18Rng As String = "L:S"  'NOT USED?!?!
Public Const gstrModule9Rng As String = "AA:AA" 'NOT USED?!?!
'----------------------------
' Report titles and text
'----------------------------
Public Const gstrDocTitle As String = "COURSE EVALUATION SURVEY REPORT"                   ' The title for the reports
Public Const gstrDocYear As String = "2015/16"                                            ' The year to display in the reports
Public Const gstrThresholdDisclaimer As String = "Please note: as the number of responses (%RESP) is below the threshold for this cohort ($THRE), these responses should be treated with caution."
Public Const gstrNAresponseDisclaimer As String = "Please note: some responses (%NAs) were not valid, which have been discounted from the statistical data"
Public Const gstrMoreInfo As String = "For more information about the Course Evaluation Surveys, please visit http://www.sussex.ac.uk/adqe/enhancement/studentengagement/studentvoice/ces"
'----------------------------
' Fixed numbers (records/width/column numbers)
'----------------------------
Public Const gintLowestStudyYear As Integer = 1
Public Const gintHighestStudyYear As Integer = 2
Public Const gintPublicationThreshold As Integer = 4                                  ' The minimum number of responses in order to report
Public Const gsngTotalPageWidthPoints As Single = 697.9                               ' The width, in points, of a landscape Word doc
Public Const gintTotalRecords As Integer = 6545                                       ' The total number of student records to be processed
Public Const gintCourseDataStartCol As Integer = 57                                   ' The first column containing course data
Public Const gintCourseDataColCount As Integer = 21                                   ' The number of columns of course data
Public Const gintModuleDataStartCol As Integer = 12                                   ' The number of the column where module code/titles are shown
Public Const gintModuleDataEndCol As Integer = 19                                     ' The END column for module code/titles
Public Const gintModuleStartRow As Integer = 2                                        ' The first row of module response data
Public Const gintHeaderRows As Integer = 2                                            ' The number of header rows in the responses sheet (origWs)
Public Const gintSheetDataRows As Integer = 1                                         ' The number of header rows to INSERT into split module/course sheets
Public Const gintStatRows As Integer = 7                                              ' The number of STATISTICAL rows inserted into course/module sheets (count0,1,2,3,valid,Mean,Median)

Public Const gintCourseDataEndCol As Integer = 78                                     ' NOT USED - The end column for course data
Public Const gsngTotalPageHeightPoints As Single = 451.3                              ' NOT USED - The height, in points, of a landscape Word doc
'----------------------------
' Table Heading Arrays
' NOTE: VBA cannot have constant arrays, so these are strings which are parsed to arrays by the ExcelTools.ParseStrToArr function
'----------------------------
Public Const gstrCourseSatisfactionHeadings As String = "Question Text,Very satisfied (1),Quite satisfied (2),Not very satisfied (3),Not at all satisfied (4),Valid responses,Mean,Median"
Public Const gstrCourseContentHeadings As String = "Question Text,Extremely (1),Moderately (2),Slightly (3),Not at all (4),Valid responses,Mean,Median"
Public Const gstrCourseAssessmentHeadings As String = "Question Text,Excellent (1),Good (2),Fair (3),Poor (4),Valid responses,Mean,Median"
Public Const gstrCourseWorkloadHeadings As String = "Question Text,Too much (1),About right (2),Too little (3),Valid responses,Mean,Median"
Public Const gstrCourseSkillsHeadings As String = "Question Text,1 Star,2 Stars,3 Stars,4 Stars,Valid responses,Mean,Median"
Public Const gstrCoursePreparationHeadings As String = "Question Text,Extremely (1),Moderately (2),Slightly (3),Not at all (4),Valid responses,Mean,Median"
Public Const gstrCourseResourcesHeadings As String = "Question Text,Excellent (1),Good (2),Fair (3),Poor (4),Valid responses,Mean,Median"
Public Const gstrCourseOrganisationHeadings As String = "Question Text,Strongly agree (1),Slightly agree (2),Slightly disagree (3),Strongly disagree (4),Valid responses,Mean,Median"
