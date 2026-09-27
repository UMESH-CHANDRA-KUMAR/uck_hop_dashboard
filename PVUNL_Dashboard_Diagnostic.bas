Attribute VB_Name = "PVUNL_Dashboard"
Option Explicit

Private Const DAILY_FILE As String = "Daily-updates.xlsx"
Private Const TASKS_FILE As String = "Calendar-tasks.xlsx"
Private Const DATA_FILE As String = "Data-file.xlsx"
Private Const REPORT_SHEET As String = "PVUNL Daily Report"

Public Sub DiagnosePVUNLFiles()
    Dim msg As String, basePath As String
    basePath = ThisWorkbook.Path
    msg = "Macro workbook:" & vbCrLf & ThisWorkbook.FullName & vbCrLf & vbCrLf
    msg = msg & "Folder:" & vbCrLf & basePath & vbCrLf & vbCrLf
    msg = msg & CheckFile(basePath, DAILY_FILE)
    msg = msg & CheckFile(basePath, TASKS_FILE)
    msg = msg & CheckFile(basePath, DATA_FILE)
    MsgBox msg, vbInformation, "PVUNL File Diagnostic"
End Sub

Private Function CheckFile(ByVal basePath As String, ByVal fileName As String) As String
    Dim p As String
    p = basePath & Application.PathSeparator & fileName
    If Len(Dir$(p)) = 0 Then
        CheckFile = "NOT FOUND: " & p & vbCrLf
    Else
        CheckFile = "FOUND: " & p & vbCrLf
    End If
End Function

Public Sub GeneratePVUNLReport()
    Dim reportDate As Date: reportDate = Date - 1
    Dim oldAlerts As Boolean, oldScreen As Boolean
    oldAlerts = Application.DisplayAlerts: oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim basePath As String: basePath = ThisWorkbook.Path
    If Len(basePath) = 0 Then Err.Raise vbObjectError + 1, , "Save the macro workbook before running the report."

    Dim dailyWb As Workbook, tasksWb As Workbook, dataWb As Workbook
    Set dailyWb = OpenWorkbookIfNeeded(basePath, DAILY_FILE)
    Set tasksWb = OpenWorkbookIfNeeded(basePath, TASKS_FILE)
    Set dataWb = OpenWorkbookIfNeeded(basePath, DATA_FILE)

    Dim wsReport As Worksheet: Set wsReport = PrepareReportSheet()
    BuildReport wsReport, dailyWb, tasksWb, dataWb, reportDate

    Dim pdfPath As String
    pdfPath = basePath & Application.PathSeparator & "PVUNL_Daily_Summary_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
    wsReport.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    wsReport.Activate
    MsgBox "Report generated successfully:" & vbCrLf & pdfPath, vbInformation

CleanExit:
    Application.DisplayAlerts = oldAlerts
    Application.ScreenUpdating = oldScreen
    Exit Sub
Failed:
    MsgBox "Report generation failed:" & vbCrLf & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & "Run DiagnosePVUNLFiles first.", vbCritical
    Resume CleanExit
End Sub

Private Function OpenWorkbookIfNeeded(ByVal basePath As String, ByVal shortName As String) As Workbook
    Dim fullPath As String, wb As Workbook
    fullPath = basePath & Application.PathSeparator & shortName
    If Len(Dir$(fullPath)) = 0 Then Err.Raise vbObjectError + 2, , "Source file not found:" & vbCrLf & fullPath
    On Error Resume Next
    Set wb = Workbooks(shortName)
    On Error GoTo 0
    If wb Is Nothing Then Set wb = Workbooks.Open(Filename:=fullPath, ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)
    Set OpenWorkbookIfNeeded = wb
End Function

Private Function PrepareReportSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets(REPORT_SHEET): On Error GoTo 0
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)): ws.Name = REPORT_SHEET
    ws.Cells.UnMerge: ws.Cells.Clear: ws.Cells.Font.Name = "Segoe UI": ws.Cells.Font.Size = 10: ws.Columns("A:K").ColumnWidth = 15
    Set PrepareReportSheet = ws
End Function

Private Sub BuildReport(ByVal ws As Worksheet, ByVal dailyWb As Workbook, ByVal tasksWb As Workbook, ByVal dataWb As Workbook, ByVal reportDate As Date)
    Dim navy As Long: navy = RGB(0, 90, 156)
    MergeTitle ws.Range("A1:K1"), "PVUNL Project Dashboard", navy, 20
    MergeSubtitle ws.Range("A2:K2"), "Daily Summary for " & Format(reportDate, "dd-mmm-yyyy") & " | Generated " & Format(Date, "dd-mmm-yyyy")
    Dim r As Long: r = 4
    MergeSection ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)), "Status Tiles - as on " & Format(reportDate, "dd-mmm-yyyy"), navy
    r = r + 1
    MakeTile ws, 1, r, 3, r + 3, "MANPOWER STATUS", "Data available", "Manpower status as on " & Format(reportDate, "dd-mmm-yyyy"), navy
    MakeTile ws, 4, r, 6, r + 3, "ACC CHART", "Min target", "ACC critical-area target", RGB(0, 120, 212)
    MakeTile ws, 7, r, 11, r + 3, "CHIMNEY SEGMENT STATUS", "Data available", "Chimney status", RGB(142, 68, 173)
    ws.PageSetup.PrintArea = "$A$1:$K$12": ws.PageSetup.Orientation = xlLandscape: ws.PageSetup.FitToPagesWide = 1: ws.PageSetup.FitToPagesTall = 1: ws.PageSetup.Zoom = False
End Sub

Private Sub MergeTitle(ByVal rg As Range, ByVal text As String, ByVal colour As Long, ByVal fontSize As Long): rg.Merge: rg.Value = text: rg.Font.Size = fontSize: rg.Font.Bold = True: rg.Font.Color = vbWhite: rg.Interior.Color = colour: rg.HorizontalAlignment = xlCenter: End Sub
Private Sub MergeSubtitle(ByVal rg As Range, ByVal text As String): rg.Merge: rg.Value = text: rg.Font.Size = 11: rg.Font.Color = RGB(90, 90, 90): rg.HorizontalAlignment = xlCenter: End Sub
Private Sub MergeSection(ByVal rg As Range, ByVal text As String, ByVal colour As Long): rg.Merge: rg.Value = text: rg.Font.Bold = True: rg.Font.Color = vbWhite: rg.Interior.Color = colour: rg.HorizontalAlignment = xlLeft: End Sub
Private Sub MakeTile(ByVal ws As Worksheet, ByVal c1 As Long, ByVal r1 As Long, ByVal c2 As Long, ByVal r2 As Long, ByVal heading As String, ByVal tileValue As String, ByVal note As String, ByVal colour As Long): ws.Range(ws.Cells(r1, c1), ws.Cells(r1, c2)).Merge: ws.Cells(r1, c1).Value = heading: ws.Cells(r1, c1).Font.Bold = True: ws.Cells(r1, c1).Font.Color = vbWhite: ws.Cells(r1, c1).Interior.Color = colour: ws.Cells(r1, c1).HorizontalAlignment = xlCenter: ws.Range(ws.Cells(r1 + 1, c1), ws.Cells(r2 - 1, c2)).Merge: ws.Cells(r1 + 1, c1).Value = tileValue: ws.Cells(r1 + 1, c1).Interior.Color = RGB(244, 247, 251): ws.Cells(r1 + 1, c1).HorizontalAlignment = xlCenter: ws.Range(ws.Cells(r2, c1), ws.Cells(r2, c2)).Merge: ws.Cells(r2, c1).Value = note: ws.Cells(r2, c1).Interior.Color = RGB(244, 247, 251): ws.Cells(r2, c1).HorizontalAlignment = xlCenter: End Sub
