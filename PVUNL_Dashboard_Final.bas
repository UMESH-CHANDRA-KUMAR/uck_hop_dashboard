Attribute VB_Name = "PVUNL_Dashboard"
Option Explicit

Private Const DAILY_FILE As String = "Daily-updates.xlsx"
Private Const TASKS_FILE As String = "Calendar-tasks.xlsx"
Private Const DATA_FILE As String = "Data-file.xlsx"
Private Const REPORT_SHEET As String = "PVUNL Daily Report"

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
    MsgBox "Report generation failed:" & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Private Function OpenWorkbookIfNeeded(ByVal basePath As String, ByVal shortName As String) As Workbook
    Dim fullPath As String, wb As Workbook
    fullPath = basePath & Application.PathSeparator & shortName
    If Len(Dir$(fullPath, vbNormal Or vbReadOnly Or vbHidden Or vbSystem)) = 0 Then
        Err.Raise vbObjectError + 2, , "Source file not found:" & vbCrLf & fullPath
    End If
    On Error Resume Next
    Set wb = Workbooks(shortName)
    On Error GoTo 0
    If wb Is Nothing Then Set wb = Workbooks.Open(Filename:=fullPath, ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)
    Set OpenWorkbookIfNeeded = wb
End Function

Private Function PrepareReportSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(REPORT_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = REPORT_SHEET
    End If
    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10
    ws.Columns("A:K").ColumnWidth = 15
    Set PrepareReportSheet = ws
End Function

Private Sub BuildReport(ByVal ws As Worksheet, ByVal dailyWb As Workbook, ByVal tasksWb As Workbook, ByVal dataWb As Workbook, ByVal reportDate As Date)
    Dim navy As Long: navy = RGB(0, 90, 156)
    Dim blue As Long: blue = RGB(0, 120, 212)
    Dim orange As Long: orange = RGB(255, 102, 0)
    Dim green As Long: green = RGB(40, 167, 69)
    Dim purple As Long: purple = RGB(142, 68, 173)
    Dim updatesWs As Worksheet, criticalWs As Worksheet, chimneyWs As Worksheet, tasksWs As Worksheet, milestonesWs As Worksheet

    Set updatesWs = GetRequiredSheet(dailyWb, "Updates")
    Set criticalWs = GetRequiredSheet(dailyWb, "Critical areas")
    Set chimneyWs = GetRequiredSheet(dailyWb, "Chimney")
    Set tasksWs = GetRequiredSheet(tasksWb, "Tasks")
    Set milestonesWs = GetRequiredSheet(dataWb, "Milestone")

    MergeTitle ws.Range("A1:K1"), "PVUNL Project Dashboard", navy, 20
    MergeSubtitle ws.Range("A2:K2"), "Daily Summary for " & Format(reportDate, "dd-mmm-yyyy") & " | Generated " & Format(Date, "dd-mmm-yyyy")
    Dim r As Long: r = 4
    MergeSection ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)), "Status Tiles - as on " & Format(reportDate, "dd-mmm-yyyy"), navy
    r = r + 1

    MakeTile ws, 1, r, 3, r + 3, "MANPOWER STATUS", GetManpowerForDate(updatesWs, reportDate), "Manpower status as on " & Format(reportDate, "dd-mmm-yyyy"), navy
    MakeTile ws, 4, r, 6, r + 3, "ACC CHART", "Min target: " & CStr(GetTarget(criticalWs, "ACC")), "ACC critical-area target", blue
    MakeChimneyTile ws, 7, r, 11, r + 3, chimneyWs, purple
    r = r + 5

    MergeSection ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)), "Daily Major Updates for " & Format(reportDate, "dd-mmm-yyyy"), navy
    r = r + 1
    AddUpdate ws, r, "ACC", GetUpdateForDate(updatesWs, "ACC", reportDate), GetTarget(criticalWs, "ACC"), blue: r = r + 2
    AddUpdate ws, r, "Boiler", GetUpdateForDate(updatesWs, "Boiler", reportDate), GetTarget(criticalWs, "Boiler - NDHT Joints"), orange: r = r + 2
    AddUpdate ws, r, "PCP Joint", GetUpdateForDate(updatesWs, "PCP Joint", reportDate), GetTarget(criticalWs, "PCP - P91/92"), green: r = r + 2

    MergeSection ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)), "This Week's Targets (" & Format(MondayOfWeek(Date), "dd-mmm") & " to " & Format(MondayOfWeek(Date) + 6, "dd-mmm-yyyy") & ")", navy
    r = r + 1
    AddWeeklyTasks ws, r, tasksWs
    r = r + 8

    MergeSection ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)), "Key Milestones", navy
    r = r + 1
    AddMilestones ws, r, milestonesWs

    ws.PageSetup.PrintArea = "$A$1:$K$" & CStr(r + 4)
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = 1
    ws.PageSetup.Zoom = False
End Sub

Private Function GetRequiredSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetRequiredSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
    If GetRequiredSheet Is Nothing Then Err.Raise vbObjectError + 10, , "Worksheet not found: " & sheetName & " in " & wb.Name
End Function

Private Sub MergeTitle(ByVal rg As Range, ByVal text As String, ByVal colour As Long, ByVal fontSize As Long)
    rg.Merge: rg.Value = text: rg.Font.Size = fontSize: rg.Font.Bold = True: rg.Font.Color = vbWhite: rg.Interior.Color = colour: rg.HorizontalAlignment = xlCenter
End Sub

Private Sub MergeSubtitle(ByVal rg As Range, ByVal text As String)
    rg.Merge: rg.Value = text: rg.Font.Size = 11: rg.Font.Color = RGB(90, 90, 90): rg.HorizontalAlignment = xlCenter
End Sub

Private Sub MergeSection(ByVal rg As Range, ByVal text As String, ByVal colour As Long)
    rg.Merge: rg.Value = text: rg.Font.Bold = True: rg.Font.Color = vbWhite: rg.Interior.Color = colour: rg.HorizontalAlignment = xlLeft: rg.VerticalAlignment = xlCenter: rg.RowHeight = 23
End Sub

Private Sub MakeTile(ByVal ws As Worksheet, ByVal c1 As Long, ByVal r1 As Long, ByVal c2 As Long, ByVal r2 As Long, ByVal heading As String, ByVal tileValue As String, ByVal note As String, ByVal colour As Long)
    ws.Range(ws.Cells(r1, c1), ws.Cells(r1, c2)).Merge: ws.Cells(r1, c1).Value = heading: ws.Cells(r1, c1).Font.Bold = True: ws.Cells(r1, c1).Font.Color = vbWhite: ws.Cells(r1, c1).Interior.Color = colour: ws.Cells(r1, c1).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(r1 + 1, c1), ws.Cells(r1 + 2, c2)).Merge: ws.Cells(r1 + 1, c1).Value = tileValue: ws.Cells(r1 + 1, c1).Font.Bold = True: ws.Cells(r1 + 1, c1).Font.Size = 15: ws.Cells(r1 + 1, c1).Font.Color = RGB(0, 90, 156): ws.Cells(r1 + 1, c1).Interior.Color = RGB(244, 247, 251): ws.Cells(r1 + 1, c1).HorizontalAlignment = xlCenter: ws.Cells(r1 + 1, c1).VerticalAlignment = xlCenter
    ws.Range(ws.Cells(r1 + 3, c1), ws.Cells(r2, c2)).Merge: ws.Cells(r1 + 3, c1).Value = note: ws.Cells(r1 + 3, c1).Font.Size = 9: ws.Cells(r1 + 3, c1).Interior.Color = RGB(244, 247, 251): ws.Cells(r1 + 3, c1).HorizontalAlignment = xlCenter
    ApplyBorder ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
End Sub

Private Sub MakeChimneyTile(ByVal ws As Worksheet, ByVal c1 As Long, ByVal r1 As Long, ByVal c2 As Long, ByVal r2 As Long, ByVal source As Worksheet, ByVal colour As Long)
    ws.Range(ws.Cells(r1, c1), ws.Cells(r1, c2)).Merge: ws.Cells(r1, c1).Value = "CHIMNEY SEGMENT STATUS": ws.Cells(r1, c1).Font.Bold = True: ws.Cells(r1, c1).Font.Color = vbWhite: ws.Cells(r1, c1).Interior.Color = colour: ws.Cells(r1, c1).HorizontalAlignment = xlCenter
    Dim i As Long, outputText As String
    For i = 2 To source.Cells(source.Rows.Count, 1).End(xlUp).Row
        If Len(CStr(source.Cells(i, 1).Value)) > 0 Then outputText = outputText & CStr(source.Cells(i, 1).Value) & ": " & source.Cells(i, 2).Text & vbCrLf
    Next i
    ws.Range(ws.Cells(r1 + 1, c1), ws.Cells(r2, c2)).Merge: ws.Cells(r1 + 1, c1).Value = outputText: ws.Cells(r1 + 1, c1).Font.Size = 9: ws.Cells(r1 + 1, c1).Interior.Color = RGB(244, 247, 251): ws.Cells(r1 + 1, c1).HorizontalAlignment = xlCenter: ws.Cells(r1 + 1, c1).VerticalAlignment = xlCenter: ws.Cells(r1 + 1, c1).WrapText = True
    ApplyBorder ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
End Sub

Private Sub AddUpdate(ByVal ws As Worksheet, ByVal r As Long, ByVal category As String, ByVal updateText As String, ByVal target As Variant, ByVal accent As Long)
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)).Merge: ws.Cells(r, 1).Value = category & "  |  Min. Target: " & CStr(target) & vbCrLf & IIf(Len(updateText) = 0, "No update logged for the report date.", updateText): ws.Cells(r, 1).WrapText = True: ws.Cells(r, 1).Interior.Color = RGB(244, 247, 251): ws.Cells(r, 1).Borders(xlEdgeLeft).Color = accent: ws.Cells(r, 1).Borders(xlEdgeLeft).Weight = xlThick: ws.Rows(r).RowHeight = 34
End Sub

Private Sub AddWeeklyTasks(ByVal ws As Worksheet, ByVal r As Long, ByVal source As Worksheet)
    ws.Cells(r, 1).Value = "Date": ws.Cells(r, 2).Value = "Activity": ws.Cells(r, 8).Value = "Status": ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)).Interior.Color = RGB(0, 90, 156): ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)).Font.Color = vbWhite: ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)).Font.Bold = True
    Dim i As Long, outRow As Long, d As Date, startDate As Date, endDate As Date: startDate = MondayOfWeek(Date): endDate = startDate + 6: outRow = r + 1
    For i = 2 To source.Cells(source.Rows.Count, 1).End(xlUp).Row
        If IsDate(source.Cells(i, 1).Value) Then
            d = DateValue(source.Cells(i, 1).Value)
            If d >= startDate And d <= endDate Then
                ws.Cells(outRow, 1).Value = d: ws.Cells(outRow, 1).NumberFormat = "dd-mmm": ws.Range(ws.Cells(outRow, 2), ws.Cells(outRow, 7)).Merge: ws.Cells(outRow, 2).Value = source.Cells(i, 2).Value: ws.Range(ws.Cells(outRow, 8), ws.Cells(outRow, 11)).Merge: ws.Cells(outRow, 8).Value = IIf(Len(CStr(source.Cells(i, 3).Value)) = 0, "Pending", source.Cells(i, 3).Value): outRow = outRow + 1
            End If
        End If
    Next i
    ApplyBorder ws.Range(ws.Cells(r, 1), ws.Cells(Application.Max(r + 1, outRow - 1), 11))
End Sub

Private Sub AddMilestones(ByVal ws As Worksheet, ByVal r As Long, ByVal source As Worksheet)
    Dim lastRow As Long, total As Long, i As Long, colIndex As Long
    lastRow = source.Cells(source.Rows.Count, 1).End(xlUp).Row: total = lastRow - 1: If total <= 0 Then Exit Sub
    For i = 2 To lastRow
        colIndex = 1 + Int((i - 2) * 11 / total): ws.Cells(r, colIndex).Value = CStr(source.Cells(i, 1).Value) & vbCrLf & CStr(source.Cells(i, 2).Value): ws.Cells(r, colIndex).WrapText = True: ws.Cells(r, colIndex).HorizontalAlignment = xlCenter: ws.Cells(r, colIndex).VerticalAlignment = xlCenter: ws.Cells(r, colIndex).Font.Size = 8: ws.Cells(r, colIndex).Interior.Color = RGB(244, 247, 251)
    Next i
    ApplyBorder ws.Range(ws.Cells(r, 1), ws.Cells(r + 1, 11))
End Sub

Private Function GetManpowerForDate(ByVal source As Worksheet, ByVal targetDate As Date) As String
    Dim updateText As String: updateText = GetUpdateForDate(source, "ACC", targetDate): GetManpowerForDate = "No data"
    If Len(updateText) > 0 Then
        Dim re As Object, matches As Object: Set re = CreateObject("VBScript.RegExp"): re.Pattern = "manpower(\s+deployment)?\s*[-:]?\s*([0-9,]+)": re.IgnoreCase = True
        If re.Test(updateText) Then Set matches = re.Execute(updateText): GetManpowerForDate = matches(0).SubMatches(1)
    End If
End Function

Private Function GetUpdateForDate(ByVal source As Worksheet, ByVal category As String, ByVal targetDate As Date) As String
    Dim i As Long
    For i = 2 To source.Cells(source.Rows.Count, 1).End(xlUp).Row
        If IsDate(source.Cells(i, 1).Value) Then
            If DateValue(source.Cells(i, 1).Value) = targetDate And Trim(CStr(source.Cells(i, 2).Value)) = category Then GetUpdateForDate = CStr(source.Cells(i, 3).Value): Exit Function
        End If
    Next i
End Function

Private Function GetTarget(ByVal source As Worksheet, ByVal headerName As String) As Variant
    Dim col As Long, i As Long: col = 0
    For i = 1 To source.Cells(1, source.Columns.Count).End(xlToLeft).Column
        If LCase(Trim(CStr(source.Cells(1, i).Value))) = LCase(headerName) Then col = i: Exit For
    Next i
    If col = 0 Then GetTarget = "N/A": Exit Function
    For i = 2 To source.Cells(source.Rows.Count, 1).End(xlUp).Row
        If LCase(Trim(CStr(source.Cells(i, 1).Value))) = "target" Then GetTarget = source.Cells(i, col).Value: Exit Function
    Next i
    GetTarget = "N/A"
End Function

Private Function MondayOfWeek(ByVal d As Date) As Date
    MondayOfWeek = d - Weekday(d, vbMonday) + 1
End Function

Private Sub ApplyBorder(ByVal rg As Range)
    With rg.Borders: .LineStyle = xlContinuous: .Color = RGB(210, 220, 230): .Weight = xlThin: End With
End Sub
