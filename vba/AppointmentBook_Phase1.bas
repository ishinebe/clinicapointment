Option Explicit

'============================================================
' ClinicAppointment
' Module: AppointmentBook
' Version: 2026.07.07-Phase1-A1J46-hotfix
'
' Important:
' - One-day template range is fixed to Template!A1:J46.
' - Do not use UsedRange for Template because stray formatting can expand
'   the copied area and create hundreds of printed pages.
'============================================================

Private Const SHEET_TEMPLATE As String = "Template"
Private Const SHEET_TEMPLATE_DRAFT As String = "TemplateDraft"
Private Const SHEET_SETTINGS As String = "Settings"
Private Const SHEET_OUTPUT As String = "Output"

Private Const SETTINGS_YEAR_CELL As String = "B2"
Private Const SETTINGS_MONTH_CELL As String = "B3"

Private Const TEMPLATE_ONE_DAY_RANGE As String = "A1:J46"
Private Const BLOCK_GAP_ROWS As Long = 2

Private Const DRAFT_TIME_START_HOUR As Long = 9
Private Const DRAFT_TIME_END_HOUR As Long = 19
Private Const DRAFT_HATCH_START_HOUR As Long = 18
Private Const NEW_TIME_COL As Long = 1
Private Const OLD_MINUTE_COL As Long = 8
Private Const FIRST_TIME_ROW As Long = 7
Private Const LAST_TIME_ROW As Long = 46

Public Sub GenerateAppointmentBook_Phase1()
    GenerateAppointmentBookCore "Phase 1"
End Sub

Public Sub GenerateAppointmentBook_Phase2()
    GenerateAppointmentBookCore "Phase 2"
End Sub

Private Sub GenerateAppointmentBookCore(ByVal phaseName As String)

    On Error GoTo ErrorHandler

    Dim wsT As Worksheet
    Dim wsS As Worksheet
    Dim wsO As Worksheet

    Set wsT = GetSheetOrError(SHEET_TEMPLATE)
    Set wsS = GetSheetOrError(SHEET_SETTINGS)
    Set wsO = GetSheetOrError(SHEET_OUTPUT)

    Dim targetYear As Long
    Dim targetMonth As Long

    targetYear = CLng(Val(wsS.Range(SETTINGS_YEAR_CELL).Value))
    targetMonth = CLng(Val(wsS.Range(SETTINGS_MONTH_CELL).Value))

    If targetYear < 2000 Or targetYear > 2100 Or targetMonth < 1 Or targetMonth > 12 Then
        MsgBox "Enter year in Settings!B2 and month in Settings!B3." & vbCrLf & _
               "Example: B2=2027, B3=1", vbExclamation
        Exit Sub
    End If

    Dim templateRange As Range
    Set templateRange = GetTemplateRange(wsT)

    If templateRange Is Nothing Then
        MsgBox "Template sheet does not contain a one-day design.", vbExclamation
        Exit Sub
    End If

    If templateRange.Rows.Count <> 46 Or templateRange.Columns.Count <> 10 Then
        MsgBox "Template range is unexpected: " & templateRange.Address(False, False) & vbCrLf & _
               "Expected: " & TEMPLATE_ONE_DAY_RANGE, vbCritical
        Exit Sub
    End If

    Dim previousScreenUpdating As Boolean
    Dim previousDisplayAlerts As Boolean
    Dim previousEnableEvents As Boolean
    Dim appStateCaptured As Boolean

    previousScreenUpdating = Application.ScreenUpdating
    previousDisplayAlerts = Application.DisplayAlerts
    previousEnableEvents = Application.EnableEvents
    appStateCaptured = True

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    ClearOutput wsO
    CopyTemplateColumnWidths wsT, wsO, templateRange

    Dim daysInMonth As Long
    daysInMonth = Day(DateSerial(targetYear, targetMonth + 1, 0))

    Dim blockHeight As Long
    blockHeight = templateRange.Rows.Count

    Dim pasteRow As Long
    Dim d As Long
    Dim currentDate As Date

    pasteRow = 1

    For d = 1 To daysInMonth
        currentDate = DateSerial(targetYear, targetMonth, d)

        CopyTemplateBlock wsT, wsO, templateRange, pasteRow
        ReplaceTemplateDateIfPossible wsO, templateRange, pasteRow, currentDate

        pasteRow = pasteRow + blockHeight + BLOCK_GAP_ROWS
    Next d

    CopyTemplatePrintSettings wsT, wsO, templateRange, pasteRow - BLOCK_GAP_ROWS - 1

    wsO.Activate

    Application.EnableEvents = previousEnableEvents
    Application.DisplayAlerts = previousDisplayAlerts
    Application.ScreenUpdating = previousScreenUpdating

    MsgBox phaseName & " complete: Template!" & TEMPLATE_ONE_DAY_RANGE & _
           " copied for " & targetYear & "/" & targetMonth & ".", vbInformation
    Exit Sub

ErrorHandler:
    If appStateCaptured Then
        Application.EnableEvents = previousEnableEvents
        Application.DisplayAlerts = previousDisplayAlerts
        Application.ScreenUpdating = previousScreenUpdating
    Else
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
    End If

    MsgBox "Error occurred in " & phaseName & "." & vbCrLf & _
           "Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbCritical

End Sub

Private Function GetTemplateRange(ByVal ws As Worksheet) As Range

    ' A1:J46 is the confirmed one-day appointment-book block.
    ' UsedRange is intentionally not used because stray formatting can expand it.
    Set GetTemplateRange = ws.Range(TEMPLATE_ONE_DAY_RANGE)

End Function

Private Sub ClearOutput(ByVal ws As Worksheet)

    ws.Cells.Clear

    On Error Resume Next
    ws.ResetAllPageBreaks
    On Error GoTo 0

End Sub

Private Sub CopyTemplateBlock(ByVal wsT As Worksheet, ByVal wsO As Worksheet, ByVal templateRange As Range, ByVal pasteRow As Long)

    Dim dst As Range
    Set dst = wsO.Cells(pasteRow, templateRange.Column)

    templateRange.Copy
    dst.PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False

    CopyTemplateRowHeights wsT, wsO, templateRange, pasteRow

End Sub

Private Sub CopyTemplateColumnWidths(ByVal wsT As Worksheet, ByVal wsO As Worksheet, ByVal templateRange As Range)

    Dim c As Long

    For c = templateRange.Column To templateRange.Column + templateRange.Columns.Count - 1
        wsO.Columns(c).ColumnWidth = wsT.Columns(c).ColumnWidth
    Next c

End Sub

Private Sub CopyTemplateRowHeights(ByVal wsT As Worksheet, ByVal wsO As Worksheet, ByVal templateRange As Range, ByVal pasteRow As Long)

    Dim r As Long

    For r = 0 To templateRange.Rows.Count - 1
        wsO.Rows(pasteRow + r).RowHeight = wsT.Rows(templateRange.Row + r).RowHeight
    Next r

End Sub

Private Sub ReplaceTemplateDateIfPossible(ByVal wsO As Worksheet, ByVal templateRange As Range, ByVal pasteRow As Long, ByVal currentDate As Date)

    Dim targetCell As Range
    Set targetCell = FindTemplateDateCell(wsO, templateRange, pasteRow)

    If targetCell Is Nothing Then
        Exit Sub
    End If

    If targetCell.MergeCells Then
        Set targetCell = targetCell.MergeArea.Cells(1, 1)
    End If

    targetCell.Value = currentDate
    targetCell.NumberFormatLocal = "yyyy/m/d (aaa)"
    wsO.Columns(targetCell.Column).ColumnWidth = Application.WorksheetFunction.Max(wsO.Columns(targetCell.Column).ColumnWidth, 14)

End Sub

Private Function FindTemplateDateCell(ByVal wsO As Worksheet, ByVal templateRange As Range, ByVal pasteRow As Long) As Range

    Dim firstSearchRow As Long
    Dim lastHeaderRow As Long

    firstSearchRow = pasteRow
    lastHeaderRow = Application.WorksheetFunction.Min(pasteRow + FIRST_TIME_ROW - 2, pasteRow + templateRange.Rows.Count - 1)

    Set FindTemplateDateCell = FindDateCandidateInRows(wsO, templateRange, firstSearchRow, lastHeaderRow)

    If FindTemplateDateCell Is Nothing Then
        Set FindTemplateDateCell = FindDateCandidateInRows(wsO, templateRange, pasteRow, pasteRow + templateRange.Rows.Count - 1)
    End If

End Function

Private Function FindDateCandidateInRows(ByVal wsO As Worksheet, ByVal templateRange As Range, ByVal firstRow As Long, ByVal lastRow As Long) As Range

    Dim r As Long
    Dim c As Long
    Dim cell As Range

    For r = firstRow To lastRow
        For c = templateRange.Column To templateRange.Column + templateRange.Columns.Count - 1
            If Not IsTemplateTimeAxisCell(r, c, firstRow) Then
                Set cell = wsO.Cells(r, c)
                If IsTemplateDateCandidate(cell) Then
                    Set FindDateCandidateInRows = cell
                    Exit Function
                End If
            End If
        Next c
    Next r

    Set FindDateCandidateInRows = Nothing

End Function

Private Function IsTemplateTimeAxisCell(ByVal absoluteRow As Long, ByVal absoluteCol As Long, ByVal blockFirstRow As Long) As Boolean

    IsTemplateTimeAxisCell = ((absoluteCol = NEW_TIME_COL Or absoluteCol = OLD_MINUTE_COL) And _
                              absoluteRow >= blockFirstRow + FIRST_TIME_ROW - 1 And _
                              absoluteRow <= blockFirstRow + LAST_TIME_ROW - 1)

End Function

Private Function IsTemplateDateCandidate(ByVal cell As Range) As Boolean

    Dim targetCell As Range

    If cell.MergeCells Then
        Set targetCell = cell.MergeArea.Cells(1, 1)
    Else
        Set targetCell = cell
    End If

    If Len(Trim$(CStr(targetCell.Value))) = 0 Then
        IsTemplateDateCandidate = False
    ElseIf IsDate(targetCell.Value) Then
        IsTemplateDateCandidate = True
    Else
        IsTemplateDateCandidate = ContainsDateMarker(CStr(targetCell.Value))
    End If

End Function

Private Function ContainsDateMarker(ByVal valueText As String) As Boolean

    ContainsDateMarker = (InStr(valueText, "/") > 0 Or _
                          InStr(valueText, "-") > 0 Or _
                          InStr(valueText, "20") > 0 Or _
                          (InStr(valueText, "年") > 0 And InStr(valueText, "月") > 0 And InStr(valueText, "日") > 0))

End Function

Private Sub CopyTemplatePrintSettings(ByVal wsT As Worksheet, ByVal wsO As Worksheet, ByVal templateRange As Range, ByVal lastRow As Long)

    With wsO.PageSetup
        .Orientation = wsT.PageSetup.Orientation
        .PaperSize = wsT.PageSetup.PaperSize
        .LeftMargin = wsT.PageSetup.LeftMargin
        .RightMargin = wsT.PageSetup.RightMargin
        .TopMargin = wsT.PageSetup.TopMargin
        .BottomMargin = wsT.PageSetup.BottomMargin
        .HeaderMargin = wsT.PageSetup.HeaderMargin
        .FooterMargin = wsT.PageSetup.FooterMargin
        .CenterHorizontally = wsT.PageSetup.CenterHorizontally
        .CenterVertically = wsT.PageSetup.CenterVertically
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintArea = wsO.Range( _
            wsO.Cells(1, templateRange.Column), _
            wsO.Cells(lastRow, templateRange.Column + templateRange.Columns.Count - 1) _
        ).Address
    End With

End Sub

Private Function GetSheetOrError(ByVal sheetName As String) As Worksheet

    On Error GoTo NotFound
    Set GetSheetOrError = ThisWorkbook.Worksheets(sheetName)
    Exit Function

NotFound:
    Err.Raise vbObjectError + 100, , "Sheet not found: " & sheetName

End Function

Public Sub CreateTemplateDraft()

    On Error GoTo ErrorHandler

    Dim wsT As Worksheet
    Set wsT = GetSheetOrError(SHEET_TEMPLATE)

    Dim wsExisting As Worksheet
    Set wsExisting = Nothing

    If SheetExists(SHEET_TEMPLATE_DRAFT) Then
        If MsgBox("TemplateDraft already exists. Delete and recreate it?", vbQuestion + vbYesNo) <> vbYes Then
            Exit Sub
        End If

        Set wsExisting = ThisWorkbook.Worksheets(SHEET_TEMPLATE_DRAFT)
    End If

    Dim previousScreenUpdating As Boolean
    Dim previousDisplayAlerts As Boolean
    Dim previousEnableEvents As Boolean
    Dim appStateCaptured As Boolean

    previousScreenUpdating = Application.ScreenUpdating
    previousDisplayAlerts = Application.DisplayAlerts
    previousEnableEvents = Application.EnableEvents
    appStateCaptured = True

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    If Not wsExisting Is Nothing Then
        wsExisting.Delete
    End If

    wsT.Copy After:=wsT

    Dim wsD As Worksheet
    Set wsD = ActiveSheet
    wsD.Name = SHEET_TEMPLATE_DRAFT

    ApplyTemplateDraftTimeAxis wsD

    wsD.Activate

    Application.EnableEvents = previousEnableEvents
    Application.DisplayAlerts = previousDisplayAlerts
    Application.ScreenUpdating = previousScreenUpdating

    MsgBox "TemplateDraft created.", vbInformation
    Exit Sub

ErrorHandler:
    If appStateCaptured Then
        Application.EnableEvents = previousEnableEvents
        Application.DisplayAlerts = previousDisplayAlerts
        Application.ScreenUpdating = previousScreenUpdating
    Else
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
    End If

    MsgBox "Error while creating TemplateDraft." & vbCrLf & _
           "Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbCritical

End Sub

Private Sub ApplyTemplateDraftTimeAxis(ByVal ws As Worksheet)

    Dim templateRange As Range
    Set templateRange = GetTemplateRange(ws)

    If templateRange Is Nothing Then
        Exit Sub
    End If

    Dim lastCol As Long
    lastCol = templateRange.Column + templateRange.Columns.Count - 1

    Dim lastTimeRowToUse As Long
    lastTimeRowToUse = Application.WorksheetFunction.Min(LAST_TIME_ROW, ws.Rows.Count)

    UnmergeTemplateDraftTimeColumns ws, NEW_TIME_COL, OLD_MINUTE_COL, FIRST_TIME_ROW, lastTimeRowToUse
    ClearTemplateDraftOldTimeValues ws, NEW_TIME_COL, OLD_MINUTE_COL, FIRST_TIME_ROW, lastTimeRowToUse

    Dim r As Long
    Dim slotTime As Date
    Dim minuteValue As Long

    r = FIRST_TIME_ROW
    slotTime = TimeSerial(DRAFT_TIME_START_HOUR, 0, 0)

    Do While slotTime <= TimeSerial(DRAFT_TIME_END_HOUR, 0, 0) And r <= lastTimeRowToUse

        minuteValue = Minute(slotTime)

        With ws.Cells(r, NEW_TIME_COL)
            .Value = Format(slotTime, "h:mm")
            .NumberFormat = "@"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = (minuteValue = 0)
        End With

        FormatTemplateDraftTimeRow ws.Range(ws.Cells(r, templateRange.Column), ws.Cells(r, lastCol)), _
                                   minuteValue, _
                                   slotTime >= TimeSerial(DRAFT_HATCH_START_HOUR, 0, 0)

        slotTime = DateAdd("n", 15, slotTime)
        r = r + 1

    Loop

    SetTemplateDraftTimeColumnLayout ws
    FixTemplateDraftDateDisplays ws, templateRange

End Sub

Private Sub UnmergeTemplateDraftTimeColumns(ByVal ws As Worksheet, ByVal timeCol As Long, ByVal oldMinuteCol As Long, ByVal firstRow As Long, ByVal lastRow As Long)

    Dim r As Long

    For r = firstRow To lastRow
        If ws.Cells(r, timeCol).MergeCells Then
            ws.Cells(r, timeCol).MergeArea.UnMerge
        End If

        If ws.Cells(r, oldMinuteCol).MergeCells Then
            ws.Cells(r, oldMinuteCol).MergeArea.UnMerge
        End If
    Next r

End Sub

Private Sub ClearTemplateDraftOldTimeValues(ByVal ws As Worksheet, ByVal timeCol As Long, ByVal oldMinuteCol As Long, ByVal firstRow As Long, ByVal lastRow As Long)

    With ws.Range(ws.Cells(firstRow, timeCol), ws.Cells(lastRow, timeCol))
        .ClearContents
        .NumberFormat = "@"
    End With

    With ws.Range(ws.Cells(firstRow, oldMinuteCol), ws.Cells(lastRow, oldMinuteCol))
        .ClearContents
        .Borders.LineStyle = xlNone
        .Interior.Pattern = xlNone
    End With

End Sub

Private Sub SetTemplateDraftTimeColumnLayout(ByVal ws As Worksheet)

    With ws.Columns(NEW_TIME_COL)
        .Hidden = False
        .ColumnWidth = 8.5
    End With

    With ws.Columns(OLD_MINUTE_COL)
        .Hidden = True
    End With

End Sub

Private Sub FixTemplateDraftDateDisplays(ByVal ws As Worksheet, ByVal templateRange As Range)

    Dim cell As Range

    For Each cell In templateRange.Cells
        If IsDate(cell.Value) Then
            cell.NumberFormatLocal = "yyyy/m/d"
            ws.Columns(cell.Column).ColumnWidth = Application.WorksheetFunction.Max(ws.Columns(cell.Column).ColumnWidth, 12)
        End If
    Next cell

End Sub

Private Sub FormatTemplateDraftTimeRow(ByVal rowRange As Range, ByVal minuteValue As Long, ByVal useHatch As Boolean)

    With rowRange.Borders(xlEdgeTop)
        If minuteValue = 0 Then
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(89, 89, 89)
        Else
            .LineStyle = xlDot
            .Weight = xlThin
            .Color = RGB(191, 191, 191)
        End If
    End With

    If useHatch Then
        With rowRange.Interior
            .Pattern = xlLightDown
            .PatternColor = RGB(217, 217, 217)
        End With
    End If

End Sub

Private Function SheetExists(ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    SheetExists = Not ws Is Nothing

End Function

Public Sub CheckAppointmentBook_Phase1()

    On Error GoTo ErrorHandler

    Dim wsT As Worksheet
    Set wsT = GetSheetOrError(SHEET_TEMPLATE)

    Dim tr As Range
    Set tr = GetTemplateRange(wsT)

    If tr Is Nothing Then
        MsgBox "Template sheet is empty.", vbExclamation
        Exit Sub
    End If

    MsgBox "Appointment book settings check" & vbCrLf & vbCrLf & _
           "Template range: " & tr.Address(False, False) & vbCrLf & _
           "Required sheets: Template / Settings / Output" & vbCrLf & _
           "Settings!B2 = Year" & vbCrLf & _
           "Settings!B3 = Month" & vbCrLf & vbCrLf & _
           "Run macro: GenerateAppointmentBook_Phase2", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error during settings check." & vbCrLf & Err.Description, vbCritical

End Sub
