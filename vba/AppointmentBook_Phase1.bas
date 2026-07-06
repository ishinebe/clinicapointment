Option Explicit

'============================================================
' ClinicAppointment
' Module: AppointmentBook
' Version: 2026.07.07-Phase4AB-staff-headers
'
' Important:
' - One-day template range is fixed to Template!A1:J46.
' - Do not use UsedRange for Template because stray formatting can expand
'   the copied area and create hundreds of printed pages.
' - Template is the finalized design master.
' - Time axis is already designed in Template, so this macro does not redraw it.
' - Clinic-hour shading is based on the actual time labels copied from Template.
' - Staff headers can be overridden from Settings!B5:F5.
'============================================================

Private Const SHEET_TEMPLATE As String = "Template"
Private Const SHEET_TEMPLATE_DRAFT As String = "TemplateDraft"
Private Const SHEET_SETTINGS As String = "Settings"
Private Const SHEET_OUTPUT As String = "Output"

Private Const SETTINGS_YEAR_CELL As String = "B2"
Private Const SETTINGS_MONTH_CELL As String = "B3"
Private Const SETTINGS_STAFF_FIRST_CELL As String = "B5"

Private Const TEMPLATE_ONE_DAY_RANGE As String = "A1:J46"
Private Const BLOCK_GAP_ROWS As Long = 2

Private Const NEW_TIME_COL As Long = 1
Private Const OLD_MINUTE_COL As Long = 8
Private Const FIRST_TIME_ROW As Long = 7
Private Const LAST_TIME_ROW As Long = 46
Private Const HEADER_ROW_IN_TEMPLATE As Long = 4

Private Const STAFF_SLOT_COUNT As Long = 5
Private Const STAFF_COL_1 As Long = 2   ' B
Private Const STAFF_COL_2 As Long = 4   ' D
Private Const STAFF_COL_3 As Long = 6   ' F
Private Const STAFF_COL_4 As Long = 8   ' H
Private Const STAFF_COL_5 As Long = 9   ' I

Public Sub GenerateAppointmentBook_Phase1()
    GenerateAppointmentBookCore "Phase 1"
End Sub

Public Sub GenerateAppointmentBook_Phase2()
    GenerateAppointmentBookCore "Phase 2"
End Sub

Public Sub GenerateAppointmentBook_Phase3()
    GenerateAppointmentBookCore "Phase 3"
End Sub

Public Sub GenerateAppointmentBook_Phase4()
    GenerateAppointmentBookCore "Phase 4"
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
        ApplyOperationalInfo wsO, wsS, templateRange, pasteRow, currentDate

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

Private Sub ApplyOperationalInfo(ByVal wsO As Worksheet, ByVal wsS As Worksheet, ByVal templateRange As Range, ByVal pasteRow As Long, ByVal currentDate As Date)

    ReplaceTemplateDateIfPossible wsO, templateRange, pasteRow, currentDate
    ReplaceStaffHeadersIfConfigured wsO, wsS, templateRange, pasteRow
    ApplyClinicHoursInBlock wsO, templateRange, pasteRow, currentDate

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

Private Sub ReplaceStaffHeadersIfConfigured(ByVal wsO As Worksheet, ByVal wsS As Worksheet, ByVal templateRange As Range, ByVal pasteRow As Long)

    Dim staffNames(1 To STAFF_SLOT_COUNT) As String
    Dim i As Long
    Dim hasAnyStaffSetting As Boolean

    For i = 1 To STAFF_SLOT_COUNT
        staffNames(i) = Trim$(CStr(wsS.Range(SETTINGS_STAFF_FIRST_CELL).Offset(0, i - 1).Value))
        If Len(staffNames(i)) > 0 Then
            hasAnyStaffSetting = True
        End If
    Next i

    ' If Settings!B5:F5 is blank, keep the finalized Template headers as-is.
    If Not hasAnyStaffSetting Then
        Exit Sub
    End If

    Dim targetCols(1 To STAFF_SLOT_COUNT) As Long
    targetCols(1) = STAFF_COL_1
    targetCols(2) = STAFF_COL_2
    targetCols(3) = STAFF_COL_3
    targetCols(4) = STAFF_COL_4
    targetCols(5) = STAFF_COL_5

    Dim headerRow As Long
    headerRow = pasteRow + HEADER_ROW_IN_TEMPLATE - 1

    Dim targetCell As Range

    For i = 1 To STAFF_SLOT_COUNT
        If Len(staffNames(i)) > 0 Then
            Set targetCell = wsO.Cells(headerRow, targetCols(i))
            If targetCell.MergeCells Then
                Set targetCell = targetCell.MergeArea.Cells(1, 1)
            End If
            targetCell.Value = staffNames(i)
        End If
    Next i

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

Private Sub ApplyClinicHoursInBlock(ByVal wsO As Worksheet, ByVal templateRange As Range, ByVal pasteRow As Long, ByVal currentDate As Date)

    Dim startRow As Long
    Dim endRow As Long
    Dim lastCol As Long

    startRow = pasteRow + FIRST_TIME_ROW - 1
    endRow = pasteRow + LAST_TIME_ROW - 1
    lastCol = templateRange.Column + templateRange.Columns.Count - 1

    Dim weekdayValue As Long
    weekdayValue = Weekday(currentDate, vbSunday)

    If IsClosedWeekday(weekdayValue) Then
        ShadeRows wsO, startRow, endRow, templateRange.Column, lastCol
        PutClosedLabel wsO, pasteRow, templateRange
        Exit Sub
    End If

    Dim closeTime As Date
    closeTime = GetClinicCloseTime(weekdayValue)

    Dim r As Long
    Dim slotTime As Date

    For r = startRow To endRow
        If TryGetTimeFromCell(wsO.Cells(r, NEW_TIME_COL), slotTime) Then
            If slotTime >= closeTime Then
                ShadeRows wsO, r, r, templateRange.Column, lastCol
            End If
        End If
    Next r

End Sub

Private Function TryGetTimeFromCell(ByVal cell As Range, ByRef parsedTime As Date) As Boolean

    Dim valueText As String
    valueText = Trim$(CStr(cell.Value))

    If Len(valueText) = 0 Then
        TryGetTimeFromCell = False
        Exit Function
    End If

    If IsDate(cell.Value) Then
        parsedTime = TimeValue(cell.Value)
        TryGetTimeFromCell = True
    ElseIf IsDate(valueText) Then
        parsedTime = TimeValue(valueText)
        TryGetTimeFromCell = True
    Else
        TryGetTimeFromCell = False
    End If

End Function

Private Function IsClosedWeekday(ByVal weekdayValue As Long) As Boolean

    IsClosedWeekday = (weekdayValue = vbThursday Or weekdayValue = vbSunday)

End Function

Private Function GetClinicCloseTime(ByVal weekdayValue As Long) As Date

    Select Case weekdayValue
        Case vbWednesday
            GetClinicCloseTime = TimeSerial(18, 0, 0)
        Case vbSaturday
            GetClinicCloseTime = TimeSerial(17, 30, 0)
        Case Else
            GetClinicCloseTime = TimeSerial(19, 0, 0)
    End Select

End Function

Private Sub ShadeRows(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstCol As Long, ByVal lastCol As Long)

    With ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastRow, lastCol)).Interior
        .Pattern = xlLightDown
        .PatternColor = RGB(217, 217, 217)
    End With

End Sub

Private Sub PutClosedLabel(ByVal ws As Worksheet, ByVal pasteRow As Long, ByVal templateRange As Range)

    Dim labelRow As Long
    Dim labelCol As Long

    labelRow = pasteRow + FIRST_TIME_ROW - 1
    labelCol = templateRange.Column + 1

    With ws.Cells(labelRow, labelCol)
        .Value = "休診"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
    End With

End Sub

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
    wsD.Activate

    Application.EnableEvents = previousEnableEvents
    Application.DisplayAlerts = previousDisplayAlerts
    Application.ScreenUpdating = previousScreenUpdating

    MsgBox "TemplateDraft created. Time axis is preserved from Template.", vbInformation
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
           "Settings!B3 = Month" & vbCrLf & _
           "Settings!B5:F5 = Staff headers" & vbCrLf & vbCrLf & _
           "Run macro: GenerateAppointmentBook_Phase4", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error during settings check." & vbCrLf & Err.Description, vbCritical

End Sub
