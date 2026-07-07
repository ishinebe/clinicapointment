Option Explicit

'============================================================
' ClinicAppointment
' Module: AppointmentBook
' Version: 2026.07.07-Phase5D-one-cell-work-pattern
'
' Important:
' - One-day template range is fixed to Template!A1:J46.
' - Template is the finalized design master.
' - Time axis is already designed in Template, so this macro does not redraw it.
' - Settings!B7:F13 is one-cell complete:
'   blank = normal, 休 = off, 午前のみ / xx:xxまで = short-time work.
' - Settings!B16:F16 is clinic-wide monthly default close time.
'============================================================

Private Const SHEET_TEMPLATE As String = "Template"
Private Const SHEET_TEMPLATE_DRAFT As String = "TemplateDraft"
Private Const SHEET_SETTINGS As String = "Settings"
Private Const SHEET_OUTPUT As String = "Output"

Private Const SETTINGS_YEAR_CELL As String = "B2"
Private Const SETTINGS_MONTH_CELL As String = "B3"
Private Const SETTINGS_STAFF_FIRST_CELL As String = "B5"
Private Const SETTINGS_WORK_PATTERN_FIRST_CELL As String = "B7"
Private Const SETTINGS_STAFF_HEADER_RANGE As String = "B5:F5"
Private Const SETTINGS_WORK_PATTERN_RANGE As String = "B7:F13"
Private Const SETTINGS_STAFF_MASTER_RANGE As String = "H5:H24"
Private Const SETTINGS_MONTHLY_CLOSE_CELL As String = "B16"
Private Const SETTINGS_MONTHLY_CLOSE_RANGE As String = "B16:F16"

Private Const TEMPLATE_ONE_DAY_RANGE As String = "A1:J46"
Private Const BLOCK_GAP_ROWS As Long = 2

Private Const NEW_TIME_COL As Long = 1
Private Const OLD_MINUTE_COL As Long = 8
Private Const FIRST_TIME_ROW As Long = 7
Private Const LAST_TIME_ROW As Long = 46
Private Const HEADER_ROW_IN_TEMPLATE As Long = 4

Private Const STAFF_SLOT_COUNT As Long = 5
Private Const STAFF_COL_1 As Long = 2    ' B: Dr 1
Private Const STAFF_COL_2 As Long = 4    ' D: Dr 2
Private Const STAFF_COL_3 As Long = 6    ' F: Reserve
Private Const STAFF_COL_4 As Long = 9    ' I: DH 1. H is intentionally skipped.
Private Const STAFF_COL_5 As Long = 10   ' J: DH 2

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

Public Sub GenerateAppointmentBook_Phase5()
    GenerateAppointmentBookCore "Phase 5"
End Sub

Public Sub SetupSettingsDropdowns()

    On Error GoTo ErrorHandler

    Dim wsS As Worksheet
    Set wsS = GetSheetOrError(SHEET_SETTINGS)

    Application.ScreenUpdating = False

    ApplySettingsLabels wsS
    InitializeStaffMasterFromCurrentHeaders wsS
    ApplyStaffHeaderDropdowns wsS
    ApplyWorkPatternDropdowns wsS
    ApplyMonthlyCloseDropdown wsS

    Application.ScreenUpdating = True

    MsgBox "Settings dropdowns are ready." & vbCrLf & _
           "Staff headers: Settings!B5:F5" & vbCrLf & _
           "Work pattern: Settings!B7:F13" & vbCrLf & _
           "Clinic monthly close: Settings!B16:F16" & vbCrLf & _
           "Blank = normal. Use one dropdown cell for off or short-time.", vbInformation
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error while setting up Settings dropdowns." & vbCrLf & _
           "Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbCritical

End Sub

Private Sub ApplySettingsLabels(ByVal wsS As Worksheet)

    wsS.Range("A7").Value = "月"
    wsS.Range("A8").Value = "火"
    wsS.Range("A9").Value = "水"
    wsS.Range("A10").Value = "木"
    wsS.Range("A11").Value = "金"
    wsS.Range("A12").Value = "土"
    wsS.Range("A13").Value = "日"

    wsS.Range("A16").Value = "医院全体の当月原則終了時刻"
    PrepareMonthlyCloseVisualRange wsS

    wsS.Range("H4").Value = "担当者マスター"

End Sub

Private Sub PrepareMonthlyCloseVisualRange(ByVal wsS As Worksheet)

    Dim keepValue As Variant
    keepValue = wsS.Range(SETTINGS_MONTHLY_CLOSE_CELL).Value

    With wsS.Range(SETTINGS_MONTHLY_CLOSE_RANGE)
        .UnMerge
        .ClearContents
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    wsS.Range(SETTINGS_MONTHLY_CLOSE_CELL).Value = keepValue

End Sub

Private Sub InitializeStaffMasterFromCurrentHeaders(ByVal wsS As Worksheet)

    If WorksheetFunction.CountA(wsS.Range(SETTINGS_STAFF_MASTER_RANGE)) > 0 Then
        Exit Sub
    End If

    Dim sourceCell As Range
    Dim nextRow As Long
    Dim candidate As String

    nextRow = wsS.Range(SETTINGS_STAFF_MASTER_RANGE).Row

    For Each sourceCell In wsS.Range(SETTINGS_STAFF_HEADER_RANGE).Cells
        candidate = Trim$(CStr(sourceCell.Value))
        If Len(candidate) > 0 Then
            If Not ValueExistsInRange(wsS.Range(SETTINGS_STAFF_MASTER_RANGE), candidate) Then
                wsS.Cells(nextRow, wsS.Range(SETTINGS_STAFF_MASTER_RANGE).Column).Value = candidate
                nextRow = nextRow + 1
            End If
        End If
    Next sourceCell

End Sub

Private Function ValueExistsInRange(ByVal targetRange As Range, ByVal valueText As String) As Boolean

    Dim cell As Range

    For Each cell In targetRange.Cells
        If Trim$(CStr(cell.Value)) = valueText Then
            ValueExistsInRange = True
            Exit Function
        End If
    Next cell

    ValueExistsInRange = False

End Function

Private Sub ApplyStaffHeaderDropdowns(ByVal wsS As Worksheet)

    Dim staffMaster As Range
    Set staffMaster = wsS.Range(SETTINGS_STAFF_MASTER_RANGE)

    With wsS.Range(SETTINGS_STAFF_HEADER_RANGE).Validation
        .Delete
    End With

    If WorksheetFunction.CountA(staffMaster) = 0 Then
        Exit Sub
    End If

    With wsS.Range(SETTINGS_STAFF_HEADER_RANGE).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="=" & staffMaster.Address(True, True, xlA1, True)
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "担当者選択"
        .InputMessage = "担当者マスターから選択してください。"
        .ErrorTitle = "入力できません"
        .ErrorMessage = "担当者マスターに登録されている名前から選択してください。"
    End With

End Sub

Private Sub ApplyWorkPatternDropdowns(ByVal wsS As Worksheet)

    With wsS.Range(SETTINGS_WORK_PATTERN_RANGE).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="休,午前のみ,15:00まで,15:30まで,16:00まで,16:30まで,17:00まで,17:30まで,18:00まで,18:30まで"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "勤務パターン"
        .InputMessage = "通常は空欄。休みは休、午前勤務は午前のみ、早上がりは終了時刻を選択してください。"
        .ErrorTitle = "入力できません"
        .ErrorMessage = "プルダウンから選択するか、空欄のままにしてください。"
    End With

End Sub

Private Sub ApplyMonthlyCloseDropdown(ByVal wsS As Worksheet)

    With wsS.Range(SETTINGS_MONTHLY_CLOSE_CELL).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="15:00,15:30,16:00,16:30,17:00,17:30,18:00,18:30,19:00"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "医院全体の当月原則終了時刻"
        .InputMessage = "通常は空欄。その月全体で医院の終了時刻を変更する場合だけ選択してください。"
        .ErrorTitle = "入力できません"
        .ErrorMessage = "プルダウンから終了時刻を選択するか、空欄のままにしてください。"
    End With

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
    Set GetTemplateRange = ws.Range(TEMPLATE_ONE_DAY_RANGE)
End Function

Private Sub ClearOutput(ByVal ws As Worksheet)
    ws.Cells.Clear
    On Error Resume Next
    ws.ResetAllPageBreaks
    On Error GoTo 0
End Sub

Private Sub CopyTemplateBlock(ByVal wsT As Worksheet, ByVal wsO As Worksheet, ByVal templateRange As Range, ByVal pasteRow As Long)

    templateRange.Copy
    wsO.Cells(pasteRow, templateRange.Column).PasteSpecial Paste:=xlPasteAll
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
    ReplaceStaffHeadersIfConfigured wsO, wsS, pasteRow
    ApplyStaffWorkPatternIfConfigured wsO, wsS, pasteRow, currentDate
    ApplyClinicHoursInBlock wsO, wsS, templateRange, pasteRow, currentDate

End Sub

Private Sub ReplaceTemplateDateIfPossible(ByVal wsO As Worksheet, ByVal templateRange As Range, ByVal pasteRow As Long, ByVal currentDate As Date)

    Dim targetCell As Range
    Set targetCell = FindTemplateDateCell(wsO, templateRange, pasteRow)

    If targetCell Is Nothing Then Exit Sub
    If targetCell.MergeCells Then Set targetCell = targetCell.MergeArea.Cells(1, 1)

    targetCell.Value = currentDate
    targetCell.NumberFormatLocal = "yyyy/m/d (aaa)"
    wsO.Columns(targetCell.Column).ColumnWidth = Application.WorksheetFunction.Max(wsO.Columns(targetCell.Column).ColumnWidth, 14)

End Sub

Private Sub ReplaceStaffHeadersIfConfigured(ByVal wsO As Worksheet, ByVal wsS As Worksheet, ByVal pasteRow As Long)

    Dim staffNames(1 To STAFF_SLOT_COUNT) As String
    Dim targetCols(1 To STAFF_SLOT_COUNT) As Long
    Dim i As Long
    Dim hasAnyStaffSetting As Boolean

    GetStaffTargetColumns targetCols

    For i = 1 To STAFF_SLOT_COUNT
        staffNames(i) = Trim$(CStr(wsS.Range(SETTINGS_STAFF_FIRST_CELL).Offset(0, i - 1).Value))
        If Len(staffNames(i)) > 0 Then hasAnyStaffSetting = True
    Next i

    If Not hasAnyStaffSetting Then Exit Sub

    Dim headerRow As Long
    headerRow = pasteRow + HEADER_ROW_IN_TEMPLATE - 1

    Dim targetCell As Range

    For i = 1 To STAFF_SLOT_COUNT
        If Len(staffNames(i)) > 0 Then
            Set targetCell = wsO.Cells(headerRow, targetCols(i))
            If targetCell.MergeCells Then Set targetCell = targetCell.MergeArea.Cells(1, 1)
            targetCell.Value = staffNames(i)
        End If
    Next i

End Sub

Private Sub ApplyStaffWorkPatternIfConfigured(ByVal wsO As Worksheet, ByVal wsS As Worksheet, ByVal pasteRow As Long, ByVal currentDate As Date)

    If WorksheetFunction.CountA(wsS.Range(SETTINGS_WORK_PATTERN_RANGE)) = 0 Then Exit Sub

    Dim weekdayIndexMondayFirst As Long
    weekdayIndexMondayFirst = Weekday(currentDate, vbMonday)

    Dim targetCols(1 To STAFF_SLOT_COUNT) As Long
    GetStaffTargetColumns targetCols

    Dim i As Long
    Dim statusValue As String
    Dim closeTime As Date

    For i = 1 To STAFF_SLOT_COUNT
        statusValue = Trim$(CStr(wsS.Range(SETTINGS_WORK_PATTERN_FIRST_CELL).Offset(weekdayIndexMondayFirst - 1, i - 1).Value))

        If IsStaffOffValue(statusValue) Then
            ShadeStaffSlot wsO, pasteRow, targetCols(i)
            MarkStaffHeaderOff wsO, pasteRow, targetCols(i)
        ElseIf TryGetCloseTimeFromWorkPattern(statusValue, closeTime) Then
            ShadeStaffSlotFromTime wsO, pasteRow, targetCols(i), closeTime
            MarkStaffHeaderUntil wsO, pasteRow, targetCols(i), Format$(closeTime, "h:mm")
        End If
    Next i

End Sub

Private Function IsStaffOffValue(ByVal valueText As String) As Boolean

    Dim normalized As String
    normalized = LCase$(Trim$(valueText))

    IsStaffOffValue = (normalized = "休" Or normalized = "休み" Or normalized = "休診" Or _
                       normalized = "off" Or normalized = "x" Or normalized = "×" Or _
                       normalized = "0" Or normalized = "-" Or normalized = "欠")

End Function

Private Function TryGetCloseTimeFromWorkPattern(ByVal valueText As String, ByRef closeTime As Date) As Boolean

    Dim normalized As String
    normalized = Trim$(valueText)

    If Len(normalized) = 0 Then
        TryGetCloseTimeFromWorkPattern = False
        Exit Function
    End If

    If normalized = "午前のみ" Then
        closeTime = TimeSerial(12, 0, 0)
        TryGetCloseTimeFromWorkPattern = True
        Exit Function
    End If

    normalized = Replace(normalized, "まで", "")
    normalized = Replace(normalized, "迄", "")

    TryGetCloseTimeFromWorkPattern = TryParseTimeText(normalized, closeTime)

End Function

Private Sub ShadeStaffSlot(ByVal ws As Worksheet, ByVal pasteRow As Long, ByVal targetCol As Long)

    Dim firstCol As Long
    Dim lastCol As Long
    GetStaffSlotColumnSpan ws, pasteRow, targetCol, firstCol, lastCol

    ShadeRows ws, pasteRow + FIRST_TIME_ROW - 1, pasteRow + LAST_TIME_ROW - 1, firstCol, lastCol

End Sub

Private Sub ShadeStaffSlotFromTime(ByVal ws As Worksheet, ByVal pasteRow As Long, ByVal targetCol As Long, ByVal closeTime As Date)

    Dim firstCol As Long
    Dim lastCol As Long
    GetStaffSlotColumnSpan ws, pasteRow, targetCol, firstCol, lastCol

    Dim r As Long
    Dim slotTime As Date

    For r = pasteRow + FIRST_TIME_ROW - 1 To pasteRow + LAST_TIME_ROW - 1
        If TryGetTimeFromCell(ws.Cells(r, NEW_TIME_COL), slotTime) Then
            If slotTime >= closeTime Then
                ShadeRows ws, r, r, firstCol, lastCol
            End If
        End If
    Next r

End Sub

Private Sub GetStaffSlotColumnSpan(ByVal ws As Worksheet, ByVal pasteRow As Long, ByVal targetCol As Long, ByRef firstCol As Long, ByRef lastCol As Long)

    Dim headerCell As Range
    Set headerCell = ws.Cells(pasteRow + HEADER_ROW_IN_TEMPLATE - 1, targetCol)

    If headerCell.MergeCells Then
        firstCol = headerCell.MergeArea.Column
        lastCol = headerCell.MergeArea.Column + headerCell.MergeArea.Columns.Count - 1
    Else
        firstCol = targetCol
        lastCol = targetCol
    End If

End Sub

Private Sub MarkStaffHeaderOff(ByVal ws As Worksheet, ByVal pasteRow As Long, ByVal targetCol As Long)

    Dim headerCell As Range
    Set headerCell = ws.Cells(pasteRow + HEADER_ROW_IN_TEMPLATE - 1, targetCol)
    If headerCell.MergeCells Then Set headerCell = headerCell.MergeArea.Cells(1, 1)

    If InStr(CStr(headerCell.Value), "休") = 0 Then
        headerCell.Value = CStr(headerCell.Value) & " 休"
    End If

End Sub

Private Sub MarkStaffHeaderUntil(ByVal ws As Worksheet, ByVal pasteRow As Long, ByVal targetCol As Long, ByVal timeText As String)

    Dim headerCell As Range
    Set headerCell = ws.Cells(pasteRow + HEADER_ROW_IN_TEMPLATE - 1, targetCol)
    If headerCell.MergeCells Then Set headerCell = headerCell.MergeArea.Cells(1, 1)

    If InStr(CStr(headerCell.Value), "まで") = 0 Then
        headerCell.Value = CStr(headerCell.Value) & " " & timeText & "まで"
    End If

End Sub

Private Sub GetStaffTargetColumns(ByRef targetCols() As Long)

    targetCols(1) = STAFF_COL_1
    targetCols(2) = STAFF_COL_2
    targetCols(3) = STAFF_COL_3
    targetCols(4) = STAFF_COL_4
    targetCols(5) = STAFF_COL_5

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

Private Sub ApplyClinicHoursInBlock(ByVal wsO As Worksheet, ByVal wsS As Worksheet, ByVal templateRange As Range, ByVal pasteRow As Long, ByVal currentDate As Date)

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
    closeTime = GetEffectiveClinicCloseTime(wsS, weekdayValue)

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

Private Function GetEffectiveClinicCloseTime(ByVal wsS As Worksheet, ByVal weekdayValue As Long) As Date

    Dim monthlyValue As String
    Dim monthlyTime As Date

    monthlyValue = Trim$(CStr(wsS.Range(SETTINGS_MONTHLY_CLOSE_CELL).Value))

    If Len(monthlyValue) > 0 Then
        If TryParseTimeText(monthlyValue, monthlyTime) Then
            GetEffectiveClinicCloseTime = monthlyTime
            Exit Function
        End If
    End If

    GetEffectiveClinicCloseTime = GetClinicCloseTime(weekdayValue)

End Function

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
    ElseIf TryParseTimeText(valueText, parsedTime) Then
        TryGetTimeFromCell = True
    Else
        TryGetTimeFromCell = False
    End If

End Function

Private Function TryParseTimeText(ByVal valueText As String, ByRef parsedTime As Date) As Boolean

    If IsDate(valueText) Then
        parsedTime = TimeValue(valueText)
        TryParseTimeText = True
    Else
        TryParseTimeText = False
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

    If Not wsExisting Is Nothing Then wsExisting.Delete

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

    MsgBox "Appointment book settings check" & vbCrLf & vbCrLf & _
           "Template range: " & tr.Address(False, False) & vbCrLf & _
           "Required sheets: Template / Settings / Output" & vbCrLf & _
           "Settings!B2 = Year" & vbCrLf & _
           "Settings!B3 = Month" & vbCrLf & _
           "Settings!B5:F5 = Staff headers" & vbCrLf & _
           "Settings!B7:F13 = Work pattern" & vbCrLf & _
           "Settings!B16:F16 = Clinic monthly close time" & vbCrLf & _
           "Blank in B7:F13 means normal work." & vbCrLf & _
           "Options include 休, 午前のみ, and xx:xxまで." & vbCrLf & _
           "Staff mapping: B5->B, C5->D, D5->F, E5->I, F5->J" & vbCrLf & vbCrLf & _
           "Run setup macro first: SetupSettingsDropdowns" & vbCrLf & _
           "Run generation macro: GenerateAppointmentBook_Phase5", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error during settings check." & vbCrLf & Err.Description, vbCritical

End Sub
