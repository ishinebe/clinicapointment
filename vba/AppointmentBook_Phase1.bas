Option Explicit

'============================================================
' ClinicAppointment
' Module: AppointmentBook
' Version: 2026.07.07-Phase1
'
' Phase 1:
' - Copy Template sheet design master to Output sheet.
' - Preserve current appointment book layout as much as possible.
'
' Run:
' Alt + F8 -> GenerateAppointmentBook_Phase1
'
' Required sheets:
' - Template
' - Settings
' - Output
'
' Settings:
' - B2: Year
' - B3: Month
'
' Note:
' Phase 1 focuses on copying Template for one month.
' Time rendering, staff replacement, shading, and exceptions are later phases.
'============================================================

Private Const SHEET_TEMPLATE As String = "Template"
Private Const SHEET_TEMPLATE_DRAFT As String = "TemplateDraft"
Private Const SHEET_SETTINGS As String = "Settings"
Private Const SHEET_OUTPUT As String = "Output"

Private Const SETTINGS_YEAR_CELL As String = "B2"
Private Const SETTINGS_MONTH_CELL As String = "B3"

Private Const BLOCK_GAP_ROWS As Long = 2

Private Const DRAFT_TIME_START_HOUR As Long = 9
Private Const DRAFT_TIME_END_HOUR As Long = 19
Private Const DRAFT_HATCH_START_HOUR As Long = 18

Public Sub GenerateAppointmentBook_Phase1()

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
        ReplaceDateIfPossible wsO, templateRange, pasteRow, currentDate

        pasteRow = pasteRow + blockHeight + BLOCK_GAP_ROWS

    Next d

    CopyTemplatePrintSettings wsT, wsO, templateRange, pasteRow - BLOCK_GAP_ROWS - 1

    wsO.Activate

    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Phase 1 complete: Template copied for " & targetYear & "/" & targetMonth & ".", vbInformation
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Error occurred." & vbCrLf & _
           "Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbCritical

End Sub

Private Function GetTemplateRange(ByVal ws As Worksheet) As Range

    Dim ur As Range
    Set ur = ws.UsedRange

    If WorksheetFunction.CountA(ur) = 0 Then
        Set GetTemplateRange = Nothing
    Else
        Set GetTemplateRange = ur
    End If

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

Private Sub ReplaceDateIfPossible(ByVal wsO As Worksheet, ByVal templateRange As Range, ByVal pasteRow As Long, ByVal currentDate As Date)

    Dim targetCell As Range
    Set targetCell = FindDateLikeCell(wsO, templateRange, pasteRow)

    If targetCell Is Nothing Then
        ' If date cell is not found, do nothing in Phase 1.
        ' Phase 2 will use anchors or fixed cell mapping.
        Exit Sub
    End If

    targetCell.Value = Format(currentDate, "yyyy年m月d日") & "（" & GetJapaneseWeekday(currentDate) & "）"

End Sub

Private Function FindDateLikeCell(ByVal wsO As Worksheet, ByVal templateRange As Range, ByVal pasteRow As Long) As Range

    Dim r As Long
    Dim c As Long
    Dim txt As String

    For r = pasteRow To pasteRow + templateRange.Rows.Count - 1
        For c = templateRange.Column To templateRange.Column + templateRange.Columns.Count - 1

            txt = CStr(wsO.Cells(r, c).Value)

            If InStr(txt, "年") > 0 And InStr(txt, "月") > 0 And InStr(txt, "日") > 0 Then
                Set FindDateLikeCell = wsO.Cells(r, c)
                Exit Function
            End If

        Next c
    Next r

    Set FindDateLikeCell = Nothing

End Function

Private Function GetJapaneseWeekday(ByVal d As Date) As String

    Select Case Weekday(d, vbSunday)
        Case 1: GetJapaneseWeekday = "日"
        Case 2: GetJapaneseWeekday = "月"
        Case 3: GetJapaneseWeekday = "火"
        Case 4: GetJapaneseWeekday = "水"
        Case 5: GetJapaneseWeekday = "木"
        Case 6: GetJapaneseWeekday = "金"
        Case 7: GetJapaneseWeekday = "土"
    End Select

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

    Dim hourCol As Long
    Dim minuteCol As Long

    If Not FindDraftTimeColumns(ws, templateRange, hourCol, minuteCol) Then
        Err.Raise vbObjectError + 200, , "Time columns were not found in Template."
    End If

    Dim lastCol As Long
    lastCol = templateRange.Column + templateRange.Columns.Count - 1

    Dim combinedTimeColumnWidth As Double
    combinedTimeColumnWidth = ws.Columns(hourCol).ColumnWidth + ws.Columns(minuteCol).ColumnWidth

    Dim r As Long
    Dim hourValue As Variant
    Dim minuteValue As Variant
    Dim slotTime As Date
    Dim slots As Collection

    Set slots = New Collection

    For r = templateRange.Row To templateRange.Row + templateRange.Rows.Count - 1

        hourValue = DraftCellNumber(ws.Cells(r, hourCol))
        minuteValue = DraftCellNumber(ws.Cells(r, minuteCol))

        If IsDraftTimeSlot(hourValue, minuteValue) Then
            slotTime = TimeSerial(CLng(hourValue), CLng(minuteValue), 0)

            If slotTime >= TimeSerial(DRAFT_TIME_START_HOUR, 0, 0) And _
               slotTime <= TimeSerial(DRAFT_TIME_END_HOUR, 0, 0) Then
                slots.Add Array(r, CLng(hourValue), CLng(minuteValue))
            End If
        End If

    Next r

    Dim slot As Variant

    UnmergeTemplateDraftTimeColumns ws, templateRange, hourCol, minuteCol

    For Each slot In slots

        r = CLng(slot(0))
        slotTime = TimeSerial(CLng(slot(1)), CLng(slot(2)), 0)

        With ws.Cells(r, hourCol)
            .Value = Format(slotTime, "h:mm")
            .NumberFormat = "@"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = (CLng(slot(2)) = 0)
        End With

        With ws.Cells(r, minuteCol)
            .ClearContents
            .Borders.LineStyle = xlNone
            .Interior.Pattern = xlNone
        End With

        FormatTemplateDraftTimeRow ws.Range(ws.Cells(r, templateRange.Column), ws.Cells(r, lastCol)), _
                                   CLng(slot(2)), _
                                   slotTime >= TimeSerial(DRAFT_HATCH_START_HOUR, 0, 0)

    Next slot

    SetTemplateDraftTimeColumnLayout ws, hourCol, minuteCol, combinedTimeColumnWidth

End Sub

Private Sub UnmergeTemplateDraftTimeColumns(ByVal ws As Worksheet, ByVal templateRange As Range, ByVal hourCol As Long, ByVal minuteCol As Long)

    Dim r As Long

    For r = templateRange.Row To templateRange.Row + templateRange.Rows.Count - 1
        If ws.Cells(r, hourCol).MergeCells Then
            ws.Cells(r, hourCol).MergeArea.UnMerge
        End If

        If ws.Cells(r, minuteCol).MergeCells Then
            ws.Cells(r, minuteCol).MergeArea.UnMerge
        End If
    Next r

End Sub

Private Sub SetTemplateDraftTimeColumnLayout(ByVal ws As Worksheet, ByVal hourCol As Long, ByVal minuteCol As Long, ByVal combinedTimeColumnWidth As Double)

    With ws.Columns(hourCol)
        .Hidden = False
        .ColumnWidth = Application.WorksheetFunction.Max(8, combinedTimeColumnWidth)
    End With

    With ws.Columns(minuteCol)
        .ClearContents
        .ColumnWidth = 0
        .Hidden = True
    End With

End Sub

Private Function FindDraftTimeColumns(ByVal ws As Worksheet, ByVal templateRange As Range, ByRef hourCol As Long, ByRef minuteCol As Long) As Boolean

    Dim bestScore As Long
    Dim bestHourCol As Long
    Dim bestMinuteCol As Long
    Dim c As Long
    Dim r As Long
    Dim score As Long
    Dim hourValue As Variant
    Dim minuteValue As Variant

    For c = templateRange.Column To templateRange.Column + templateRange.Columns.Count - 2
        score = 0

        For r = templateRange.Row To templateRange.Row + templateRange.Rows.Count - 1
            hourValue = DraftCellNumber(ws.Cells(r, c))
            minuteValue = DraftCellNumber(ws.Cells(r, c + 1))

            If IsDraftTimeSlot(hourValue, minuteValue) Then
                score = score + 1
            End If
        Next r

        If score > bestScore Then
            bestScore = score
            bestHourCol = c
            bestMinuteCol = c + 1
        End If
    Next c

    If bestScore >= 4 Then
        hourCol = bestHourCol
        minuteCol = bestMinuteCol
        FindDraftTimeColumns = True
    Else
        FindDraftTimeColumns = False
    End If

End Function

Private Function DraftCellNumber(ByVal cell As Range) As Variant

    Dim sourceCell As Range

    If cell.MergeCells Then
        Set sourceCell = cell.MergeArea.Cells(1, 1)
    Else
        Set sourceCell = cell
    End If

    Dim txt As String
    txt = Trim$(CStr(sourceCell.Value))

    If Len(txt) = 0 Then
        DraftCellNumber = Empty
    ElseIf IsNumeric(txt) Then
        DraftCellNumber = CLng(Val(txt))
    ElseIf IsDate(txt) Then
        DraftCellNumber = Hour(CDate(txt))
    Else
        DraftCellNumber = Empty
    End If

End Function

Private Function IsDraftTimeSlot(ByVal hourValue As Variant, ByVal minuteValue As Variant) As Boolean

    If IsEmpty(hourValue) Or IsEmpty(minuteValue) Then
        IsDraftTimeSlot = False
        Exit Function
    End If

    If CLng(hourValue) < 0 Or CLng(hourValue) > 23 Then
        IsDraftTimeSlot = False
        Exit Function
    End If

    Select Case CLng(minuteValue)
        Case 0, 15, 30, 45
            IsDraftTimeSlot = True
        Case Else
            IsDraftTimeSlot = False
    End Select

End Function

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

    MsgBox "Phase 1 settings check" & vbCrLf & vbCrLf & _
           "Template range: " & tr.Address(False, False) & vbCrLf & _
           "Required sheets: Template / Settings / Output" & vbCrLf & _
           "Settings!B2 = Year" & vbCrLf & _
           "Settings!B3 = Month" & vbCrLf & vbCrLf & _
           "Run macro: GenerateAppointmentBook_Phase1", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error during settings check." & vbCrLf & Err.Description, vbCritical

End Sub
