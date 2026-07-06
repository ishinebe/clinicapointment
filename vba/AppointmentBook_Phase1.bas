Attribute VB_Name = "AppointmentBook"
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
Private Const SHEET_SETTINGS As String = "Settings"
Private Const SHEET_OUTPUT As String = "Output"

Private Const SETTINGS_YEAR_CELL As String = "B2"
Private Const SETTINGS_MONTH_CELL As String = "B3"

Private Const BLOCK_GAP_ROWS As Long = 2

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
