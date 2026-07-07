Option Explicit

Private Const SHEET_OUTPUT As String = "Output"
Private Const OUTPUT_FIRST_COL As Long = 1
Private Const OUTPUT_LAST_COL As Long = 10
Private Const TEMPLATE_ROW_COUNT As Long = 46
Private Const BLOCK_GAP_ROWS As Long = 2
Private Const ONE_DAY_BLOCK_STEP As Long = TEMPLATE_ROW_COUNT + BLOCK_GAP_ROWS

Public Sub ApplyOneDayOnePagePrintSettings()
    On Error GoTo ErrorHandler

    Dim wsO As Worksheet
    Set wsO = ThisWorkbook.Worksheets(SHEET_OUTPUT)

    Dim lastRow As Long
    lastRow = GetLastOutputRow(wsO)

    If lastRow < TEMPLATE_ROW_COUNT Then
        MsgBox "Output sheet has no appointment book. Please generate it first.", vbExclamation
        Exit Sub
    End If

    Dim daysInOutput As Long
    daysInOutput = CountGeneratedDayBlocks(lastRow)

    ApplyPrintSettingsCore wsO, lastRow, daysInOutput

    MsgBox "Print settings updated. Please check Page Break Preview before printing.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error while updating print settings." & vbCrLf & _
           "Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbCritical
End Sub

Public Sub GenerateAppointmentBook_Phase5_WithOneDayPrintSettings()
    GenerateAppointmentBook_Phase5
    ApplyOneDayOnePagePrintSettings
End Sub

Private Sub ApplyPrintSettingsCore(ByVal wsO As Worksheet, ByVal lastRow As Long, ByVal daysInOutput As Long)
    Dim dayIndex As Long
    Dim nextDayStartRow As Long

    wsO.Activate

    On Error Resume Next
    wsO.ResetAllPageBreaks
    On Error GoTo 0

    With wsO.PageSetup
        .PrintArea = wsO.Range(wsO.Cells(1, OUTPUT_FIRST_COL), wsO.Cells(lastRow, OUTPUT_LAST_COL)).Address
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
    End With

    For dayIndex = 1 To daysInOutput - 1
        nextDayStartRow = 1 + dayIndex * ONE_DAY_BLOCK_STEP
        If nextDayStartRow <= lastRow Then
            wsO.HPageBreaks.Add Before:=wsO.Rows(nextDayStartRow)
        End If
    Next dayIndex
End Sub

Private Function GetLastOutputRow(ByVal wsO As Worksheet) As Long
    Dim lastCell As Range

    Set lastCell = wsO.Range(wsO.Cells(1, OUTPUT_FIRST_COL), wsO.Cells(wsO.Rows.Count, OUTPUT_LAST_COL)) _
        .Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)

    If lastCell Is Nothing Then
        GetLastOutputRow = 0
    Else
        GetLastOutputRow = lastCell.Row
    End If
End Function

Private Function CountGeneratedDayBlocks(ByVal lastRow As Long) As Long
    CountGeneratedDayBlocks = ((lastRow - 1) \ ONE_DAY_BLOCK_STEP) + 1
End Function
