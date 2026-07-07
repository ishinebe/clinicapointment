Option Explicit

'============================================================
' ClinicAppointment
' Module: ExceptionsDateDropdown
' Version: 2026.07.07-Phase5G-exception-date-dropdown
'
' Purpose:
' - Create a date dropdown for Exceptions!A2:A100 based on Settings!B2/B3.
' - This avoids free text date input and reduces reception-side mistakes.
'============================================================

Private Const SHEET_SETTINGS As String = "Settings"
Private Const SHEET_EXCEPTIONS As String = "Exceptions"

Private Const SETTINGS_YEAR_CELL As String = "B2"
Private Const SETTINGS_MONTH_CELL As String = "B3"

Private Const EXCEPTION_INPUT_DATE_RANGE As String = "A2:A100"
Private Const EXCEPTION_DATE_LIST_FIRST_CELL As String = "G2"
Private Const EXCEPTION_DATE_LIST_HEADER_CELL As String = "G1"

Public Sub SetupExceptionsDateDropdowns()

    On Error GoTo ErrorHandler

    Dim wsS As Worksheet
    Dim wsE As Worksheet

    Set wsS = GetSheetOrError(SHEET_SETTINGS)
    Set wsE = GetSheetOrError(SHEET_EXCEPTIONS)

    Dim targetYear As Long
    Dim targetMonth As Long

    targetYear = CLng(Val(wsS.Range(SETTINGS_YEAR_CELL).Value))
    targetMonth = CLng(Val(wsS.Range(SETTINGS_MONTH_CELL).Value))

    If targetYear < 2000 Or targetYear > 2100 Or targetMonth < 1 Or targetMonth > 12 Then
        MsgBox "Settings!B2に年、Settings!B3に月を入力してから実行してください。" & vbCrLf & _
               "例：B2=2027, B3=2", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    PrepareExceptionDateList wsE, targetYear, targetMonth
    ApplyExceptionDateValidation wsE, targetYear, targetMonth

    Application.ScreenUpdating = True

    MsgBox "臨時予定の日付候補を更新しました。" & vbCrLf & _
           targetYear & "年" & targetMonth & "月の日付から選択できます。", vbInformation
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "臨時予定の日付候補更新中にエラーが発生しました。" & vbCrLf & _
           "Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, vbCritical

End Sub

Private Sub PrepareExceptionDateList(ByVal wsE As Worksheet, ByVal targetYear As Long, ByVal targetMonth As Long)

    Dim daysInMonth As Long
    daysInMonth = Day(DateSerial(targetYear, targetMonth + 1, 0))

    wsE.Range("G1:G40").ClearContents

    wsE.Range(EXCEPTION_DATE_LIST_HEADER_CELL).Value = "日付候補"
    wsE.Range(EXCEPTION_DATE_LIST_HEADER_CELL).Font.Bold = True

    Dim d As Long

    For d = 1 To daysInMonth
        wsE.Range(EXCEPTION_DATE_LIST_FIRST_CELL).Offset(d - 1, 0).Value = DateSerial(targetYear, targetMonth, d)
        wsE.Range(EXCEPTION_DATE_LIST_FIRST_CELL).Offset(d - 1, 0).NumberFormatLocal = "yyyy/m/d (aaa)"
    Next d

    wsE.Columns("G").Hidden = True

End Sub

Private Sub ApplyExceptionDateValidation(ByVal wsE As Worksheet, ByVal targetYear As Long, ByVal targetMonth As Long)

    Dim daysInMonth As Long
    daysInMonth = Day(DateSerial(targetYear, targetMonth + 1, 0))

    Dim dateListRange As Range
    Set dateListRange = wsE.Range(EXCEPTION_DATE_LIST_FIRST_CELL).Resize(daysInMonth, 1)

    With wsE.Range(EXCEPTION_INPUT_DATE_RANGE)
        .NumberFormatLocal = "yyyy/m/d (aaa)"
        With .Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
                 Formula1:="=" & dateListRange.Address(True, True, xlA1, True)
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = "日付選択"
            .InputMessage = "当月の日付から選択してください。"
            .ErrorTitle = "入力できません"
            .ErrorMessage = "Settingsで指定した年月の日付から選択してください。"
        End With
    End With

End Sub

Private Function GetSheetOrError(ByVal sheetName As String) As Worksheet

    On Error GoTo NotFound
    Set GetSheetOrError = ThisWorkbook.Worksheets(sheetName)
    Exit Function

NotFound:
    Err.Raise vbObjectError + 100, , "Sheet not found: " & sheetName

End Function
