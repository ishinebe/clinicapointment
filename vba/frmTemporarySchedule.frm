VERSION 5.00
Begin VB.UserForm frmTemporarySchedule
   Caption         =   "臨時予定を編集"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11200
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTemporarySchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SHEET_SETTINGS As String = "Settings"
Private Const SHEET_EXCEPTIONS As String = "Exceptions"
Private Const SETTINGS_YEAR_CELL As String = "B2"
Private Const SETTINGS_MONTH_CELL As String = "B3"

Private Const INPUT_FIRST_ROW As Long = 2
Private Const INPUT_LAST_ROW As Long = 100

Private cmbTempDate As MSForms.ComboBox
Private cmbTempType As MSForms.ComboBox
Private cmbTempTarget As MSForms.ComboBox
Private cmbTempCloseTime As MSForms.ComboBox
Private txtTempMemo As MSForms.TextBox
Private WithEvents lstTemporarySchedules As MSForms.ListBox

Private WithEvents btnAddTemporarySchedule As MSForms.CommandButton
Private WithEvents btnUpdateTemporarySchedule As MSForms.CommandButton
Private WithEvents btnDeleteTemporarySchedule As MSForms.CommandButton
Private WithEvents btnRefreshTemporaryDates As MSForms.CommandButton
Private WithEvents btnCloseTemporarySchedule As MSForms.CommandButton

Private Sub UserForm_Initialize()

    On Error GoTo ErrorHandler

    BuildForm
    LoadDateCandidates
    LoadTemporarySchedules
    Exit Sub

ErrorHandler:
    MsgBox "臨時予定フォームの初期化中にエラーが発生しました。" & vbCrLf & _
           Err.Description, vbCritical

End Sub

Private Sub BuildForm()

    Me.Caption = "臨時予定を編集"
    Me.Width = 760
    Me.Height = 520

    AddLabel "lblTitle", "臨時予定を編集", 12, 10, 180, 20, True, 14

    AddLabel "lblDate", "日付", 22, 48, 80, 16, True, 10
    Set cmbTempDate = AddCombo("cmbTempDate", 110, 44, 150, 20)

    AddLabel "lblType", "内容", 22, 78, 80, 16, True, 10
    Set cmbTempType = AddCombo("cmbTempType", 110, 74, 190, 20)
    AddTypeItems cmbTempType

    AddLabel "lblTarget", "対象", 22, 108, 80, 16, True, 10
    Set cmbTempTarget = AddCombo("cmbTempTarget", 110, 104, 100, 20)
    AddTargetItems cmbTempTarget

    AddLabel "lblClose", "終了時刻", 230, 108, 80, 16, True, 10
    Set cmbTempCloseTime = AddCombo("cmbTempCloseTime", 310, 104, 90, 20)
    AddCloseTimeItems cmbTempCloseTime

    AddLabel "lblMemo", "メモ", 22, 138, 80, 16, True, 10
    Set txtTempMemo = AddTextBox("txtTempMemo", 110, 134, 290, 20)

    Set btnAddTemporarySchedule = AddButton("btnAddTemporarySchedule", "追加", 420, 44, 80, 26)
    Set btnUpdateTemporarySchedule = AddButton("btnUpdateTemporarySchedule", "更新", 510, 44, 80, 26)
    Set btnDeleteTemporarySchedule = AddButton("btnDeleteTemporarySchedule", "削除", 600, 44, 80, 26)
    Set btnRefreshTemporaryDates = AddButton("btnRefreshTemporaryDates", "日付候補を更新", 420, 84, 130, 26)
    Set btnCloseTemporarySchedule = AddButton("btnCloseTemporarySchedule", "閉じる", 560, 84, 80, 26)

    AddLabel "lblList", "登録済みの臨時予定", 22, 178, 150, 16, True, 10
    Set lstTemporarySchedules = Me.Controls.Add("Forms.ListBox.1", "lstTemporarySchedules", True)
    With lstTemporarySchedules
        .Left = 22
        .Top = 200
        .Width = 680
        .Height = 250
        .ColumnCount = 6
        .ColumnWidths = "90 pt;150 pt;60 pt;70 pt;210 pt;0 pt"
    End With

End Sub

Private Function AddLabel(ByVal controlName As String, ByVal captionText As String, _
                          ByVal leftPos As Double, ByVal topPos As Double, _
                          ByVal controlWidth As Double, ByVal controlHeight As Double, _
                          ByVal isBold As Boolean, ByVal fontSize As Double) As MSForms.Label

    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", controlName, True)

    With lbl
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = controlWidth
        .Height = controlHeight
        .Font.Bold = isBold
        .Font.Size = fontSize
    End With

    Set AddLabel = lbl

End Function

Private Function AddCombo(ByVal controlName As String, ByVal leftPos As Double, ByVal topPos As Double, _
                          ByVal controlWidth As Double, ByVal controlHeight As Double) As MSForms.ComboBox

    Dim cmb As MSForms.ComboBox
    Set cmb = Me.Controls.Add("Forms.ComboBox.1", controlName, True)

    With cmb
        .Left = leftPos
        .Top = topPos
        .Width = controlWidth
        .Height = controlHeight
        .Style = fmStyleDropDownList
        .MatchRequired = False
    End With

    Set AddCombo = cmb

End Function

Private Function AddTextBox(ByVal controlName As String, ByVal leftPos As Double, ByVal topPos As Double, _
                            ByVal controlWidth As Double, ByVal controlHeight As Double) As MSForms.TextBox

    Dim txt As MSForms.TextBox
    Set txt = Me.Controls.Add("Forms.TextBox.1", controlName, True)

    With txt
        .Left = leftPos
        .Top = topPos
        .Width = controlWidth
        .Height = controlHeight
    End With

    Set AddTextBox = txt

End Function

Private Function AddButton(ByVal controlName As String, ByVal captionText As String, _
                           ByVal leftPos As Double, ByVal topPos As Double, _
                           ByVal controlWidth As Double, ByVal controlHeight As Double) As MSForms.CommandButton

    Dim btn As MSForms.CommandButton
    Set btn = Me.Controls.Add("Forms.CommandButton.1", controlName, True)

    With btn
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = controlWidth
        .Height = controlHeight
        .Font.Bold = True
    End With

    Set AddButton = btn

End Function

Private Sub AddTypeItems(ByVal cmb As MSForms.ComboBox)

    cmb.Clear
    cmb.AddItem ""
    cmb.AddItem "医院全体を休みにする"
    cmb.AddItem "医院全体を早く閉める"
    cmb.AddItem "スタッフが休み"
    cmb.AddItem "スタッフが早上がり"

End Sub

Private Sub AddTargetItems(ByVal cmb As MSForms.ComboBox)

    cmb.Clear
    cmb.AddItem "全体"
    cmb.AddItem "Dr1"
    cmb.AddItem "Dr2"
    cmb.AddItem "予備"
    cmb.AddItem "DH1"
    cmb.AddItem "DH2"

End Sub

Private Sub AddCloseTimeItems(ByVal cmb As MSForms.ComboBox)

    Dim items As Variant
    items = Array("", "12:00", "15:00", "15:30", "16:00", "16:30", "17:00", "17:30", "18:00", "18:30", "19:00")

    Dim i As Long
    cmb.Clear
    For i = LBound(items) To UBound(items)
        cmb.AddItem CStr(items(i))
    Next i

End Sub

Private Sub LoadDateCandidates()

    cmbTempDate.Clear
    cmbTempDate.AddItem ""

    Dim wsS As Worksheet
    Set wsS = GetSheetOrError(SHEET_SETTINGS)

    Dim targetYear As Long
    Dim targetMonth As Long
    targetYear = CLng(Val(wsS.Range(SETTINGS_YEAR_CELL).Value))
    targetMonth = CLng(Val(wsS.Range(SETTINGS_MONTH_CELL).Value))

    If targetYear >= 2000 And targetYear <= 2100 And targetMonth >= 1 And targetMonth <= 12 Then
        Dim daysInMonth As Long
        Dim d As Long
        daysInMonth = Day(DateSerial(targetYear, targetMonth + 1, 0))

        For d = 1 To daysInMonth
            cmbTempDate.AddItem Format$(DateSerial(targetYear, targetMonth, d), "yyyy/m/d (aaa)")
        Next d
        Exit Sub
    End If

    If SheetExists(SHEET_EXCEPTIONS) Then
        Dim wsE As Worksheet
        Dim cell As Range
        Set wsE = ThisWorkbook.Worksheets(SHEET_EXCEPTIONS)

        For Each cell In wsE.Range("A" & INPUT_FIRST_ROW & ":A" & INPUT_LAST_ROW).Cells
            If IsDate(cell.Value) Then cmbTempDate.AddItem Format$(DateValue(cell.Value), "yyyy/m/d (aaa)")
        Next cell
    End If

End Sub

Private Sub LoadTemporarySchedules()

    EnsureTemporaryScheduleSheet

    Dim wsE As Worksheet
    Set wsE = ThisWorkbook.Worksheets(SHEET_EXCEPTIONS)

    lstTemporarySchedules.Clear

    Dim r As Long
    Dim rowIndex As Long

    For r = INPUT_FIRST_ROW To INPUT_LAST_ROW
        If Application.WorksheetFunction.CountA(wsE.Range("A" & r & ":E" & r)) > 0 Then
            lstTemporarySchedules.AddItem FormatDisplayDate(wsE.Cells(r, "A").Value)
            rowIndex = lstTemporarySchedules.ListCount - 1
            lstTemporarySchedules.List(rowIndex, 1) = InternalTypeToDisplay(CStr(wsE.Cells(r, "B").Value))
            lstTemporarySchedules.List(rowIndex, 2) = CStr(wsE.Cells(r, "C").Value)
            lstTemporarySchedules.List(rowIndex, 3) = CStr(wsE.Cells(r, "D").Text)
            lstTemporarySchedules.List(rowIndex, 4) = CStr(wsE.Cells(r, "E").Value)
            lstTemporarySchedules.List(rowIndex, 5) = CStr(r)
        End If
    Next r

End Sub

Private Sub EnsureTemporaryScheduleSheet()

    If Not SheetExists(SHEET_EXCEPTIONS) Then
        ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)).Name = SHEET_EXCEPTIONS
    End If

    Dim wsE As Worksheet
    Set wsE = ThisWorkbook.Worksheets(SHEET_EXCEPTIONS)

    wsE.Range("A1").Value = "日付"
    wsE.Range("B1").Value = "種別"
    wsE.Range("C1").Value = "対象"
    wsE.Range("D1").Value = "終了時刻"
    wsE.Range("E1").Value = "メモ"

End Sub

Private Function FormatDisplayDate(ByVal valueDate As Variant) As String

    If IsDate(valueDate) Then
        FormatDisplayDate = Format$(DateValue(valueDate), "yyyy/m/d (aaa)")
    Else
        FormatDisplayDate = CStr(valueDate)
    End If

End Function

Private Function DisplayTypeToInternal(ByVal displayText As String) As String

    Select Case Trim$(displayText)
        Case "医院全体を休みにする"
            DisplayTypeToInternal = "休診"
        Case "医院全体を早く閉める"
            DisplayTypeToInternal = "時短診療"
        Case "スタッフが休み"
            DisplayTypeToInternal = "休み"
        Case "スタッフが早上がり"
            DisplayTypeToInternal = "時短勤務"
        Case Else
            DisplayTypeToInternal = ""
    End Select

End Function

Private Function InternalTypeToDisplay(ByVal internalText As String) As String

    Select Case Trim$(internalText)
        Case "休診"
            InternalTypeToDisplay = "医院全体を休みにする"
        Case "時短診療"
            InternalTypeToDisplay = "医院全体を早く閉める"
        Case "休み"
            InternalTypeToDisplay = "スタッフが休み"
        Case "時短勤務"
            InternalTypeToDisplay = "スタッフが早上がり"
        Case Else
            InternalTypeToDisplay = Trim$(internalText)
    End Select

End Function

Private Function TryGetSelectedDate(ByRef selectedDate As Date) As Boolean

    Dim valueText As String
    valueText = Trim$(cmbTempDate.Value)

    If Len(valueText) = 0 Then
        TryGetSelectedDate = False
        Exit Function
    End If

    If InStr(valueText, " ") > 0 Then valueText = Left$(valueText, InStr(valueText, " ") - 1)

    If IsDate(valueText) Then
        selectedDate = DateValue(valueText)
        TryGetSelectedDate = True
    Else
        TryGetSelectedDate = False
    End If

End Function

Private Function ValidateTemporarySchedule(ByRef selectedDate As Date, ByRef internalType As String) As Boolean

    If Not TryGetSelectedDate(selectedDate) Then
        MsgBox "日付を選択してください。", vbExclamation
        ValidateTemporarySchedule = False
        Exit Function
    End If

    internalType = DisplayTypeToInternal(cmbTempType.Value)
    If Len(internalType) = 0 Then
        MsgBox "内容を選択してください。", vbExclamation
        ValidateTemporarySchedule = False
        Exit Function
    End If

    Select Case internalType
        Case "休診"
            cmbTempTarget.Value = "全体"
            cmbTempCloseTime.Value = ""

        Case "時短診療"
            cmbTempTarget.Value = "全体"
            If Len(Trim$(cmbTempCloseTime.Value)) = 0 Then
                MsgBox "終了時刻を選択してください。", vbExclamation
                ValidateTemporarySchedule = False
                Exit Function
            End If

        Case "休み"
            If Not IsStaffTarget(cmbTempTarget.Value) Then
                MsgBox "対象スタッフを選択してください。", vbExclamation
                ValidateTemporarySchedule = False
                Exit Function
            End If
            cmbTempCloseTime.Value = ""

        Case "時短勤務"
            If Not IsStaffTarget(cmbTempTarget.Value) Then
                MsgBox "対象スタッフを選択してください。", vbExclamation
                ValidateTemporarySchedule = False
                Exit Function
            End If
            If Len(Trim$(cmbTempCloseTime.Value)) = 0 Then
                MsgBox "終了時刻を選択してください。", vbExclamation
                ValidateTemporarySchedule = False
                Exit Function
            End If
    End Select

    ValidateTemporarySchedule = True

End Function

Private Function IsStaffTarget(ByVal targetText As String) As Boolean

    Select Case Trim$(targetText)
        Case "Dr1", "Dr2", "予備", "DH1", "DH2"
            IsStaffTarget = True
        Case Else
            IsStaffTarget = False
    End Select

End Function

Private Sub WriteTemporaryScheduleRow(ByVal rowNumber As Long, ByVal selectedDate As Date, ByVal internalType As String)

    Dim wsE As Worksheet
    Set wsE = ThisWorkbook.Worksheets(SHEET_EXCEPTIONS)

    wsE.Cells(rowNumber, "A").Value = selectedDate
    wsE.Cells(rowNumber, "A").NumberFormatLocal = "yyyy/m/d (aaa)"
    wsE.Cells(rowNumber, "B").Value = internalType
    wsE.Cells(rowNumber, "C").Value = cmbTempTarget.Value
    wsE.Cells(rowNumber, "D").Value = cmbTempCloseTime.Value
    wsE.Cells(rowNumber, "E").Value = txtTempMemo.Value

End Sub

Private Function NextInputRow(ByVal wsE As Worksheet) As Long

    Dim r As Long

    For r = INPUT_FIRST_ROW To INPUT_LAST_ROW
        If Application.WorksheetFunction.CountA(wsE.Range("A" & r & ":E" & r)) = 0 Then
            NextInputRow = r
            Exit Function
        End If
    Next r

    NextInputRow = INPUT_LAST_ROW + 1

End Function

Private Sub ClearEntryFields()

    cmbTempDate.Value = ""
    cmbTempType.Value = ""
    cmbTempTarget.Value = "全体"
    cmbTempCloseTime.Value = ""
    txtTempMemo.Value = ""

End Sub

Private Sub lstTemporarySchedules_Click()

    If lstTemporarySchedules.ListIndex < 0 Then Exit Sub

    SetComboValue cmbTempDate, lstTemporarySchedules.List(lstTemporarySchedules.ListIndex, 0)
    SetComboValue cmbTempType, lstTemporarySchedules.List(lstTemporarySchedules.ListIndex, 1)
    SetComboValue cmbTempTarget, lstTemporarySchedules.List(lstTemporarySchedules.ListIndex, 2)
    SetComboValue cmbTempCloseTime, lstTemporarySchedules.List(lstTemporarySchedules.ListIndex, 3)
    txtTempMemo.Value = lstTemporarySchedules.List(lstTemporarySchedules.ListIndex, 4)

End Sub

Private Function ComboContains(ByVal cmb As MSForms.ComboBox, ByVal valueText As String) As Boolean

    Dim i As Long

    For i = 0 To cmb.ListCount - 1
        If CStr(cmb.List(i)) = valueText Then
            ComboContains = True
            Exit Function
        End If
    Next i

    ComboContains = False

End Function

Private Sub SetComboValue(ByVal cmb As MSForms.ComboBox, ByVal valueText As String)

    If Not ComboContains(cmb, valueText) Then cmb.AddItem valueText
    cmb.Value = valueText

End Sub

Private Sub btnAddTemporarySchedule_Click()

    On Error GoTo ErrorHandler

    EnsureTemporaryScheduleSheet

    Dim selectedDate As Date
    Dim internalType As String
    If Not ValidateTemporarySchedule(selectedDate, internalType) Then Exit Sub

    Dim wsE As Worksheet
    Set wsE = ThisWorkbook.Worksheets(SHEET_EXCEPTIONS)

    Dim rowNumber As Long
    rowNumber = NextInputRow(wsE)

    If rowNumber > INPUT_LAST_ROW Then
        MsgBox "入力できる行がありません。", vbExclamation
        Exit Sub
    End If

    WriteTemporaryScheduleRow rowNumber, selectedDate, internalType
    LoadTemporarySchedules
    ClearEntryFields
    Exit Sub

ErrorHandler:
    MsgBox "臨時予定の追加中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical

End Sub

Private Sub btnUpdateTemporarySchedule_Click()

    On Error GoTo ErrorHandler

    If lstTemporarySchedules.ListIndex < 0 Then
        MsgBox "更新する臨時予定を一覧から選択してください。", vbExclamation
        Exit Sub
    End If

    Dim selectedDate As Date
    Dim internalType As String
    If Not ValidateTemporarySchedule(selectedDate, internalType) Then Exit Sub

    Dim rowNumber As Long
    rowNumber = CLng(lstTemporarySchedules.List(lstTemporarySchedules.ListIndex, 5))

    WriteTemporaryScheduleRow rowNumber, selectedDate, internalType
    LoadTemporarySchedules
    ClearEntryFields
    Exit Sub

ErrorHandler:
    MsgBox "臨時予定の更新中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical

End Sub

Private Sub btnDeleteTemporarySchedule_Click()

    On Error GoTo ErrorHandler

    If lstTemporarySchedules.ListIndex < 0 Then
        MsgBox "削除する臨時予定を一覧から選択してください。", vbExclamation
        Exit Sub
    End If

    If MsgBox("選択中の臨時予定を削除しますか。", vbQuestion + vbYesNo) <> vbYes Then Exit Sub

    Dim rowNumber As Long
    rowNumber = CLng(lstTemporarySchedules.List(lstTemporarySchedules.ListIndex, 5))

    Dim wsE As Worksheet
    Set wsE = ThisWorkbook.Worksheets(SHEET_EXCEPTIONS)
    wsE.Range("A" & rowNumber & ":E" & rowNumber).Delete Shift:=xlUp

    LoadTemporarySchedules
    ClearEntryFields
    Exit Sub

ErrorHandler:
    MsgBox "臨時予定の削除中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical

End Sub

Private Sub btnRefreshTemporaryDates_Click()

    SetupExceptionsDateDropdowns
    LoadDateCandidates

End Sub

Private Sub btnCloseTemporarySchedule_Click()

    Unload Me

End Sub

Private Function SheetExists(ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    SheetExists = Not ws Is Nothing

End Function

Private Function GetSheetOrError(ByVal sheetName As String) As Worksheet

    On Error GoTo NotFound
    Set GetSheetOrError = ThisWorkbook.Worksheets(sheetName)
    Exit Function

NotFound:
    Err.Raise vbObjectError + 300, , "必要なシートが見つかりません: " & sheetName

End Function
