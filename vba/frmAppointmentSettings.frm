VERSION 5.00
Begin VB.UserForm frmAppointmentSettings
   Caption         =   "アポ帳作成"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10500
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAppointmentSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SHEET_SETTINGS As String = "Settings"
Private Const SETTINGS_YEAR_CELL As String = "B2"
Private Const SETTINGS_MONTH_CELL As String = "B3"
Private Const SETTINGS_STAFF_HEADER_RANGE As String = "B5:F5"
Private Const SETTINGS_WORK_PATTERN_RANGE As String = "B7:F13"
Private Const SETTINGS_STAFF_MASTER_RANGE As String = "H5:H24"
Private Const SETTINGS_MONTHLY_CLOSE_CELL As String = "B16"

Private Const STAFF_SLOT_COUNT As Long = 5
Private Const WEEKDAY_COUNT As Long = 7

Private cmbYear As MSForms.ComboBox
Private cmbMonth As MSForms.ComboBox
Private cmbDr1 As MSForms.ComboBox
Private cmbDr2 As MSForms.ComboBox
Private cmbReserve As MSForms.ComboBox
Private cmbDH1 As MSForms.ComboBox
Private cmbDH2 As MSForms.ComboBox
Private cmbMonthlyClose As MSForms.ComboBox
Private staffCombos(1 To STAFF_SLOT_COUNT) As MSForms.ComboBox
Private workCombos(1 To WEEKDAY_COUNT, 1 To STAFF_SLOT_COUNT) As MSForms.ComboBox

Private WithEvents btnSave As MSForms.CommandButton
Private WithEvents btnCreate As MSForms.CommandButton
Private WithEvents btnRefreshExceptionDates As MSForms.CommandButton
Private WithEvents btnClose As MSForms.CommandButton

Private Sub UserForm_Initialize()

    On Error GoTo ErrorHandler

    BuildForm
    LoadSettingsToForm
    Exit Sub

ErrorHandler:
    MsgBox "アポ帳作成フォームの初期化中にエラーが発生しました。" & vbCrLf & _
           Err.Description, vbCritical

End Sub

Private Sub BuildForm()

    Me.Caption = "アポ帳作成"
    Me.Width = 720
    Me.Height = 520

    AddLabel "lblTitle", "月次アポ帳作成フォーム", 12, 10, 260, 20, True, 14

    AddLabel "lblYearMonth", "作成年月", 12, 42, 80, 18, True, 11
    AddLabel "lblYear", "年", 28, 68, 24, 18, True, 10
    Set cmbYear = AddCombo("cmbYear", 58, 64, 86, 20)
    AddYearItems cmbYear

    AddLabel "lblMonth", "月", 166, 68, 24, 18, True, 10
    Set cmbMonth = AddCombo("cmbMonth", 196, 64, 62, 20)
    AddMonthItems cmbMonth

    AddLabel "lblStaff", "担当者", 12, 104, 80, 18, True, 11
    AddLabel "lblStaffNote", "担当者マスターから選択します。予備枠は空欄または予備枠も選べます。", 88, 104, 430, 18, False, 9
    AddStaffLabels 32, 130
    Set cmbDr1 = AddCombo("cmbDr1", 82, 150, 88, 20)
    Set cmbDr2 = AddCombo("cmbDr2", 178, 150, 88, 20)
    Set cmbReserve = AddCombo("cmbReserve", 274, 150, 88, 20)
    Set cmbDH1 = AddCombo("cmbDH1", 370, 150, 88, 20)
    Set cmbDH2 = AddCombo("cmbDH2", 466, 150, 88, 20)
    Set staffCombos(1) = cmbDr1
    Set staffCombos(2) = cmbDr2
    Set staffCombos(3) = cmbReserve
    Set staffCombos(4) = cmbDH1
    Set staffCombos(5) = cmbDH2

    AddLabel "lblWork", "曜日別の休み・早上がり", 12, 190, 170, 18, True, 11
    AddLabel "lblWorkNote", "空欄=通常、休=休み、午前のみ=午前勤務、xx:xxまで=早上がり", 184, 190, 430, 18, False, 9
    AddWorkPatternGrid 26, 218

    AddLabel "lblMonthlyClose", "医院全体の当月終了時刻", 12, 388, 170, 18, True, 11
    AddLabel "lblMonthlyCloseNote", "通常は空欄。その月だけ早く閉める場合に選択します。", 184, 388, 360, 18, False, 9
    Set cmbMonthlyClose = AddCombo("cmbMonthlyClose", 184, 414, 92, 20)
    AddMonthlyCloseItems cmbMonthlyClose

    Set btnSave = AddButton("btnSave", "設定を保存", 24, 458, 110, 28)
    Set btnCreate = AddButton("btnCreate", "アポ帳を作成", 148, 458, 120, 28)
    Set btnRefreshExceptionDates = AddButton("btnRefreshExceptionDates", "臨時予定を編集", 282, 458, 130, 28)
    Set btnClose = AddButton("btnClose", "閉じる", 426, 458, 90, 28)

End Sub

Private Sub AddStaffLabels(ByVal leftStart As Double, ByVal topPos As Double)

    AddLabel "lblDr1", "Dr1", leftStart + 50, topPos, 88, 16, True, 9
    AddLabel "lblDr2", "Dr2", leftStart + 146, topPos, 88, 16, True, 9
    AddLabel "lblReserve", "予備枠", leftStart + 242, topPos, 88, 16, True, 9
    AddLabel "lblDH1", "DH1", leftStart + 338, topPos, 88, 16, True, 9
    AddLabel "lblDH2", "DH2", leftStart + 434, topPos, 88, 16, True, 9

End Sub

Private Sub AddWorkPatternGrid(ByVal leftStart As Double, ByVal topStart As Double)

    Dim slotLabels As Variant
    slotLabels = Array("Dr1", "Dr2", "予備枠", "DH1", "DH2")

    Dim weekdayLabels As Variant
    weekdayLabels = Array("月", "火", "水", "木", "金", "土", "日")

    Dim c As Long
    Dim r As Long
    Dim x As Double
    Dim y As Double

    For c = 1 To STAFF_SLOT_COUNT
        x = leftStart + 42 + ((c - 1) * 96)
        AddLabel "lblWorkHeader" & CStr(c), CStr(slotLabels(c - 1)), x, topStart, 88, 16, True, 9
    Next c

    For r = 1 To WEEKDAY_COUNT
        y = topStart + 22 + ((r - 1) * 22)
        AddLabel "lblWeekday" & CStr(r), CStr(weekdayLabels(r - 1)), leftStart, y + 2, 30, 16, True, 9
        For c = 1 To STAFF_SLOT_COUNT
            x = leftStart + 42 + ((c - 1) * 96)
            Set workCombos(r, c) = AddCombo("cmbWork_" & CStr(r) & "_" & CStr(c), x, y, 88, 18)
            AddWorkPatternItems workCombos(r, c)
        Next c
    Next r

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

Private Sub LoadSettingsToForm()

    Dim wsS As Worksheet
    Set wsS = GetSettingsSheet()

    FillStaffCandidateCombos wsS

    SetComboValue cmbYear, CStr(wsS.Range(SETTINGS_YEAR_CELL).Value)
    SetComboValue cmbMonth, CStr(wsS.Range(SETTINGS_MONTH_CELL).Value)

    Dim i As Long
    For i = 1 To STAFF_SLOT_COUNT
        SetComboValue staffCombos(i), CStr(wsS.Range(SETTINGS_STAFF_HEADER_RANGE).Cells(1, i).Value)
    Next i

    Dim r As Long
    Dim c As Long
    For r = 1 To WEEKDAY_COUNT
        For c = 1 To STAFF_SLOT_COUNT
            SetComboValue workCombos(r, c), CStr(wsS.Range(SETTINGS_WORK_PATTERN_RANGE).Cells(r, c).Value)
        Next c
    Next r

    SetComboValue cmbMonthlyClose, CStr(wsS.Range(SETTINGS_MONTHLY_CLOSE_CELL).Value)

End Sub

Private Sub FillStaffCandidateCombos(ByVal wsS As Worksheet)

    Dim candidates As Collection
    Set candidates = New Collection

    Dim cell As Range
    For Each cell In wsS.Range(SETTINGS_STAFF_MASTER_RANGE).Cells
        AddUniqueCandidate candidates, Trim$(CStr(cell.Value))
    Next cell

    For Each cell In wsS.Range(SETTINGS_STAFF_HEADER_RANGE).Cells
        AddUniqueCandidate candidates, Trim$(CStr(cell.Value))
    Next cell

    Dim i As Long
    For i = 1 To STAFF_SLOT_COUNT
        FillStaffCombo staffCombos(i), candidates, (i = 3)
    Next i

End Sub

Private Sub FillStaffCombo(ByVal cmb As MSForms.ComboBox, ByVal candidates As Collection, ByVal includeReserve As Boolean)

    cmb.Clear
    cmb.AddItem ""

    Dim item As Variant
    For Each item In candidates
        If Len(CStr(item)) > 0 Then cmb.AddItem CStr(item)
    Next item

    If includeReserve Then
        If Not ComboContains(cmb, "予備枠") Then cmb.AddItem "予備枠"
    End If

End Sub

Private Sub AddYearItems(ByVal cmb As MSForms.ComboBox)

    Dim y As Long
    For y = 2026 To 2035
        cmb.AddItem CStr(y)
    Next y

End Sub

Private Sub AddMonthItems(ByVal cmb As MSForms.ComboBox)

    Dim m As Long
    For m = 1 To 12
        cmb.AddItem CStr(m)
    Next m

End Sub

Private Sub AddWorkPatternItems(ByVal cmb As MSForms.ComboBox)

    Dim items As Variant
    items = Array("", "休", "午前のみ", "15:00まで", "15:30まで", "16:00まで", _
                  "16:30まで", "17:00まで", "17:30まで", "18:00まで", "18:30まで")

    AddItemsFromArray cmb, items

End Sub

Private Sub AddMonthlyCloseItems(ByVal cmb As MSForms.ComboBox)

    Dim items As Variant
    items = Array("", "15:00", "15:30", "16:00", "16:30", "17:00", "17:30", "18:00", "18:30", "19:00")

    AddItemsFromArray cmb, items

End Sub

Private Sub AddItemsFromArray(ByVal cmb As MSForms.ComboBox, ByVal items As Variant)

    Dim i As Long
    cmb.Clear

    For i = LBound(items) To UBound(items)
        cmb.AddItem CStr(items(i))
    Next i

End Sub

Private Sub AddUniqueCandidate(ByVal candidates As Collection, ByVal valueText As String)

    If Len(valueText) = 0 Then Exit Sub

    On Error Resume Next
    candidates.Add valueText, valueText
    On Error GoTo 0

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

Private Function SaveSettings(Optional ByVal showMessage As Boolean = True, _
                              Optional ByVal saveYearMonthOnly As Boolean = False) As Boolean

    On Error GoTo ErrorHandler

    Dim targetYear As Long
    Dim targetMonth As Long

    If Not ValidateYearMonth(targetYear, targetMonth) Then
        SaveSettings = False
        Exit Function
    End If

    Dim wsS As Worksheet
    Set wsS = GetSettingsSheet()

    wsS.Range(SETTINGS_YEAR_CELL).Value = targetYear
    wsS.Range(SETTINGS_MONTH_CELL).Value = targetMonth

    If Not saveYearMonthOnly Then
        Dim i As Long
        For i = 1 To STAFF_SLOT_COUNT
            wsS.Range(SETTINGS_STAFF_HEADER_RANGE).Cells(1, i).Value = staffCombos(i).Value
        Next i

        Dim r As Long
        Dim c As Long
        For r = 1 To WEEKDAY_COUNT
            For c = 1 To STAFF_SLOT_COUNT
                wsS.Range(SETTINGS_WORK_PATTERN_RANGE).Cells(r, c).Value = workCombos(r, c).Value
            Next c
        Next r

        wsS.Range(SETTINGS_MONTHLY_CLOSE_CELL).Value = cmbMonthlyClose.Value
    End If

    If showMessage Then MsgBox "設定を保存しました。", vbInformation
    SaveSettings = True
    Exit Function

ErrorHandler:
    MsgBox "設定の保存中にエラーが発生しました。" & vbCrLf & Err.Description, vbCritical
    SaveSettings = False

End Function

Private Function ValidateYearMonth(ByRef targetYear As Long, ByRef targetMonth As Long) As Boolean

    targetYear = CLng(Val(cmbYear.Value))
    targetMonth = CLng(Val(cmbMonth.Value))

    If targetYear < 2000 Or targetYear > 2100 Or targetMonth < 1 Or targetMonth > 12 Then
        MsgBox "年と月を正しく選択してください。", vbExclamation
        ValidateYearMonth = False
        Exit Function
    End If

    ValidateYearMonth = True

End Function

Private Function GetSettingsSheet() As Worksheet

    On Error GoTo NotFound
    Set GetSettingsSheet = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    Exit Function

NotFound:
    Err.Raise vbObjectError + 200, , "Settingsシートが見つかりません。"

End Function

Private Sub btnSave_Click()

    SaveSettings True, False

End Sub

Private Sub btnCreate_Click()

    If SaveSettings(False, False) Then
        Me.Hide
        GenerateAppointmentBook_Phase5
        Unload Me
    End If

End Sub

Private Sub btnRefreshExceptionDates_Click()

    If SaveSettings(False, True) Then
        SetupExceptionsDateDropdowns
        frmTemporarySchedule.Show
    End If

End Sub

Private Sub btnClose_Click()

    Unload Me

End Sub
