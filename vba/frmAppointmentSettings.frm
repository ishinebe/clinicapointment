VERSION 5.00
Begin VB.UserForm frmAppointmentSettings
   Caption         =   "月次アポ帳作成ウィザード"
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

    Me.Caption = "月次アポ帳作成ウィザード"
    Me.Width = 760
    Me.Height = 540
    Me.BackColor = RGB(248, 250, 252)

    AddCard "cardYearMonth", 16, 16, 270, 86, RGB(244, 249, 255), RGB(157, 195, 230)
    AddStepHeader "lblYearMonth", "1. 作成する年月", 32, 28, 230, RGB(47, 117, 181)
    AddNoteLabel "lblYearMonthNote", "作成する月を選びます。", 42, 54, 190, 14
    AddLabel "lblYear", "年", 42, 78, 24, 16, True, 9
    Set cmbYear = AddCombo("cmbYear", 72, 74, 72, 18)
    AddYearItems cmbYear

    AddLabel "lblMonth", "月", 162, 78, 24, 16, True, 9
    Set cmbMonth = AddCombo("cmbMonth", 192, 74, 54, 18)
    AddMonthItems cmbMonth

    AddCard "cardStaff", 300, 16, 420, 96, RGB(246, 253, 247), RGB(169, 208, 142)
    AddStepHeader "lblStaff", "2. 担当者", 316, 28, 380, RGB(34, 139, 76)
    AddNoteLabel "lblStaffNote", "各列に表示する担当者を選びます。", 326, 54, 300, 14
    AddStaffLabels 326, 72
    Set cmbDr1 = AddCombo("cmbDr1", 326, 88, 66, 18)
    Set cmbDr2 = AddCombo("cmbDr2", 400, 88, 66, 18)
    Set cmbReserve = AddCombo("cmbReserve", 474, 88, 66, 18)
    Set cmbDH1 = AddCombo("cmbDH1", 548, 88, 66, 18)
    Set cmbDH2 = AddCombo("cmbDH2", 622, 88, 66, 18)
    Set staffCombos(1) = cmbDr1
    Set staffCombos(2) = cmbDr2
    Set staffCombos(3) = cmbReserve
    Set staffCombos(4) = cmbDH1
    Set staffCombos(5) = cmbDH2

    AddCard "cardWork", 16, 122, 704, 230, RGB(248, 251, 255), RGB(157, 195, 230)
    AddStepHeader "lblWork", "3. 毎週の休み・早上がり", 32, 132, 640, RGB(47, 117, 181)
    AddNoteLabel "lblWorkNote", "空欄は通常勤務。毎週同じ休みや早上がりだけ選びます。", 42, 156, 520, 14
    AddWorkPatternGrid 42, 182
    AddGuideBox 586, 232, 110, 78

    AddCard "cardMonthlyClose", 16, 362, 320, 88, RGB(255, 250, 240), RGB(244, 176, 132)
    AddStepHeader "lblMonthlyClose", "4. 医院全体の当月終了時刻", 32, 372, 280, RGB(197, 90, 17)
    AddNoteLabel "lblMonthlyCloseNote", "その月だけ早く閉める場合に選択。", 42, 398, 230, 14
    Set cmbMonthlyClose = AddCombo("cmbMonthlyClose", 42, 420, 92, 18)
    AddMonthlyCloseItems cmbMonthlyClose

    AddCard "cardTemporary", 350, 362, 370, 88, RGB(250, 247, 255), RGB(180, 167, 214)
    AddStepHeader "lblTemporary", "5. 臨時予定の確認・編集", 366, 372, 320, RGB(112, 48, 160)
    AddNoteLabel "lblTemporaryNote", "休診、早上がり、スタッフ休みなどを確認します。", 376, 398, 280, 14
    Set btnRefreshExceptionDates = AddButton("btnRefreshExceptionDates", "臨時予定を確認・編集", 440, 414, 150, 26)

    Set btnCreate = AddButton("btnCreate", "この内容でアポ帳を作成", 180, 466, 220, 32)
    Set btnSave = AddButton("btnSave", "保存して閉じる", 430, 466, 150, 32)
    StylePrimaryButton btnCreate
    StyleSecondaryButton btnRefreshExceptionDates
    StyleSecondaryButton btnSave

End Sub

Private Sub AddCard(ByVal controlName As String, ByVal leftPos As Double, ByVal topPos As Double, _
                    ByVal controlWidth As Double, ByVal controlHeight As Double, _
                    ByVal fillColor As Long, ByVal borderColor As Long)

    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", controlName, True)

    With lbl
        .Caption = ""
        .Left = leftPos
        .Top = topPos
        .Width = controlWidth
        .Height = controlHeight
        .BackStyle = fmBackStyleOpaque
        .BackColor = fillColor
        .BorderColor = borderColor
        .BorderStyle = fmBorderStyleSingle
        .ZOrder 1
    End With

End Sub

Private Sub AddStepHeader(ByVal controlName As String, ByVal captionText As String, _
                          ByVal leftPos As Double, ByVal topPos As Double, _
                          ByVal controlWidth As Double, ByVal accentColor As Long)

    Dim stepText As String
    Dim titleText As String
    Dim dotPos As Long

    dotPos = InStr(1, captionText, ".")
    If dotPos > 0 Then
        stepText = Trim$(Left$(captionText, dotPos - 1))
        titleText = Trim$(Mid$(captionText, dotPos + 1))
    Else
        stepText = ""
        titleText = captionText
    End If

    AddStepBadge controlName & "_badge", stepText, leftPos, topPos, accentColor

    Dim lbl As MSForms.Label
    Set lbl = AddLabel(controlName, titleText, leftPos + 28, topPos + 2, controlWidth - 32, 18, True, 11)
    lbl.ForeColor = accentColor

End Sub

Private Sub AddStepBadge(ByVal controlName As String, ByVal captionText As String, _
                         ByVal leftPos As Double, ByVal topPos As Double, _
                         ByVal accentColor As Long)

    Dim lbl As MSForms.Label
    Set lbl = AddLabel(controlName, captionText, leftPos, topPos, 20, 20, True, 11)

    With lbl
        .BackStyle = fmBackStyleOpaque
        .BackColor = accentColor
        .ForeColor = RGB(255, 255, 255)
        .TextAlign = fmTextAlignCenter
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = accentColor
    End With

End Sub

Private Sub AddGuideBox(ByVal leftPos As Double, ByVal topPos As Double, _
                        ByVal controlWidth As Double, ByVal controlHeight As Double)

    AddCard "cardWorkGuide", leftPos, topPos, controlWidth, controlHeight, RGB(255, 255, 255), RGB(157, 195, 230)
    AddLabel "lblWorkGuideTitle", "選択の目安", leftPos + 10, topPos + 8, controlWidth - 20, 14, True, 9
    AddNoteLabel "lblWorkGuide1", "休: 終日休診", leftPos + 10, topPos + 28, controlWidth - 20, 12
    AddNoteLabel "lblWorkGuide2", "〜まで: その時刻で終了", leftPos + 10, topPos + 43, controlWidth - 20, 12
    AddNoteLabel "lblWorkGuide3", "空欄: 通常勤務", leftPos + 10, topPos + 58, controlWidth - 20, 12

End Sub

Private Sub AddNoteLabel(ByVal controlName As String, ByVal captionText As String, _
                         ByVal leftPos As Double, ByVal topPos As Double, _
                         ByVal controlWidth As Double, ByVal controlHeight As Double)

    Dim lbl As MSForms.Label
    Set lbl = AddLabel(controlName, captionText, leftPos, topPos, controlWidth, controlHeight, False, 9)
    lbl.ForeColor = RGB(89, 89, 89)

End Sub

Private Sub AddStaffLabels(ByVal leftStart As Double, ByVal topPos As Double)

    AddLabel "lblDr1", "Dr1", leftStart, topPos, 66, 16, True, 9
    AddLabel "lblDr2", "Dr2", leftStart + 74, topPos, 66, 16, True, 9
    AddLabel "lblReserve", "予備枠", leftStart + 148, topPos, 66, 16, True, 9
    AddLabel "lblDH1", "DH1", leftStart + 222, topPos, 66, 16, True, 9
    AddLabel "lblDH2", "DH2", leftStart + 296, topPos, 66, 16, True, 9

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
    Dim rowStep As Double

    rowStep = 19

    For c = 1 To STAFF_SLOT_COUNT
        x = leftStart + 42 + ((c - 1) * 96)
        AddLabel "lblWorkHeader" & CStr(c), CStr(slotLabels(c - 1)), x, topStart, 88, 14, True, 9
    Next c

    For r = 1 To WEEKDAY_COUNT
        y = topStart + 18 + ((r - 1) * rowStep)
        AddLabel "lblWeekday" & CStr(r), CStr(weekdayLabels(r - 1)), leftStart, y + 1, 30, 14, True, 9
        For c = 1 To STAFF_SLOT_COUNT
            x = leftStart + 42 + ((c - 1) * 96)
            Set workCombos(r, c) = AddCombo("cmbWork_" & CStr(r) & "_" & CStr(c), x, y, 88, 17)
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
        .BackStyle = fmBackStyleTransparent
        .Font.Bold = isBold
        .Font.Size = fontSize
        .WordWrap = True
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
        .Font.Size = 9
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
        .Font.Size = 9
    End With

    Set AddButton = btn

End Function

Private Sub StylePrimaryButton(ByVal btn As MSForms.CommandButton)

    With btn
        .BackColor = RGB(255, 73, 85)
        .ForeColor = RGB(255, 255, 255)
        .Font.Size = 10
        .Font.Bold = True
    End With

End Sub

Private Sub StyleSecondaryButton(ByVal btn As MSForms.CommandButton)

    With btn
        .BackColor = RGB(236, 243, 255)
        .ForeColor = RGB(31, 78, 121)
        .Font.Bold = True
    End With

End Sub

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

    If SaveSettings(True, False) Then Unload Me

End Sub

Private Sub btnCreate_Click()

    If SaveSettings(False, False) Then
        Me.Hide
        Call GenerateBook_Phase5_Print
        Unload Me
    End If

End Sub

Private Sub btnRefreshExceptionDates_Click()

    If SaveSettings(False, True) Then
        SetupExceptionsDateDropdowns
        frmTemporarySchedule.Show
    End If

End Sub
