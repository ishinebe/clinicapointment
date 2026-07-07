VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()

    On Error GoTo ErrorHandler

    Application.OnTime EarliestTime:=Now + TimeSerial(0, 0, 1), Procedure:="ShowAppointmentSettingsForm"
    Exit Sub

ErrorHandler:
    MsgBox "アポ帳作成フォームの自動表示に失敗しました。" & vbCrLf & _
           Err.Description, vbExclamation

End Sub
