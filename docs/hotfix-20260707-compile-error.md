# Hotfix: ApplyOneDayPageBreaks compile error

## 症状

VBEコンパイル時に以下のエラーが出る。

```text
コンパイル エラー:
Sub, Function, または Property が必要です。
```

該当箇所は `GenerateAppointmentBookCore` 内の以下。

```vba
If applyOneDayPageBreaks Then
    ApplyOneDayPageBreaks wsO, templateRange, daysInMonth
End If
```

## 原因

`GenerateAppointmentBookCore` の引数名 `applyOneDayPageBreaks` と、Sub名 `ApplyOneDayPageBreaks` が同名扱いになっている。

VBAは大文字・小文字を区別しないため、`ApplyOneDayPageBreaks wsO, ...` がSub呼び出しではなく、Boolean引数への呼び出しのように解釈され、コンパイルエラーになる。

## 修正内容

`vba/AppointmentBook_Phase1.bas` で、引数名だけを変更する。

### 修正前

```vba
Private Sub GenerateAppointmentBookCore(ByVal phaseName As String, Optional ByVal applyOneDayPageBreaks As Boolean = False)
```

```vba
If applyOneDayPageBreaks Then
    ApplyOneDayPageBreaks wsO, templateRange, daysInMonth
End If
```

### 修正後

```vba
Private Sub GenerateAppointmentBookCore(ByVal phaseName As String, Optional ByVal shouldApplyOneDayPageBreaks As Boolean = False)
```

```vba
If shouldApplyOneDayPageBreaks Then
    ApplyOneDayPageBreaks wsO, templateRange, daysInMonth
End If
```

## 注意

Sub名 `ApplyOneDayPageBreaks` は変更しない。
変更するのは `GenerateAppointmentBookCore` の Optional 引数名と、その判定箇所のみ。
