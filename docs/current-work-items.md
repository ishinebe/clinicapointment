# Current work items

ClinicAppointment の現在の実装状況と、次に対応する作業を整理する。

## 現在の状態

### 実装済み

- `Template!A1:J46` を1日分として、1か月分の `Output` を生成する。
- 曜日別診療時間、スタッフ別勤務パターン、医院全体の当月原則終了時刻を反映する。
- `Exceptions` シートに保存した臨時予定を反映する。
- `frmAppointmentSettings` を月次アポ帳作成ウィザードに近い構成へ刷新済み。
- `frmTemporarySchedule` による臨時予定フォームを追加済み。
- `AppointmentBook_Phase1.bas` の `GenerateBook_Phase5_Print` に、行高コピー、印刷範囲設定、1日1ページの改ページ設定、網掛け開始セルの上罫線追加を統合済み。

### 注意点

- フォームと `Settings` シート上の作成ボタンは `GenerateBook_Phase5_Print` に統一する。従来の `GenerateAppointmentBook_Phase5` は互換用ラッパーとして残す。
- 起動時フォーム表示は `ThisWorkbook.cls` の `Workbook_Open` 貼り付けが前提であり、実ファイル側での反映確認が必要。
- UserFormはVBA標準部品の範囲で、手順見出し、説明文、現場向けボタン文言を整理済み。

## 次に対応するIssue

### #2 網掛け開始セルの上罫線をさらに太くして視認性を上げる

現在は `xlMedium` を使用している。印刷時の視認性を高めるため、`xlThick` への変更を検討する。

### #3 フォームからアポ帳作成した場合も印刷設定・罫線処理を適用する

フォームの作成ボタンと `Settings` シートの作成ボタンの呼び出し先を、以下へ統一する。

```vba
GenerateBook_Phase5_Print
```

### #4 Excel起動時に月次アポ帳作成フォームが確実に開くようにする

`ThisWorkbook.cls` の `Workbook_Open` が実ファイル側で正しく貼り付けられているか確認し、必要に応じてエラーハンドリングを強化する。

### #5 月次アポ帳作成フォームのUIを現場向けに刷新する

受付・歯科助手が初見でも迷わないように、フォームを「月次作成ウィザード」に近い構成へ整理済み。

## 暫定運用

フォームや既存ボタンから作成して罫線・改ページが反映されない場合は、`Alt + F8` から以下を実行する。

```vba
GenerateBook_Phase5_Print
```

`Output` を再生成し、網掛け開始セルの罫線も含めて整える場合は以下を実行する。

```vba
GenerateBook_Phase5_Print
```

`Output` を再生成し、行高・改ページ・印刷範囲まで整えたい場合は以下を実行する。

```vba
GenerateBook_Phase5_Print
```
