# ClinicAppointment

ClinicAppointment は、クリニックで使用する紙運用前提のアポイント帳を Excel VBA で自動生成するプロジェクトです。

## 目的

現在の紙アポ帳運用を維持したまま、Excelによるアポ帳作成を自動化し、毎月の作成時間と人的ミスを削減します。

## 基本方針

- `Template` シートを確定済みデザインマスターとして扱う
- `Settings` シートを現場向け操作画面として整備し、年月、担当者、通常診療時間などを管理する
- `Exceptions` シートを内部データ保存先として使い、ユーザー向けには「臨時予定」として管理する
- `Output` シートへ印刷用アポ帳を自動生成する
- 現行アポ帳の Dr 列、予備枠、DH 列の幅や配置は Template を踏襲する
- 時間列は、視認性を優先した `9:00`, `9:15`, `9:30`, `9:45` の形式を採用する
- 印刷時は `1日 = 1ページ` を原則とし、Excelの自動改ページに任せず、生成後に日別ブロック単位で手動改ページを設定する
- 休診・時短・スタッフ休みなどの予約不可範囲は、網掛けに加えて開始位置の上罫線で視認性を補助する

## 初期構成

```text
clinicapointment/
├─ README.md
├─ docs/
│  ├─ design-principles-v1.3.md
│  ├─ phase-plan.md
│  └─ current-work-items.md
├─ vba/
│  ├─ README.md
│  ├─ AppointmentBook_Phase1.bas
│  ├─ PrintPageSetup.bas
│  ├─ ExceptionsDateDropdown.bas
│  ├─ ThisWorkbook.cls
│  ├─ frmAppointmentSettings.frm
│  └─ frmTemporarySchedule.frm
├─ samples/
│  └─ README.md
└─ .gitignore
```

## 現在の実装段階

Phase 8として、月次アポ帳作成フォームに加えて臨時予定フォームを追加しています。

`Template!A1:J46` を1日分テンプレートとして固定し、`Settings!B2` の年、`Settings!B3` の月に基づいて、1か月分のアポ帳を `Output` シートへ生成できることを確認済みです。

現行アポ帳に合わせるための書式調整は、テンプレート作成作業時に完了済みです。今後は、確定済みTemplateの見た目を崩さず、日付・曜日・担当者・曜日別診療時間・臨時予定などを安全に反映する方向で進めます。

現在は、受付・歯科助手でも迷わず操作できるように、`frmAppointmentSettings` を月次アポ帳作成フォームとして、`frmTemporarySchedule` を臨時予定フォームとして追加しています。UserFormは `Settings` シートを置き換えるものではなく、既存セルへ入力する補助画面です。既存ロジックとの互換性維持のため、内部参照セルである `B2`、`B3`、`B5:F5`、`B7:F13`、`B16:F16`、`H5:H24` は変更しません。

印刷設定については、`vba/PrintPageSetup.bas` に `ApplyOneDayOnePagePrintSettings` と `GenerateAppointmentBook_Phase5_WithOneDayPrintSettings` を追加しています。これにより、行高の再正規化、印刷範囲設定、横1ページ固定、日別ブロックごとの手動改ページ、網掛け開始セルの上罫線追加をまとめて実行できます。

一方で、フォームや `Settings` シート上の既存ボタンがまだ従来の `GenerateAppointmentBook_Phase5` を呼んでいる場合、印刷設定・罫線処理が適用されません。今後はフォーム側とボタン側の呼び出し先を `GenerateAppointmentBook_Phase5_WithOneDayPrintSettings` に統一します。

`.frm` の直接インポートで文字コードや形式の問題が出る場合は、VBE上で空のUserFormを作成し、`frmAppointmentSettings.frm` の `Option Explicit` 以降をコード画面へ貼り付けて使用します。

Excelファイルを開いたときに月次アポ帳作成フォームを自動表示するには、初回にExcel上部の警告バーから「コンテンツの有効化」を押してマクロを有効にします。マクロが無効な状態では自動表示されません。実ファイル側では `ThisWorkbook.cls` の `Workbook_Open` が正しく貼り付けられているか確認が必要です。

## 通常の使用手順

1. 初回はExcelの「コンテンツの有効化」でマクロを有効にする。
2. 初回設定として `SetupUserFriendlySettings` または `SetupSettingsDropdowns` を実行する。
3. 次回以降はファイルを開くと月次アポ帳作成フォームが自動で開く。
4. フォームを閉じた後は、`ShowAppointmentSettingsForm` または `Settings` の「設定フォームを開く」ボタンで再度開ける。
5. フォームで年月、担当者、曜日別勤務パターン、医院全体の終了時刻を設定する。
6. 必要ならフォームの「臨時予定を編集」から休診・早上がりなどを入力する。
7. フォームの「アポ帳を作成」ボタンを押す。
8. `Output` シートを改ページプレビューで確認し、1日分が1ページに収まっているか確認する。

現時点では、フォームの「アポ帳を作成」ボタンから実行した場合に印刷設定・罫線処理が適用されない可能性があります。その場合は、暫定的に `Alt + F8` から `GenerateAppointmentBook_Phase5_WithOneDayPrintSettings` を実行してください。

従来通り、`Settings` シートを直接編集して既存ボタンから作成する運用も残しています。ただし、既存ボタンの呼び出し先が `GenerateAppointmentBook_Phase5` のままの場合は、印刷設定・罫線処理が反映されません。

## 印刷・改ページ・網掛け境界

`vba/PrintPageSetup.bas` では、以下の補助マクロを提供しています。

```text
ApplyOneDayOnePagePrintSettings
GenerateAppointmentBook_Phase5_WithOneDayPrintSettings
ApplyShadeStartTopBordersToOutput
```

`ApplyOneDayOnePagePrintSettings` は、生成済みの `Output` に対して以下を行います。

```text
1. Template由来の行高を各日ブロックへ再適用する
2. 日別ブロック間の空白行高を固定する
3. 網掛け開始セルの上罫線を太くする
4. 印刷範囲を A1:J最終行 に設定する
5. 横方向を1ページに収める
6. 各日ブロックの開始行に手動改ページを入れる
```

`GenerateAppointmentBook_Phase5_WithOneDayPrintSettings` は、アポ帳生成と印刷設定をまとめて実行する入口です。今後はフォーム・Settingsボタンとも、このマクロを呼ぶように統一します。

`ApplyShadeStartTopBordersToOutput` は、すでに `Output` が生成済みの場合に、網掛け開始セルの上罫線だけを後から追加するための補助マクロです。

## 現在の主な未対応事項

現在の微修正・改善候補は GitHub Issue に整理しています。

- #2 網掛け開始セルの上罫線をさらに太くして視認性を上げる
- #3 フォームからアポ帳作成した場合も印刷設定・罫線処理を適用する
- #4 Excel起動時に月次アポ帳作成フォームが確実に開くようにする
- #5 月次アポ帳作成フォームのUIを現場向けに刷新する

## UserFormの作成手順

`frmAppointmentSettings.frm` の直接インポートに失敗する場合は、VBE上でUserFormを手作成します。

1. VBEで「挿入」→「ユーザーフォーム」を選択する。
2. 作成されたUserFormのオブジェクト名を `frmAppointmentSettings` に変更する。
3. GitHubの `vba/frmAppointmentSettings.frm` を開く。
4. `Option Explicit` 以降のVBAコード部分をコピーする。
5. VBE上の `frmAppointmentSettings` のコード画面へ貼り付ける。
6. `AppointmentBook_Phase1.bas` 側に `ShowAppointmentSettingsForm` があることを確認する。
7. `Alt + F8` で `ShowAppointmentSettingsForm` を実行し、フォームが開くことを確認する。

`.frm` 先頭の `VERSION` / `Begin VB.UserForm` / `Attribute` 行は貼り付けません。フォーム部品はコードで動的生成するため、手配置は不要です。

`frmTemporarySchedule.frm` も同じ手順で作成できます。UserFormのオブジェクト名を `frmTemporarySchedule` に変更し、`vba/frmTemporarySchedule.frm` の `Option Explicit` 以降を貼り付けます。

起動時にフォームを自動表示する場合は、VBEの `ThisWorkbook` コード画面へ `vba/ThisWorkbook.cls` の `Option Explicit` 以降を貼り付けます。`Workbook_Open` が `ShowAppointmentSettingsForm` を呼び出します。
