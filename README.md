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
- 印刷時は `1日 = 1ページ` を原則とし、Excelの自動改ページに任せず、生成後に日別ブロック単位で手動改ページを設定する方針とする

## 初期構成

```text
clinicapointment/
├─ README.md
├─ docs/
│  ├─ design-principles-v1.3.md
│  └─ phase-plan.md
├─ vba/
│  ├─ README.md
│  └─ AppointmentBook_Phase1.bas
├─ samples/
│  └─ README.md
└─ .gitignore
```

## 現在の実装段階

Phase 8として、月次アポ帳作成フォームに加えて臨時予定フォームを追加しています。

`Template!A1:J46` を1日分テンプレートとして固定し、`Settings!B2` の年、`Settings!B3` の月に基づいて、1か月分のアポ帳を `Output` シートへ生成できることを確認済みです。

現行アポ帳に合わせるための書式調整は、テンプレート作成作業時に完了済みです。今後は、確定済みTemplateの見た目を崩さず、日付・曜日・担当者・曜日別診療時間・臨時予定などを安全に反映する方向で進めます。

現在は、受付・歯科助手でも迷わず操作できるように、`frmAppointmentSettings` を月次アポ帳作成フォームとして、`frmTemporarySchedule` を臨時予定フォームとして追加しています。UserFormは `Settings` シートを置き換えるものではなく、既存セルへ入力する補助画面です。既存ロジックとの互換性維持のため、内部参照セルである `B2`、`B3`、`B5:F5`、`B7:F13`、`B16:F16`、`H5:H24` は変更しません。

印刷設定については、現時点ではExcelの自動改ページにより日別ブロックの途中でページが分割される場合があります。実運用では `1日 = 1ページ` が必要なため、今後の優先作業として、生成時に `Output` の印刷範囲・横1ページ設定・日別ブロック末尾の手動改ページを自動設定する必要があります。

`.frm` の直接インポートで文字コードや形式の問題が出る場合は、VBE上で空のUserFormを作成し、`frmAppointmentSettings.frm` の `Option Explicit` 以降をコード画面へ貼り付けて使用します。

Excelファイルを開いたときに月次アポ帳作成フォームを自動表示するには、初回にExcel上部の警告バーから「コンテンツの有効化」を押してマクロを有効にします。マクロが無効な状態では自動表示されません。

## 通常の使用手順

1. 初回はExcelの「コンテンツの有効化」でマクロを有効にする。
2. 初回設定として `SetupUserFriendlySettings` または `SetupSettingsDropdowns` を実行する。
3. 次回以降はファイルを開くと月次アポ帳作成フォームが自動で開く。
4. フォームを閉じた後は、`ShowAppointmentSettingsForm` または `Settings` の「設定フォームを開く」ボタンで再度開ける。
5. フォームで年月、担当者、曜日別勤務パターン、医院全体の終了時刻を設定する。
6. 必要ならフォームの「臨時予定を編集」から休診・早上がりなどを入力する。
7. フォームの「アポ帳を作成」ボタンを押す。
8. `Output` シートを改ページプレビューで確認し、1日分が1ページに収まっているか確認する。現時点では必要に応じて手動で改ページを修正する。

従来通り、`Settings` シートを直接編集して既存ボタンから作成する運用も残しています。

## 印刷・改ページの暫定対応

現時点で1日分の途中に青い点線の改ページが入る場合は、Excelの自動改ページが原因です。

暫定的には以下の手順で修正します。

1. `Output` シートを開く。
2. `表示` → `改ページプレビュー` を選択する。
3. 青い改ページ線を、各日ブロックの最後の行の直後へ移動する。
4. 1日分が1ページに収まることを確認してから印刷する。

恒久対応としては、VBA側で `ResetAllPageBreaks` を実行した後、各日ブロック末尾に `HPageBreaks.Add` または `Rows(...).PageBreak = xlPageBreakManual` を設定し、`PrintArea` と横1ページ設定を同時に適用する方針です。

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
