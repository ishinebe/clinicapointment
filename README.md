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

`.frm` の直接インポートで文字コードや形式の問題が出る場合は、VBE上で空のUserFormを作成し、`frmAppointmentSettings.frm` の `Option Explicit` 以降をコード画面へ貼り付けて使用します。

## 通常の使用手順

1. 初回は `SetupUserFriendlySettings` または `SetupSettingsDropdowns` を実行する。
2. `ShowAppointmentSettingsForm` または `Settings` の「設定フォームを開く」ボタンでフォームを開く。
3. フォームで年月、担当者、曜日別勤務パターン、医院全体の終了時刻を設定する。
4. 必要ならフォームの「臨時予定を編集」から休診・早上がりなどを入力する。
5. フォームの「アポ帳を作成」ボタンを押す。

従来通り、`Settings` シートを直接編集して既存ボタンから作成する運用も残しています。

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
