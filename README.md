# ClinicAppointment

ClinicAppointment は、クリニックで使用する紙運用前提のアポイント帳を Excel VBA で自動生成するプロジェクトです。

## 目的

現在の紙アポ帳運用を維持したまま、Excelによるアポ帳作成を自動化し、毎月の作成時間と人的ミスを削減します。

## 基本方針

- `Template` シートを確定済みデザインマスターとして扱う
- `Settings` シートを現場向け操作画面として整備し、年月、担当者、通常診療時間などを管理する
- `Exceptions` シートで会議、早退、休診などの例外を管理する
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

Settings UI改善フェーズ。

`Template!A1:J46` を1日分テンプレートとして固定し、`Settings!B2` の年、`Settings!B3` の月に基づいて、1か月分のアポ帳を `Output` シートへ生成できることを確認済みです。

現行アポ帳に合わせるための書式調整は、テンプレート作成作業時に完了済みです。今後は、確定済みTemplateの見た目を崩さず、日付・曜日・担当者・曜日別診療時間・例外設定などを安全に反映する方向で進めます。

現在は、受付・歯科助手でも迷わず操作できるように、`Settings` シートを「アポ帳作成 設定画面」として整備しています。既存ロジックとの互換性維持のため、内部参照セルである `B2`、`B3`、`B5:F5`、`B7:F13`、`B16:F16`、`H5:H24` は変更しません。

## 通常の使用手順

1. `SetupUserFriendlySettings` または `SetupSettingsDropdowns` を実行する。
2. `Settings` で年・月を入力する。
3. 担当者をプルダウンから選択する。
4. 必要なら曜日別勤務パターンを設定する。
5. 必要なら `Exceptions` に日付別例外を入力する。
6. `Settings` の「アポ帳を作成」ボタンを押す。
