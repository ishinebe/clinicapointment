# ClinicAppointment

ClinicAppointment は、クリニックで使用する紙運用前提のアポイント帳を Excel VBA で自動生成するプロジェクトです。

## 目的

現在の紙アポ帳運用を維持したまま、Excelによるアポ帳作成を自動化し、毎月の作成時間と人的ミスを削減します。

## 基本方針

- `Template` シートをデザインマスターとして扱う
- `Settings` シートで年月、担当者、通常診療時間などを管理する
- `Exceptions` シートで会議、早退、休診などの例外を管理する
- `Output` シートへ印刷用アポ帳を自動生成する
- 現行アポ帳の Dr 列、予備枠、DH 列の幅や配置は Template を踏襲する
- 時間列のみ、視認性を優先して `9:00`, `9:15`, `9:30`, `9:45` の形式でマクロが描画する

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

Phase 1完了。

`Template!A1:J46` を1日分テンプレートとして固定し、`Settings!B2` の年、`Settings!B3` の月に基づいて、1か月分のアポ帳を `Output` シートへ生成できることを確認済みです。

次は、生成結果を現行アポ帳の実用書式に近づける工程に進みます。特に、時間列以外の見た目は現行ファイルを踏襲し、時間列のみ視認性を改善する方針です。
