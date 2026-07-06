# Codex依頼プロンプト: TemplateDraft作成フェーズ

## 背景

ClinicAppointment は、紙運用前提のクリニック用アポ帳を Excel VBA で自動生成するプロジェクトです。

現在の設計方針は以下です。

- `Template` はデザインマスターとして扱う。
- `Output` はTemplateをもとに生成する印刷用シート。
- 現行アポ帳のDr列、予備枠、DH列の列幅・配置・セル結合・罫線はTemplateを正とする。
- 一方で、現行Templateの時間表示は旧様式で視認性が低いため、時間列だけはマクロで改善する。
- Template本体を直接壊さないため、まず `TemplateDraft` シートを作成し、そこで改良版テンプレートを試作する。

## 依頼内容

Excel VBAで `CreateTemplateDraft` マクロを実装してください。

## 追加・修正してほしいファイル

- `vba/AppointmentBook_Phase1.bas`

必要に応じて関数を追加・分割してください。

## 実装要件

### 1. TemplateDraftシートの作成

- `Template` シートをコピーして `TemplateDraft` シートを作成する。
- 既に `TemplateDraft` が存在する場合は、確認メッセージを出した上で削除・再作成する。
- `Template` シート本体は絶対に変更しない。

### 2. レイアウト維持

`TemplateDraft` では、以下をTemplateから踏襲してください。

- Dr列の列幅
- 予備枠の列幅
- DH列の列幅
- セル結合
- 罫線
- フォント
- 行高
- 印刷設定

### 3. 時間列の改善

現行Templateでは、左側に「9, 10, 11...」、別列に「00, 15, 30, 45」のように分かれている場合があります。

`TemplateDraft` では、時間表示を以下のように15分刻みで見やすくしてください。

```text
9:00
9:15
9:30
9:45
10:00
10:15
...
```

要件：

- 開始時刻は 9:00。
- 終了表示は少なくとも 19:00 まで対応する。
- 15分刻み。
- 00分は太実線または実線で視認性を高める。
- 15分・30分・45分は薄い点線にする。
- 時間列以外のDr列、予備枠、DH列の幅や配置はTemplateを踏襲する。

### 4. 時間外網掛けの試作

`TemplateDraft` 上で、視認性確認用に時間外網掛けを適用できるようにしてください。

初期値：

- 月火金想定：19:00以降
- 水曜想定：18:00以降
- 土曜想定：17:30以降

ただし、TemplateDraftは特定曜日ではなく「1日分テンプレート」なので、まずは任意の終了時刻をコード内定数として持たせ、例として `18:00` 以降を薄い網掛けにしてください。

網掛けはボールペンで記入できる薄さにしてください。

### 5. 実行マクロ

以下のマクロを用意してください。

```vb
Public Sub CreateTemplateDraft()
```

既存のPhase1マクロは残してください。

### 6. 安全性

- `Application.ScreenUpdating`, `Application.DisplayAlerts`, `Application.EnableEvents` は適切に制御し、エラー時にも復元してください。
- `Option Explicit` を維持してください。
- 既存の `GenerateAppointmentBook_Phase1` を壊さないでください。

## 完了条件

- `CreateTemplateDraft` を実行すると `TemplateDraft` シートが作成される。
- Template本体は変更されない。
- Dr列、予備枠、DH列の見た目はTemplateと同等に維持される。
- 時間列が `9:00`, `9:15`, `9:30`, `9:45` 形式になる。
- 00分と15分刻みの罫線が視認性よく設定される。
- 薄い網掛けが適用できる。

## 補足

現時点では、完璧な最終Output生成よりも、TemplateDraft上で見た目を検証できることを優先してください。
