# VBA modules

このフォルダには、ClinicAppointmentで使用するExcel VBAモジュールを格納します。

## 現在のモジュール

- `AppointmentBook_Phase1.bas`

## インポート方法

1. Excelを `.xlsm` 形式で保存する。
2. `Alt + F11` でVBAエディタを開く。
3. 古いモジュールがある場合は右クリックして削除する。
4. `ファイル` → `ファイルのインポート` から `.bas` ファイルを選択する。
5. `Alt + F8` で `GenerateAppointmentBook_Phase4` を実行する。

## 必要シート

- `Template`
- `Settings`
- `Output`

## Settings

- `B2`: 年
- `B3`: 月
- `B5:F5`: 列見出し。

列見出しの対応は以下の通り。

- `B5` → Output列B：1列目Dr
- `C5` → Output列D：2列目Dr
- `D5` → Output列F：予備枠
- `E5` → Output列I：1列目DH
- `F5` → Output列J：2列目DH

TemplateのH列は意図的に幅が狭いスペーサー列として扱うため、担当者名の出力先には使用しない。

`B5:F5` がすべて空欄の場合は、Templateに入っている列見出しをそのまま使用する。
