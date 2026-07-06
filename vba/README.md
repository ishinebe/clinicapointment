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
- `B5:F5`: 列見出し。左から順に、1列目Dr、2列目Dr、予備枠、1列目DH、2列目DHとして扱う。

`B5:F5` がすべて空欄の場合は、Templateに入っている列見出しをそのまま使用する。
