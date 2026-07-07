# VBA modules

このフォルダには、ClinicAppointmentで使用するExcel VBAモジュールを格納します。

## 現在のモジュール

- `AppointmentBook_Phase1.bas`
- `ExceptionsDateDropdown.bas`
- `ThisWorkbook.cls`
- `frmAppointmentSettings.frm`
- `frmTemporarySchedule.frm`

## インポート方法

1. Excelを `.xlsm` 形式で保存する。
2. `Alt + F11` でVBAエディタを開く。
3. 古いモジュールがある場合は右クリックして削除する。
4. `ファイル` → `ファイルのインポート` から `.bas` ファイルを選択する。
5. `ThisWorkbook.cls` はインポートではなく、VBEの `ThisWorkbook` コード画面へ `Option Explicit` 以降を貼り付ける。
6. `frmAppointmentSettings.frm` は、直接インポートできる場合はインポートする。
7. `.frm` インポートに失敗する場合は、後述の「UserFormを手作成してコードを貼り付ける手順」で作成する。
8. 初回設定として `Alt + F8` で `SetupUserFriendlySettings` または `SetupSettingsDropdowns` を実行する。
9. 初回はExcel上部の警告バーから「コンテンツの有効化」を押してマクロを有効にする。マクロが無効な状態では自動表示されない。
10. ファイルを開くと月次アポ帳作成フォームが自動で開く。閉じた後は、`Alt + F8` で `ShowAppointmentSettingsForm` を実行するか、`Settings` の「設定フォームを開く」ボタンを押す。
11. アポ帳生成時はフォームの「アポ帳を作成」ボタン、または従来の `Settings` の「アポ帳を作成」ボタンを押す。

## 実運用フロー

受付・事務担当者が月次のアポ帳を作成する場合は、以下の順番で操作する。

```text
1. 初回は Excel の「コンテンツの有効化」でマクロを有効にする。
2. 初回設定として SetupUserFriendlySettings または SetupSettingsDropdowns を実行する。
3. 次回以降はファイルを開くと月次アポ帳作成フォームが自動で開く。
4. フォームを閉じた後は、ShowAppointmentSettingsForm または Settings の「設定フォームを開く」ボタンで再度開く。
5. フォームで年・月、担当者、曜日別勤務、医院全体の当月原則終了時刻を設定する。
6. 必要なら「臨時予定を編集」フォームで特定日だけの休診・早上がり・スタッフ休みを設定する。
7. フォームの「アポ帳を作成」ボタンを押す。
8. Output の見た目、網掛け、改ページ、印刷範囲を確認する。
```

現時点では、改ページと印刷範囲の完全自動設定は未完了である。1日分の途中で青い点線の改ページが入る場合は、Excelの自動改ページが原因なので、改ページプレビューで各日ブロック末尾へ手動調整してから印刷する。

今後、入力チェックマクロを追加した場合は、`GenerateAppointmentBook_Phase5` の前に入力チェックを実行する。

## 設計上の役割分担

### Template

`Template!A1:J46` は、1日分の完成済みデザインマスターである。

時間列、罫線、列幅、紙面上の見た目は、基本的に `Template` を正とする。

マクロ側では、Templateの見た目を大きく再構成しない。

### Settings

`Settings` は、その月全体に繰り返し適用する設定を扱う。

UserForm化後も、`Settings` は内部設定保存先として残す。`frmAppointmentSettings` は `Settings` セルへの入力補助画面であり、既存セル番地や生成ロジックを置き換えない。

```text
Settings = 毎週繰り返す基本パターン + 医院全体の当月原則終了時刻
```

例：

```text
毎週水曜はDH1が16:00まで
今月だけ医院全体が17:00終了
```

### 臨時予定

ユーザー向けには「臨時予定」と呼ぶ。

内部データ保存先として `Exceptions` シートを使う。シート名とA:E列の既存構造は変更しない。

```text
臨時予定 = 特定日の休診・時短診療・スタッフ休み・スタッフ時短勤務
```

例：

```text
2/10だけ学会で休診
2/17だけ院内研修で16:00終了
2/19だけDr2が16:00まで
```

## 設定の優先順位

同じ日に複数の設定が関係する場合、考え方としては以下を優先する。

```text
1. 臨時予定の休診
2. 臨時予定の時短診療
3. 臨時予定のスタッフ休み
4. 臨時予定のスタッフ時短勤務
5. Settings の曜日別勤務パターン
6. Settings の医院全体の当月原則終了時刻
7. 通常の曜日別診療時間
```

医院全体の休診・時短診療は、スタッフ個別設定より強い設定として扱う。

## 必要シート

- `Template`
- `Settings`
- `Output`
- `Exceptions`

## Settings

以下のセル番地は既存生成ロジックとの互換性のため固定する。画面上の見た目を整えても、内部参照セルは変更しない。

- `B2`: 年
- `B3`: 月
- `B5:F5`: 列見出し。担当者マスターからプルダウン選択する。
- `B7:F13`: 曜日別の勤務パターン。1セル完結型プルダウン。
- `B16:F16`: 医院全体の当月原則終了時刻。通常は空欄。医院全体の診療終了時刻をその月だけ変更する場合に選択する。
- `A7:A13`: 曜日ラベル。
- `H5:H24`: 担当者マスター。

## 列対応

- `B5` / `B7:B13` → Output列B：1列目Dr
- `C5` / `C7:C13` → Output列D：2列目Dr
- `D5` / `D7:D13` → Output列F：予備枠
- `E5` / `E7:E13` → Output列I：1列目DH
- `F5` / `F7:F13` → Output列J：2列目DH

TemplateのH列は幅が狭いスペーサー列として扱うため、担当者名の出力先には使用しない。

## プルダウン設定

`SetupSettingsDropdowns` を実行すると、以下が設定される。

- `Settings` シートを「アポ帳作成 設定画面」として整える。
- `A7:A13` に曜日ラベルを配置する。
- `B5:F5` に担当者マスター参照のプルダウンを設定する。
- `B7:F13` に勤務パターンのプルダウンを設定する。
- `B16:F16` を結合し、医院全体の終了時刻プルダウンを設定する。
- `Exceptions` シートを作成し、臨時予定の種別・対象・終了時刻プルダウンを設定する。
- `Settings` シートに「アポ帳を作成」ボタンと「日付候補を更新」ボタンを配置する。
- `Settings` シートに `ShowTemporaryScheduleForm` を実行する「臨時予定を編集」ボタンを配置する。
- `Settings` シートに `ShowAppointmentSettingsForm` を実行する「設定フォームを開く」ボタンを配置する。
- 再実行時は既存の同名ボタンを削除してから作成するため、ボタンは重複しない。

`SetupUserFriendlySettings` は、現場向けSettings画面の整備、プルダウン設定、ボタン配置をまとめて実行する入口である。内部的には `SetupSettingsDropdowns` と同じ整備を行う。

`CreateAppointmentBookButton` は、`Settings` シートに `GenerateAppointmentBook_Phase5` を実行する「アポ帳を作成」ボタンを配置する。

`CreateRefreshExceptionDatesButton` は、`Settings` シートに `SetupExceptionsDateDropdowns` を実行する「日付候補を更新」ボタンを配置する。

`SetupExceptionsDateDropdowns` を実行すると、`Settings!B2/B3` の年月をもとに、`Exceptions!A2:A100` に当月日付のプルダウンを設定する。

## UserForm

`ShowAppointmentSettingsForm` を実行すると、`frmAppointmentSettings` が開く。

`ThisWorkbook.Workbook_Open` でも `ShowAppointmentSettingsForm` を呼び出すため、マクロ有効状態でExcelファイルを開くと月次アポ帳作成フォームが自動で開く。フォームを閉じても、`Settings` シートの「設定フォームを開く」ボタンまたは `Alt + F8` の `ShowAppointmentSettingsForm` から再度開ける。

`frmAppointmentSettings.frm` には日本語UI文字列を含める。`.frm` の直接インポートで失敗する場合は、VBE上で空のUserFormを作成し、コード部分だけを貼り付ける。

フォームは起動時に `Settings` の既存値を読み込む。

- 年: `Settings!B2`
- 月: `Settings!B3`
- 担当者: `Settings!B5:F5`
- 曜日別勤務パターン: `Settings!B7:F13`
- 医院全体の当月原則終了時刻: `Settings!B16`
- 担当者候補: `Settings!H5:H24`

フォームの「設定を保存」は、入力値を `Settings` へ書き戻す。

フォームの「アポ帳を作成」は、入力値を `Settings` へ保存してから `GenerateAppointmentBook_Phase5` を実行する。

フォームの「臨時予定を編集」は、フォーム上の年月を `Settings` へ保存してから `SetupExceptionsDateDropdowns` を実行し、`frmTemporarySchedule` を開く。

`ShowTemporaryScheduleForm` を実行すると、`frmTemporarySchedule` が開く。

臨時予定フォームでは、内部保存先である `Exceptions!A2:E100` を一覧表示し、追加・更新・削除できる。

- 日付: `Exceptions!A列`
- 内容: `Exceptions!B列`
- 対象: `Exceptions!C列`
- 終了時刻: `Exceptions!D列`
- メモ: `Exceptions!E列`

画面上では「医院全体を休みにする」「医院全体を早く閉める」「スタッフが休み」「スタッフが早上がり」と表示し、保存時は既存生成ロジックと互換の `休診` / `時短診療` / `休み` / `時短勤務` に変換する。

既存の `Settings` シート直接入力と既存ボタンも引き続き使用できる。

### UserFormを手作成してコードを貼り付ける手順

`.frm` ファイルをVBEへ直接インポートできない場合は、以下の手順で作成する。

1. VBEで「挿入」→「ユーザーフォーム」を選択する。
2. 作成されたUserFormのオブジェクト名を `frmAppointmentSettings` に変更する。
3. GitHubの `vba/frmAppointmentSettings.frm` を開く。
4. `Option Explicit` 以降のVBAコード部分をコピーする。
5. VBE上の `frmAppointmentSettings` のコード画面へ貼り付ける。
6. `AppointmentBook_Phase1.bas` 側に `ShowAppointmentSettingsForm` があることを確認する。
7. `Alt + F8` で `ShowAppointmentSettingsForm` を実行し、フォームが開くことを確認する。

`.frm` の先頭にある `VERSION` / `Begin VB.UserForm` / `Attribute` 行は、コード貼り付け時には貼り付けない。貼り付けるのは `Option Explicit` 以降のコード部分のみ。

フォーム上の部品はコードで動的生成しているため、VBE上でLabelやComboBoxなどを手配置する必要はない。

`frmTemporarySchedule` も同じ手順で作成する。UserFormのオブジェクト名を `frmTemporarySchedule` に変更し、`vba/frmTemporarySchedule.frm` の `Option Explicit` 以降を貼り付ける。

### ThisWorkbookのコード貼り付け手順

起動時にフォームを自動表示する場合は、以下の手順で `Workbook_Open` を追加する。

1. VBEのプロジェクトツリーで `ThisWorkbook` を開く。
2. GitHubの `vba/ThisWorkbook.cls` を開く。
3. `Option Explicit` 以降のVBAコード部分をコピーする。
4. VBE上の `ThisWorkbook` コード画面へ貼り付ける。
5. Excelファイルを `.xlsm` として保存し、開き直す。
6. Excelの警告バーで「コンテンツの有効化」を押す。
7. 月次アポ帳作成フォームが自動で開くことを確認する。

## 曜日別勤務パターン

通常の出勤日は空欄のままにする。

`B7:F13` のプルダウンは以下。

```text
休
午前のみ
15:00まで
15:30まで
16:00まで
16:30まで
17:00まで
17:30まで
18:00まで
18:30まで
```

例：2列目Drが水曜日だけ16:00あがりの場合。

```text
C9 = 16:00まで
```

この場合、その月の水曜日だけ、2列目Drの列が16:00以降網掛けされる。担当者ヘッダーには `16:00まで` などの時刻は追記しない。

`午前のみ` は現時点では12:00以降を網掛けする。

## 医院全体の当月原則終了時刻

`B16:F16` は、その月だけ医院全体の診療終了時刻を一括で変更するための設定欄。

通常は空欄のままにする。空欄の場合は、曜日別の通常診療時間を使う。

```text
B16:F16 = 17:00
```

この場合、その月の診療日は原則として17:00以降を全体網掛けする。Dr1、Dr2、予備、DH1、DH2の全列に適用される。

## 臨時予定の内部保存先

臨時予定は、内部的には `Exceptions` シートのA:E列に保存する。

列構成は以下。

```text
A列: 日付
B列: 種別
C列: 対象
D列: 終了時刻
E列: メモ
```

`A列: 日付` は、`SetupExceptionsDateDropdowns` を実行すると、当月の日付からプルダウン選択できる。

種別は以下。

```text
休診
時短診療
休み
時短勤務
```

対象は以下。

```text
全体
Dr1
Dr2
予備
DH1
DH2
```

## 種別の使い分け

### 休診

医院全体が休みの日に使う。

```text
A2 = 2027/2/10
B2 = 休診
C2 = 全体
E2 = 学会
```

その日のブロック全体を網掛けし、休診表示を入れる。

### 時短診療

特定日だけ医院全体の終了時刻が早い場合に使う。

```text
A2 = 2027/2/17
B2 = 時短診療
C2 = 全体
D2 = 16:00
E2 = 院内研修
```

その日の全列で16:00以降を網掛けする。

### 休み

特定日だけ特定スタッフが休みの場合に使う。

```text
A2 = 2027/2/12
B2 = 休み
C2 = DH1
E2 = 有休
```

その日のDH1列だけを終日網掛けする。担当者ヘッダーへの `休` 追記は行わない。

### 時短勤務

特定日だけ特定スタッフが早上がりする場合に使う。

```text
A2 = 2027/2/19
B2 = 時短勤務
C2 = Dr2
D2 = 16:00
E2 = 早退
```

その日のDr2列だけを16:00以降網掛けする。担当者ヘッダーへの `16:00まで` 追記は行わない。

## 入力ルール

- 医院全体に関する臨時予定は、種別を `休診` または `時短診療`、対象を `全体` にする。
- スタッフ個別の臨時予定は、種別を `休み` または `時短勤務`、対象を `Dr1` / `Dr2` / `予備` / `DH1` / `DH2` から選ぶ。
- `時短診療` と `時短勤務` の場合だけ、終了時刻を選択する。
- `休診` と `休み` の場合、終了時刻は空欄でよい。

## 印刷・改ページ要件

紙運用では、`1日 = 1ページ` を原則とする。

現在の `Output` は `Template!A1:J46` を1日分として、各日ブロックの間に `BLOCK_GAP_ROWS = 2` 行の空白を入れて縦方向に連結している。そのため、日別ブロックの高さは以下として扱う。

```text
1日分のTemplate行数 = 46行
日別ブロック間の空白 = 2行
次の日の開始行 = 前日の開始行 + 48行
```

印刷設定の恒久対応では、生成完了後に以下を自動適用する。

```text
1. Output の既存改ページをリセットする
2. 印刷範囲を A1:J最終行 に固定する
3. 横方向は1ページに収める
4. 縦方向は自動、ただし各日ブロック末尾で手動改ページする
5. 各日ブロックの最後、つまり次の日ブロック開始行の直前に改ページを入れる
```

実装候補の考え方は以下。

```vba
wsO.ResetAllPageBreaks
wsO.PageSetup.PrintArea = wsO.Range("A1:J" & lastOutputRow).Address
wsO.PageSetup.Zoom = False
wsO.PageSetup.FitToPagesWide = 1
wsO.PageSetup.FitToPagesTall = False

For dayIndex = 1 To daysInMonth - 1
    nextDayStartRow = 1 + dayIndex * (TEMPLATE_ROW_COUNT + BLOCK_GAP_ROWS)
    wsO.HPageBreaks.Add Before:=wsO.Rows(nextDayStartRow)
Next dayIndex
```

この処理は、網掛け・担当者名・臨時予定の反映が完了した後、`GenerateAppointmentBookCore` の最後で実行するのが望ましい。

暫定対応として、手動で確認する場合は以下の手順とする。

```text
1. Output を開く
2. 表示 → 改ページプレビュー を選択する
3. 1日分の途中にある青い点線を、次の日ブロック開始直前へ移動する
4. 1日分が1ページに収まることを確認する
5. 印刷する
```

## 現在実装済みの機能

- `Template!A1:J46` を1日分として固定コピーする。
- `UsedRange` を使わず、不要な巨大コピーやページ増加を避ける。
- 1か月分の日別ブロックを `Output` に生成する。
- `Settings!B2/B3` の年月から対象月の日数を判定する。
- 日付・曜日を各日ブロックに反映する。
- Templateの時間列・罫線・列幅・基本デザインを維持する。
- 曜日別の通常診療時間を反映する。
- 土曜17:30終了、水曜18:00終了、月火金19:00終了、木日休診に対応する。
- A列の実際の時間表示を読んで網掛けを判定する。
- `Settings!B5:F5` の担当者名をOutputの見出しに反映する。
- H列をスペーサーとして扱い、DH1/DH2はI列/J列に出力する。
- `Settings!B7:F13` の1セル完結型勤務パターンに対応する。
- `Settings!B16:F16` の医院全体の当月原則終了時刻に対応する。
- 臨時予定として保存した休診・時短診療・休み・時短勤務に対応する。
- `Exceptions!A2:A100` の当月日付プルダウンに対応する。

## 今後の作業候補

### 1. 印刷・改ページの安定化

`Output` 生成後に、`1日 = 1ページ` となるように印刷設定を自動適用する。

対応内容：

```text
ResetAllPageBreaks
PrintArea = A1:J最終行
FitToPagesWide = 1
FitToPagesTall = False
各日ブロック末尾に手動改ページ
```

現在確認されている問題：

```text
Excelの自動改ページにより、夕方の時間帯など日別ブロックの途中でページが切れる場合がある。
次の日付ブロックが同一ページ下部に入り、紙アポ帳として使いづらい。
```

完了条件：

```text
各日が必ず1ページに収まる
次の日付が同じページへ入り込まない
印刷プレビューで月の日数分のページ数になる
横方向が1ページに収まる
```

### 2. 入力ミス検出マクロ

生成前に `Settings` と臨時予定の入力内容を確認する。

候補マクロ名：

```text
ValidateAppointmentBookSettings
```

検出したい例：

```text
時短診療なのに終了時刻が空欄
時短勤務なのに終了時刻が空欄
休診なのに対象が全体以外
休みなのに対象が全体
日付が空欄なのに種別だけ入っている
対象が空欄なのに種別が入っている
```

問題がある場合は、どのシートの何行目を直すべきか表示する。

### 3. 臨時予定の入力済み行の見える化

入力済みの臨時予定行を種別ごとに色分けする。

例：

```text
休診       = 強めの網掛け
時短診療   = 薄い網掛け
休み       = スタッフ個別休みとして別色
時短勤務   = スタッフ個別時短として別色
```

### 4. 祝日・年末年始対応

祝日や医院独自の年末年始休診を半自動で `Exceptions` に入れるか検討する。

ただし、祝日は年ごとに変わり、医院独自休診もあるため、当面は手動入力でもよい。
