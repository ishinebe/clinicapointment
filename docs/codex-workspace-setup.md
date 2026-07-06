# Codex workspace setup

Codexで実装する前に、GitHub上のファイルをローカルworkspaceへ取得してください。

## 発生した問題

Codex workspace上に以下のファイルが存在しない場合、実装に進めません。

```text
docs/codex-prompt-template-draft.md
vba/AppointmentBook_Phase1.bas
```

これは、GitHubリポジトリの作成・更新は完了していても、Codexのworkspaceにリポジトリ一式がcloneされていない場合に起こります。

## 対応方法

Codexのworkspaceで、以下のいずれかを行ってください。

### 方法A: GitHubリポジトリをcloneする

```bash
git clone https://github.com/ishinebe/clinicapointment.git
cd clinicappointment
```

その後、Codexに以下を依頼してください。

```text
docs/codex-prompt-template-draft.md の内容に従って、TemplateDraft作成フェーズを実装してください。
既存の vba/AppointmentBook_Phase1.bas を壊さず、CreateTemplateDraft マクロを追加してください。
```

### 方法B: 既存workspaceにpullする

すでに `clinicapointment` をclone済みの場合は、対象ディレクトリに移動して以下を実行します。

```bash
git pull origin main
```

## Codexに読ませるべきファイル

```text
docs/design-principles-v1.3.md
docs/phase-plan.md
docs/codex-prompt-template-draft.md
vba/AppointmentBook_Phase1.bas
```

## 注意

Codexに単一のMarkdownファイルだけを渡すと、既存VBAファイルを確認できないため、
「既存の AppointmentBook_Phase1.bas を壊さず」という条件を満たせません。
必ずリポジトリ全体をworkspaceに配置してください。
