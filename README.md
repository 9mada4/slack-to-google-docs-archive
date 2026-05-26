# slack-to-google-docs-archive

Slack 投稿を Google Docs に保存する Google Apps Script です。
参考: https://zenn.dev/gemcook/articles/38beb65aa8371c

## 実装済み

### 手動実行(過去ログ)
- Bot 参加チャンネルの過去ログを 5 分ごとに分割取得する

### 自動実行(新規メッセージ)
- Slack Events API の投稿イベントを受け取る
- 新規投稿をキューに積み、1 分ごとに Google Docs へ保存する

### 共通
- チャンネル別フォルダ、年別ドキュメントに保存する
- 投稿者、時刻、本文、スレッド返信を保存する
- メンションをユーザー名に変換する
- 画像添付を Google Docs に埋め込む
- ユーザー名、チャンネル情報、処理済み状態を6hキャッシュ・記録する
- 各種トリガーの作成・削除、過去ログ進捗リセットに対応する

## 設定

Apps Script の 左の歯車 > スクリプトプロパティ に設定します。

1. `SLACK_TOKEN`

Slack App は Bot User OAuth Token (`xoxb-...`) を使います。

OAuth & Permissions > Bot Token Scopes:

- `channels:read`
- `channels:history`
- `groups:read`
- `groups:history`
- `files:read`
- `users:read`

Event Subscriptions > Subscribe to bot events:

- Enable Events: On
- Subscribe to bot events:
  - `message.channels`
  - `message.groups`

参考記事の `reactions:read` / `reaction_added` は、現行実装では未使用です。

2. `DOC_FOLDER_ID`
- GoogleDriveのFolderのリンク>https://drive.google.com/drive/folders/ここがDOC_FOLDER_ID

## 使い方

1. 過去ログ取得: `createImportPastMessagesTrigger()` を 1 回実行する
2. 今後の投稿保存: `createSlackEventQueueTrigger()` を 1 回実行する
3. 過去ログをやり直す: `resetImportPastMessages()` 実行後、`createImportPastMessagesTrigger()` を再実行する

## 注意点

- Slack 署名検証は未実装
- 長いスレッドのページング取得は未対応
- 大量投稿では Script Properties の容量上限に注意
