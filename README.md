# slack-to-google-docs-archive

Slack 投稿を Google Docs に保存する Google Apps Script です．
**参考: https://zenn.dev/gemcook/articles/38beb65aa8371c**
GASコードは `Code.gs` を適当なスプレッドシートの Apps Sript にコピペ．ファイル名は任意．
⚠️ 一時変数を保存するのでGoogleドキュメント不可．必ずスプレッドシートをGASファイルに設定．

## 使い方

1. 過去ログ取得: ボタンから `confirmCreateImportPastMessagesTrigger()` を 1 回実行する
2. 今後の投稿保存: 新規投稿タイミングで自動実行
3. 過去ログをやり直す: ボタンから `confirmResetImportPastMessages()` 実行後、`confirmCreateImportPastMessagesTrigger()` を再実行する
   - `resetImportPastMessages()` は過去ログ取得の進捗と処理済み記録を削除する

スプレッドシートの図形ボタンには，直接実行用の `createImportPastMessagesTrigger()` / `resetSlackEventQueue()` / `resetImportPastMessages()` ではなく，上記の `confirm...` 関数を割り当ててください。押下時に確認ダイアログが出て，OK の場合だけ実処理を呼び出します。

## 実装済み

### 手動実行(過去ログ)
- Bot 参加チャンネルの過去ログを 5 分ごとに分割取得する
- 親メッセージを基準に古い投稿が上、新しい投稿が下になるように保存する

### 自動実行(新規メッセージ)
- Slack Events API の投稿イベントを受け取る
- 新規投稿をキューに積み、1 分ごとに Google Docs へ保存する

### 共通
- チャンネル別フォルダ、年別ドキュメントに保存する
- 投稿者、時刻、本文、スレッド返信を保存する
- メンションをユーザー名に変換する
- 画像添付を Google Docs に埋め込む
- ユーザー名、チャンネル情報、処理済み状態を6hキャッシュ・記録する
- 各種トリガーの作成・削除、過去ログ進捗リセット、Slack イベントキューのリセットに対応する

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


## 注意点

- Slack 署名検証は未実装
- 既存 Docs の並び順は自動では並べ替えない
- 処理済み記録は `_slack_processed_messages` シートに保存する
- 過去ログ取得が異常停止して自動保存が進まない場合は、`resetImportPastMessages()` で `IMPORT_ACTIVE` を解除する
- 自動保存が復旧しない場合の最終手段として `resetSlackEventQueue()` を使えるが、未処理キューは削除される
