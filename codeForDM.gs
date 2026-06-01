// ===============================================================
// 初期設定・復旧メモ

// 1. 対象指定: 過去ログを取得したいDM/チャンネル内でBotをメンションする
// 2. 過去ログ取得: Botメンションで指定された会話だけを5分ごとのトリガーで分割取得する
// 3. 過去ログをやり直す: ボタンから `confirmResetImportPastMessages()` 実行後、対象会話でBotを再メンションする
//    - `resetImportPastMessages()` は過去ログ取得の進捗と処理済み記録を削除する
// ===============================================================
// 歯車>スクリプトプロパティに`SLACK_TOKEN`, `DOC_FOLDER_ID`を設定

// 1. SlackAPI管理画面>OAuth & Permissions>OAuth Tokens
const SLACK_TOKEN = PropertiesService.getScriptProperties().getProperty('SLACK_TOKEN');
if (!SLACK_TOKEN) {
  throw new Error("Script Properties に SLACK_TOKEN が設定されていません");
}

// 2. GoogleDriveのFolderのリンク>https://drive.google.com/drive/folders/ここがDOC_FOLDER_ID
const DOC_FOLDER_ID = PropertiesService.getScriptProperties().getProperty('DOC_FOLDER_ID');
if (!DOC_FOLDER_ID) {
  throw new Error("Script Properties に DOC_FOLDER_ID が設定されていません");
}
const PROCESSED_MESSAGES_SHEET_NAME = "_slack_processed_messages";
const SLACK_EVENT_QUEUE_PROP = "SLACK_EVENT_QUEUE";
const SLACK_EVENT_DONE_CACHE_KEYS_PROP = "SLACK_EVENT_DONE_CACHE_KEYS";
const SLACK_EVENT_DONE_CACHE_KEY_LIMIT = 100;
const IMPORT_TARGET_CHANNEL_PROP = "IMPORT_TARGET_CHANNEL";
const SLACK_SELF_USER_ID_CACHE_KEY = "SLACK_SELF_USER_ID";
let processedMessageKeyCache = null;
// ==============================================================
// 参考

// https://zenn.dev/gemcook/articles/38beb65aa8371c
// ===============================================================

// ボタン用関数 =======================================
// 確認ダイアログを表示
function confirmCreateImportPastMessagesTrigger() {
  confirmAndRun_(
    "指定会話の過去ログ取得を開始しますか？",
    "Botメンションで指定された会話だけを対象に，importPastMessages の5分ごとのトリガーを作成し，初回処理をすぐ実行します。",
    createImportPastMessagesTrigger
  );
}

function confirmResetSlackEventQueue() {
  confirmAndRun_(
    "Slackイベントキューをリセットしますか？",
    "未処理キュー，自動実行用プロパティ，関連キャッシュ，processSlackEventQueue トリガーを削除します。",
    resetSlackEventQueue
  );
}

function confirmResetImportPastMessages() {
  confirmAndRun_(
    "過去ログ取得の進捗をリセットしますか？",
    "IMPORT_ACTIVE，過去ログ取得の進捗，処理済み記録，importPastMessages トリガーを削除します。",
    resetImportPastMessages
  );
}

function confirmAndRun_(title, message, callback) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(title, message, ui.ButtonSet.OK_CANCEL);

  if (response !== ui.Button.OK) {
    Logger.log(`${title} キャンセルされました`);
    return;
  }

  callback();
}


// 1-1 ===================================================================
// =======================================================================
function importPastMessages() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    Logger.log("importPastMessages は前回実行中のためスキップします");
    return;
  }

  try{
    const props = PropertiesService.getScriptProperties();

    const channels = getImportTargetChannels();
    if (!channels.length) {
      Logger.log("過去ログ取得対象の会話が指定されていません。対象DM/チャンネルでBotをメンションしてください");
      props.deleteProperty("IMPORT_ACTIVE");
      deleteImportPastMessagesTrigger();
      return;
    }

    let channelIndex = Number(props.getProperty("IMPORT_CHANNEL_INDEX") || 0);
    let cursor = props.getProperty("IMPORT_CURSOR") || "";

    if (channelIndex >= channels.length) {
      Logger.log("指定会話の過去ログインポートが完了しました!");
      cleanupRuntimeProperties();
      deleteImportPastMessagesTrigger();
      return;
    }

    const channel = channels[channelIndex];
    const channelId = channel.id;
    const channelName = channel.name || channelId;

    Logger.log(`指定会話処理中: ${channelName} (${channelId})`);

    const result = getAllChannelMessages(channelId, cursor);
    const messages = result.messages;

    Logger.log(`${channelName}: この回では ${messages.length} 件を処理します`);

    const shouldPrepend = Boolean(cursor);
    const insertIndexesByYear = {};

    const writeImportedMessage = msg => {
      const options = {};

      if (shouldPrepend) {
        const year = getMessageThreadYear(msg);

        if (insertIndexesByYear[year] === undefined) {
          const doc = getOrCreateDocForYear(year, channel);
          insertIndexesByYear[year] = getHistoricalInsertIndex(doc.getBody());
        }

        options.insertIndex = insertIndexesByYear[year];
      }

      const result = writeMessageToDoc(msg, channel, options);

      if (shouldPrepend) {
        insertIndexesByYear[getMessageThreadYear(msg)] = result.nextInsertIndex;
      }
    };

    messages.forEach(msg => {
      const key = `${channelId}:${msg.ts}`;
      if (isMessageProcessed(key)) return;

      if (msg.reply_count && msg.reply_count > 0) {
        const threadMessages = getThreadMessages(channelId, msg.ts);

        if (threadMessages) {
          threadMessages.forEach(tMsg => {
            const threadKey = `${channelId}:${tMsg.ts}`;
            if (isMessageProcessed(threadKey)) return;

            writeImportedMessage(tMsg);
            markMessageProcessed(threadKey);
          });
        }
      } else {
        writeImportedMessage(msg);
        markMessageProcessed(key);
      }
    });

    if (result.nextCursor) {
      props.setProperty("IMPORT_CHANNEL_INDEX", String(channelIndex));
      props.setProperty("IMPORT_CURSOR", result.nextCursor);
      Logger.log("次回，同じ指定会話の続きを処理します");
      return;
    }

    const nextChannelIndex = channelIndex + 1;
    props.setProperty("IMPORT_CHANNEL_INDEX", String(nextChannelIndex));
    props.deleteProperty("IMPORT_CURSOR");

    if (nextChannelIndex >= channels.length) {
      Logger.log("指定会話の過去ログインポートが完了しました!");
      cleanupRuntimeProperties();
      deleteImportPastMessagesTrigger();
      return;
    }

    Logger.log("指定会話の処理が完了しました");
  } finally {
    lock.releaseLock();
  }
}

// 1-2 importPastMessages を5分ごとに自動実行するトリガーを作成
function createImportPastMessagesTrigger(runImmediately) {
  if (!getImportTargetChannel()) {
    Logger.log("過去ログ取得対象の会話が未指定のため，importPastMessages トリガーを作成しません");
    return;
  }

  deleteImportPastMessagesTrigger();
  PropertiesService.getScriptProperties().setProperty("IMPORT_ACTIVE", "1");

  ScriptApp
    .newTrigger("importPastMessages")
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log("importPastMessages を5分ごとに実行するトリガーを作成しました");

  if (runImmediately === false) {
    Logger.log("Slackイベント応答を優先するため，初回処理は次回トリガー実行に任せます");
    return;
  }

  // ボタン実行時は初回だけ待たずに実行する
  importPastMessages();
}

// 1-3 importPastMessages の時間主導トリガーを削除
function deleteImportPastMessagesTrigger() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "importPastMessages") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  Logger.log("importPastMessages の既存トリガーを削除しました");
}
// =======================================================================




// 2 =====================================================================
// =======================================================================

function doPost(e) {
  const params = JSON.parse(e.postData.contents);
  Logger.log(`doPost: request received / type=${params.type || "(none)"}`);

  if (params.type === 'url_verification') {
    Logger.log("doPost: url_verification challenge を返します");
    return ContentService
      .createTextOutput(params.challenge)
      .setMimeType(ContentService.MimeType.TEXT);
  }

  const event = params.event;
  if (!event) {
    Logger.log("doPost: event がないため終了します");
    return ContentService.createTextOutput("ok");
  }

  Logger.log(
    `doPost: event=${formatSlackMessageForLog_(event)} / subtype=${event.subtype || "(none)"}`
  );

  if (isImportTriggerMentionEvent(event)) {
    startImportForMentionedConversation_(event);
  } else {
    Logger.log(`doPost: Botメンションではないため過去ログ取得を開始しません / type=${event.type || "(none)"} / subtype=${event.subtype || "(none)"}`);
  }

  return ContentService.createTextOutput("ok");
}

function isImportTriggerMentionEvent(event) {
  if (!event || !event.channel) {
    return false;
  }

  if (event.type === "app_mention") {
    return true;
  }

  if (event.type !== "message" || event.subtype) {
    return false;
  }

  const botUserId = getSlackSelfUserId();
  if (!botUserId) {
    Logger.log("isImportTriggerMentionEvent: BotユーザーIDを取得できないためメンション判定できません");
    return false;
  }

  const mentionPattern = new RegExp(`<@${escapeRegExp_(botUserId)}(?:\\|[^>]+)?>`);
  return mentionPattern.test(String(event.text || ""));
}

function startImportForMentionedConversation_(event) {
  const channelId = event.channel;
  const props = PropertiesService.getScriptProperties();

  if (!channelId) {
    Logger.log("startImportForMentionedConversation_: channel がないため開始できません");
    return;
  }

  if (props.getProperty("IMPORT_ACTIVE") === "1") {
    Logger.log(`startImportForMentionedConversation_: importPastMessages 実行中のため新しい対象指定を無視します / channelId=${channelId}`);
    return;
  }

  const channel = getChannelInfo(channelId);
  setImportTargetChannel(channel);
  props.deleteProperty("IMPORT_CHANNEL_INDEX");
  props.deleteProperty("IMPORT_CURSOR");

  Logger.log(`startImportForMentionedConversation_: 過去ログ取得対象を設定しました / id=${channel.id} / name=${channel.name}`);
  createImportPastMessagesTrigger(false);
}

function escapeRegExp_(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// 2-1 doPostで積んだSlackイベントを後処理する
function processSlackEventQueue() {
  const props = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();

  let targets = [];
  Logger.log("processSlackEventQueue: 開始");

  try {
    if (!lock.tryLock(1000)) {
      Logger.log("processSlackEventQueue: lock取得失敗");
      return;
    }

    if (props.getProperty("IMPORT_ACTIVE") === "1") {
      Logger.log("processSlackEventQueue: importPastMessages 実行中のためキュー処理を保留します");
      return;
    }

    const queue = getSlackEventQueue_();
    Logger.log(`processSlackEventQueue: キュー取得 / 件数=${queue.length}`);

    if (!queue.length) {
      Logger.log("processSlackEventQueue: キューが空のためトリガーを削除して終了します");
      deleteSlackEventQueueTrigger();
      saveSlackEventQueue_([]);
      return;
    }

    // 1回で処理しすぎない
    targets = queue.slice(0, 5);
    const rest = queue.slice(5);

    saveSlackEventQueue_(rest);
    Logger.log(`processSlackEventQueue: 処理対象=${targets.length}件 / 残り=${rest.length}件`);

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
      Logger.log("processSlackEventQueue: lock解放");
    }
  }

  const failed = [];

  targets.forEach(item => {
    try {
      Logger.log(`processSlackEventQueue: item処理開始 / ${formatSlackEventQueueItemForLog_(item)}`);
      const cache = CacheService.getScriptCache();
      const doneKey = getSlackEventDoneCacheKey_(item.key);

      // 直近で処理済みならスキップ
      if (isMessageProcessed(item.key)) {
        Logger.log(`processSlackEventQueue: 処理済みシートによりスキップ / key=${item.key}`);
        return;
      }

      if (cache.get(doneKey)) {
        Logger.log(`processSlackEventQueue: 直近処理済みキャッシュによりスキップ / key=${item.key}`);
        return;
      }

      // Slackから該当メッセージを取り直す
      Logger.log(`processSlackEventQueue: Slack replies 取得開始 / channelId=${item.channelId} / threadTs=${item.threadTs}`);
      const threadMessages = getThreadMessages(item.channelId, item.threadTs);
      if (!threadMessages) {
        throw new Error("threadMessages is null");
      }
      Logger.log(`processSlackEventQueue: Slack replies 取得完了 / key=${item.key} / 件数=${threadMessages.length}`);

      const msg = threadMessages.find(m => m.ts === item.ts);
      if (!msg) {
        throw new Error(`message not found: ${item.key}`);
      }
      Logger.log(`processSlackEventQueue: 対象メッセージ検出 / ${formatSlackMessageForLog_(msg)}`);

      const channel = getChannelInfo(item.channelId);
      Logger.log(`processSlackEventQueue: チャンネル情報 / id=${channel.id} / name=${channel.name}`);

      const isReply = msg.thread_ts && msg.thread_ts !== msg.ts;
      Logger.log(`processSlackEventQueue: 書き込み判定 / key=${item.key} / isReply=${Boolean(isReply)}`);

      if (isReply) {
        const appended = appendReplyToExistingThread(msg, channel, threadMessages[0]);
        Logger.log(`processSlackEventQueue: 既存スレッド追記結果 / key=${item.key} / appended=${appended}`);

        if (!appended) {
          Logger.log(`processSlackEventQueue: 親位置が見つからないためスレッド全体を書き込み / key=${item.key} / 件数=${threadMessages.length}`);
          threadMessages.forEach(tMsg => {
            writeMessageToDoc(tMsg, channel);
            markMessageProcessed(`${item.channelId}:${tMsg.ts}`);
            Logger.log(`processSlackEventQueue: スレッド全体書き込み済み / key=${item.channelId}:${tMsg.ts}`);
          });
        } else {
          markMessageProcessed(item.key);
          Logger.log(`processSlackEventQueue: 返信追記済みとして記録 / key=${item.key}`);
        }
      } else {
        writeMessageToDoc(msg, channel);
        markMessageProcessed(item.key);
        Logger.log(`processSlackEventQueue: 新規投稿を書き込み済みとして記録 / key=${item.key}`);
      }

      cache.put(doneKey, "done", 21600); // 6時間だけ重複防止
      rememberSlackEventDoneCacheKey_(doneKey);
      Logger.log(`processSlackEventQueue: item処理完了 / key=${item.key}`);
    } catch (e) {
      Logger.log(`Slackイベント処理失敗: ${item.key} / ${e}`);
      failed.push(item);
    }
  });

  // 失敗分だけキューに戻す
  if (failed.length) {
    try {
      if (!lock.tryLock(1000)) {
        Logger.log("失敗キュー戻しのlock取得失敗");
        return;
      }

      const queue = getSlackEventQueue_();
      saveSlackEventQueue_(failed.concat(queue));
      Logger.log(`processSlackEventQueue: 失敗分をキューへ戻しました / 失敗=${failed.length}件 / キュー件数=${failed.length + queue.length}`);

    } finally {
      if (lock.hasLock()) {
        lock.releaseLock();
        Logger.log("processSlackEventQueue: 失敗キュー戻しlock解放");
      }
    }
  }

  // 未処理キューが残っていなければ，毎分トリガーを止める
  const remainingQueue = getSlackEventQueue_();
  Logger.log(`processSlackEventQueue: 終了前キュー確認 / 残り=${remainingQueue.length}件`);
  if (!remainingQueue.length) {
    deleteSlackEventQueueTrigger();
  }
  Logger.log("processSlackEventQueue: 終了");
}

function formatSlackEventQueueItemForLog_(item) {
  if (!item) return "(empty item)";

  return [
    `key=${item.key || "(none)"}`,
    `channelId=${item.channelId || "(none)"}`,
    `ts=${item.ts || "(none)"}`,
    `threadTs=${item.threadTs || "(none)"}`
  ].join(" / ");
}

function formatSlackMessageForLog_(msg) {
  if (!msg) return "(empty message)";

  return [
    `type=${msg.type || "(none)"}`,
    `channel=${msg.channel || "(none)"}`,
    `user=${msg.user || "(none)"}`,
    `ts=${msg.ts || "(none)"}`,
    `threadTs=${msg.thread_ts || "(none)"}`,
    `text=${truncateForLog_(msg.text || "", 120)}`
  ].join(" / ");
}

function truncateForLog_(value, maxLength) {
  const text = String(value || "").replace(/\s+/g, " ").trim();

  if (text.length <= maxLength) {
    return text || "(empty)";
  }

  return `${text.slice(0, maxLength)}...`;
}

// 2-2 processSlackEventQueue を1分ごとに実行するトリガーを作成
function createSlackEventQueueTrigger() {
  const triggers = ScriptApp.getProjectTriggers();

  const exists = triggers.some(trigger =>
    trigger.getHandlerFunction() === "processSlackEventQueue"
  );

  if (exists) {
    Logger.log("processSlackEventQueue のトリガーは既に存在します");
    return;
  }

  ScriptApp
    .newTrigger("processSlackEventQueue")
    .timeBased()
    .everyMinutes(1)
    .create();

  Logger.log("processSlackEventQueue を1分ごとに実行するトリガーを作成しました");
}

// 2-3 processSlackEventQueue の既存トリガーを削除
function deleteSlackEventQueueTrigger() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "processSlackEventQueue") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  Logger.log("processSlackEventQueue の既存トリガーを削除しました");
}

// 2-4 
function resetSlackEventQueue() {
  const queue = getSlackEventQueue_();
  const doneCacheKeys = getSlackEventDoneCacheKeys_(queue);

  if (doneCacheKeys.length) {
    removeSlackEventDoneCaches_(doneCacheKeys);
  }

  cleanupSlackEventQueueProperties();
  deleteSlackEventQueueTrigger();
  Logger.log("Slackイベントキュー，自動実行用プロパティ，関連キャッシュ，トリガーをリセットしました");
}

// 3 キャッシュ管理用関数 ================================================
// =======================================================================

function getSlackEventQueue_() {
  const rawQueue = PropertiesService.getScriptProperties().getProperty(SLACK_EVENT_QUEUE_PROP);
  if (!rawQueue) return [];

  try {
    const queue = JSON.parse(rawQueue);
    return Array.isArray(queue) ? queue : [];
  } catch (e) {
    Logger.log(`SlackイベントキューのJSON解析に失敗しました: ${e}`);
    return [];
  }
}

function saveSlackEventQueue_(queue) {
  const props = PropertiesService.getScriptProperties();

  if (!queue.length) {
    props.deleteProperty(SLACK_EVENT_QUEUE_PROP);
    return;
  }

  props.setProperty(SLACK_EVENT_QUEUE_PROP, JSON.stringify(queue));
}

function getSlackEventDoneCacheKey_(key) {
  return `DONE_${key}`;
}

function removeSlackEventDoneCaches_(doneCacheKeys) {
  const cache = CacheService.getScriptCache();

  for (let i = 0; i < doneCacheKeys.length; i += SLACK_EVENT_DONE_CACHE_KEY_LIMIT) {
    cache.removeAll(doneCacheKeys.slice(i, i + SLACK_EVENT_DONE_CACHE_KEY_LIMIT));
  }
}

function rememberSlackEventDoneCacheKey_(doneKey) {
  try {
    const props = PropertiesService.getScriptProperties();
    const keys = getSlackEventDoneCacheKeysFromProperties_();

    if (keys.indexOf(doneKey) === -1) {
      keys.push(doneKey);
    }

    props.setProperty(
      SLACK_EVENT_DONE_CACHE_KEYS_PROP,
      JSON.stringify(keys.slice(-SLACK_EVENT_DONE_CACHE_KEY_LIMIT))
    );
  } catch (e) {
    Logger.log(`Slackイベント処理済みキャッシュキーの記録に失敗しました: ${e}`);
  }
}

function getSlackEventDoneCacheKeys_(queue) {
  const queuedKeys = queue
    .map(item => item && item.key ? getSlackEventDoneCacheKey_(item.key) : "")
    .filter(Boolean);

  return Array.from(new Set(getSlackEventDoneCacheKeysFromProperties_().concat(queuedKeys)));
}

function getSlackEventDoneCacheKeysFromProperties_() {
  const rawKeys = PropertiesService
    .getScriptProperties()
    .getProperty(SLACK_EVENT_DONE_CACHE_KEYS_PROP);

  if (!rawKeys) return [];

  try {
    const keys = JSON.parse(rawKeys);
    return Array.isArray(keys) ? keys.filter(Boolean) : [];
  } catch (e) {
    Logger.log(`Slackイベント処理済みキャッシュキーのJSON解析に失敗しました: ${e}`);
    return [];
  }
}

// =======================================================================


// 4-1 既存スレッド位置を探して返信を挿入 ================================
// =======================================================================

function appendReplyToExistingThread(msg, channel, parentMsg) {
  const channelId = channel && channel.id ? channel.id : msg.channel || "unknown";
  const threadTs = msg.thread_ts || msg.ts;
  const year = Utilities.formatDate(new Date(parseFloat(threadTs) * 1000), "JST", "yyyy");
  Logger.log(
    `appendReplyToExistingThread: 開始 / channelId=${channelId} / threadTs=${threadTs} / year=${year} / msgTs=${msg.ts}`
  );

  const doc = getOrCreateDocForYear(year, channel || {
    id: channelId,
    name: channelId
  });
  const body = doc.getBody();
  let parentIndex = findParagraphIndex(body, getThreadMarker(channelId, threadTs));
  Logger.log(
    `appendReplyToExistingThread: marker検索 / marker=${getThreadMarker(channelId, threadTs)} / parentIndex=${parentIndex}`
  );

  if (parentIndex === -1 && parentMsg) {
    const parentName = getUserName(parentMsg.user);
    const parentTime = Utilities.formatDate(
      new Date(parseFloat(parentMsg.ts) * 1000),
      "JST",
      "yyyy/MM/dd HH:mm"
    );

    parentIndex = findParagraphIndex(body, `${parentName} (${parentTime})`);
    Logger.log(
      `appendReplyToExistingThread: 親見出し検索 / text=${parentName} (${parentTime}) / parentIndex=${parentIndex}`
    );
  }

  if (parentIndex === -1) {
    Logger.log(`appendReplyToExistingThread: 親位置が見つからないため追記不可 / channelId=${channelId} / threadTs=${threadTs}`);
    return false;
  }

  const insertIndex = getThreadInsertIndex(body, parentIndex);
  Logger.log(
    `appendReplyToExistingThread: 挿入位置決定 / parentIndex=${parentIndex} / insertIndex=${insertIndex}`
  );

  const result = writeMessageToDoc(msg, channel, {
    insertIndex: insertIndex
  });

  Logger.log(
    `appendReplyToExistingThread: 完了 / channelId=${channelId} / threadTs=${threadTs} / msgTs=${msg.ts} / result=${Boolean(result)}`
  );
  return Boolean(result);
}

// 4-2 Slack メッセージを Docs に書く

function writeMessageToDoc(msg, channel, options) {
  const threadTs = msg.thread_ts ? parseFloat(msg.thread_ts) : parseFloat(msg.ts);
  const year = Utilities.formatDate(new Date(threadTs * 1000), "JST", "yyyy");
  const channelIdForLog = channel && channel.id ? channel.id : msg.channel || "unknown";
  const isReplyForLog = msg.thread_ts && msg.thread_ts !== msg.ts;
  Logger.log(
    `writeMessageToDoc: 開始 / channelId=${channelIdForLog} / year=${year} / ts=${msg.ts} / threadTs=${msg.thread_ts || msg.ts} / isReply=${Boolean(isReplyForLog)} / insertIndex=${options && typeof options.insertIndex === "number" ? options.insertIndex : "(append)"}`
  );

  const doc = getOrCreateDocForYear(year, channel || {
    id: msg.channel || "unknown",
    name: msg.channel || "unknown"
  });
  const body = doc.getBody();
  let insertIndex = options && typeof options.insertIndex === "number"
    ? options.insertIndex
    : null;

  const addParagraph = text => {
    if (insertIndex === null) {
      return body.appendParagraph(text);
    }

    const paragraph = body.insertParagraph(insertIndex, text);
    insertIndex++;
    return paragraph;
  };

  const addImage = blob => {
    if (insertIndex === null) {
      return body.appendImage(blob);
    }

    const image = body.insertImage(insertIndex, blob);
    insertIndex++;
    return image;
  };

  const name = getUserName(msg.user);

  // 本文中の <@U...> メンションを @ユーザー名 に置換
  const text = String(msg.text || "").replace(
    /<@([A-Z0-9]+)(?:\|[^>]+)?>/g,
    (match, userId) => {
      return `@${getUserName(userId)}`;
    }
  );

  const time = Utilities.formatDate(
    new Date(parseFloat(msg.ts) * 1000),
    "JST",
    "yyyy/MM/dd HH:mm"
  );

  const isReply = msg.thread_ts && msg.thread_ts !== msg.ts;

  if (!isReply) {
    const channelId = channel && channel.id ? channel.id : msg.channel || "unknown";

    addParagraph(`\n============================`);
    addParagraph(getThreadMarker(channelId, msg.ts))
      .setFontSize(8)
      .setForegroundColor("#999999");
    addParagraph(`${name} (${time})`).setBold(true);
    addParagraph(text || "(テキストなし)");
  } else {
    const parentTime = Utilities.formatDate(
      new Date(parseFloat(msg.thread_ts) * 1000),
      "JST",
      "MM/dd HH:mm"
    );

    addParagraph(
      `  ┗ [${parentTime}] ${name} (${time}): ${text}`
    ).setIndentStart(20);
  }

  if (msg.files) {
    msg.files.forEach(file => {
      if (file.mimetype && file.mimetype.includes('image')) {
        try {
          const imgBlob = fetchSlackImageBlob(file);

          if (!imgBlob) {
            throw new Error("画像Blobを取得できませんでした");
          }

          const inlineImg = addImage(imgBlob);
          const width = 300;
          const height = inlineImg.getHeight() * (width / inlineImg.getWidth());

          inlineImg.setWidth(width).setHeight(height);

          if (isReply) {
            inlineImg.getParent().asParagraph().setIndentStart(40);
          }

        } catch (e) {
          const label = file.title || file.name || file.id || "unknown";
          const p = addParagraph(`  (画像の取得に失敗しました: ${label})`)
            .setItalic(true);

          if (isReply) {
            p.setIndentStart(40);
          }

          if (file.permalink) {
            const linkP = addParagraph(`  Slackファイル: ${file.permalink}`)
              .setItalic(true);

            if (isReply) {
              linkP.setIndentStart(40);
            }
          }

          Logger.log(`画像取得失敗: ${label} / ${e}`);
        }
      }
    });
  }

  Logger.log(
    `writeMessageToDoc: 完了 / channelId=${channelIdForLog} / ts=${msg.ts} / nextInsertIndex=${insertIndex === null ? "(append)" : insertIndex}`
  );

  return {
    nextInsertIndex: insertIndex
  };
}

// 4-3 Docs 内の挿入位置探索用

function getThreadMarker(channelId, threadTs) {
  return `[slack-thread:${channelId}:${threadTs}]`;
}

function findParagraphIndex(body, text) {
  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);

    if (
      child.getType() === DocumentApp.ElementType.PARAGRAPH &&
      child.asParagraph().getText() === text
    ) {
      return i;
    }
  }

  return -1;
}

function getThreadInsertIndex(body, markerIndex) {
  for (let i = markerIndex + 1; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);

    if (
      child.getType() === DocumentApp.ElementType.PARAGRAPH &&
      child.asParagraph().getText().includes("============================")
    ) {
      return i;
    }
  }

  return null;
}

function getMessageThreadYear(msg) {
  const threadTs = msg.thread_ts ? parseFloat(msg.thread_ts) : parseFloat(msg.ts);
  return Utilities.formatDate(new Date(threadTs * 1000), "JST", "yyyy");
}

function getHistoricalInsertIndex(body) {
  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);

    if (
      child.getType() === DocumentApp.ElementType.PARAGRAPH &&
      child.asParagraph().getHeading() === DocumentApp.ParagraphHeading.HEADING1
    ) {
      return i + 1;
    }
  }

  return 0;
}

// 5 =====================================================================
// =======================================================================

// Slack画像URLを複数候補から軽量に取得する
function fetchSlackImageBlob(file) {
  const urls = [
    file.url_private_download,
    file.url_private,
    file.thumb_1024,
    file.thumb_720,
    file.thumb_480,
    file.thumb_360
  ].filter(Boolean);

  for (const url of urls) {
    const res = UrlFetchApp.fetch(url, {
      headers: {
        "Authorization": "Bearer " + SLACK_TOKEN
      },
      followRedirects: true,
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();

    if (code >= 200 && code < 300) {
      return res.getBlob().setName(file.name || file.title || "slack_image");
    }

    Logger.log(`Slack画像URL取得失敗: HTTP ${code} / ${url}`);
  }

  return null;
}

// 6 =====================================================================
// =======================================================================
function getThreadMessages(channel, ts) {
  const messages = [];
  let cursor = "";

  do {
    const params = [
      `channel=${encodeURIComponent(channel)}`,
      `ts=${encodeURIComponent(ts)}`,
      "limit=200"
    ];

    if (cursor) {
      params.push(`cursor=${encodeURIComponent(cursor)}`);
    }

    Logger.log(`getThreadMessages: Slack API呼び出し / channel=${channel} / ts=${ts} / cursor=${cursor ? "あり" : "なし"}`);
    const url = `https://slack.com/api/conversations.replies?${params.join("&")}`;
    const res = UrlFetchApp.fetch(url, { "headers": { "Authorization": "Bearer " + SLACK_TOKEN } });
    const json = JSON.parse(res.getContentText());

    if (!json.ok) {
      Logger.log(`getThreadMessages: Slack API失敗 / channel=${channel} / ts=${ts} / error=${json.error}`);
      return null;
    }

    messages.push(...json.messages);

    cursor = json.response_metadata && json.response_metadata.next_cursor
      ? json.response_metadata.next_cursor
      : "";

    Logger.log(`getThreadMessages: Slack API成功 / channel=${channel} / ts=${ts} / 今回=${json.messages.length}件 / 累計=${messages.length}件 / nextCursor=${cursor ? "あり" : "なし"}`);

    if (cursor) {
      Utilities.sleep(1200);
    }

  } while (cursor);

  Logger.log(`getThreadMessages: Slack API取得完了 / channel=${channel} / ts=${ts} / 件数=${messages.length}`);
  return messages;
}

// 7 =====================================================================
// =======================================================================

function getUserName(userId) {
  if (!userId) return "unknown";

  const cache = CacheService.getScriptCache();
  const cachedName = cache.get(userId);
  if (cachedName) return cachedName;
    
  try {
    const url = `https://slack.com/api/users.info?user=${userId}`;
    const res = UrlFetchApp.fetch(url, { "headers": { "Authorization": "Bearer " + SLACK_TOKEN } });
    const json = JSON.parse(res.getContentText());

    if (json.ok) {
      const name = json.user.profile.display_name || json.user.real_name || userId;
      cache.put(userId, name, 21600);   // 6時間ユーザー名をキャッシュに保存
      return name;
    }
    return userId;
  } catch (e) { return userId; }
}

// 8 =====================================================================
// =======================================================================

function getOrCreateDocForYear(year, channel) {
  const channelId = channel && channel.id ? channel.id : "unknown";
  const channelName = channel && channel.name ? channel.name : channelId;

  const safeChannelName = String(channelName).replace(/[\\/:*?"<>|#]/g, "_");
  const channelFolderName = `${safeChannelName}_${channelId}`;
  Logger.log(`getOrCreateDocForYear: 開始 / year=${year} / channelFolder=${channelFolderName}`);

  const rootFolder = DriveApp.getFolderById(DOC_FOLDER_ID);
  const channelFolders = rootFolder.getFoldersByName(channelFolderName);

  const channelFolder = channelFolders.hasNext()
    ? channelFolders.next()
    : rootFolder.createFolder(channelFolderName);
  Logger.log(`getOrCreateDocForYear: チャンネルフォルダ取得 / name=${channelFolderName} / id=${channelFolder.getId()}`);

  const fileName = `Slack_Log_${year}`;
  const files = channelFolder.getFilesByName(fileName);

  if (files.hasNext()) {
    const file = files.next();
    Logger.log(`getOrCreateDocForYear: 既存Docを開きます / fileName=${fileName} / docId=${file.getId()}`);
    return DocumentApp.openById(file.getId());
  }

  const doc = DocumentApp.create(fileName);
  DriveApp.getFileById(doc.getId()).moveTo(channelFolder);
  Logger.log(`getOrCreateDocForYear: 新規Docを作成しました / fileName=${fileName} / docId=${doc.getId()}`);

  doc.getBody()
    .appendParagraph(`${year}年_Slack記録ドキュメント #${channelName}`)
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  return doc;
}

// 9 =====================================================================
// =======================================================================

function getImportTargetChannels() {
  const channel = getImportTargetChannel();
  return channel ? [channel] : [];
}

function getImportTargetChannel() {
  const rawChannel = PropertiesService
    .getScriptProperties()
    .getProperty(IMPORT_TARGET_CHANNEL_PROP);

  if (!rawChannel) {
    return null;
  }

  try {
    const channel = JSON.parse(rawChannel);

    if (channel && channel.id) {
      return {
        id: channel.id,
        name: channel.name || channel.id,
        type: channel.type || ""
      };
    }
  } catch (e) {
    Logger.log(`過去ログ取得対象のJSON解析に失敗しました: ${e}`);
  }

  return null;
}

function setImportTargetChannel(channel) {
  if (!channel || !channel.id) {
    throw new Error("過去ログ取得対象の会話IDがありません");
  }

  PropertiesService.getScriptProperties().setProperty(
    IMPORT_TARGET_CHANNEL_PROP,
    JSON.stringify({
      id: channel.id,
      name: channel.name || channel.id,
      type: channel.type || ""
    })
  );
}

function getBotJoinedChannels() {
  const channels = [];
  let cursor = "";

  do {
    const params = [
      "types=public_channel,private_channel",
      "exclude_archived=true",
      "limit=200"
    ];

    if (cursor) {
      params.push(`cursor=${encodeURIComponent(cursor)}`);
    }

    const url = `https://slack.com/api/users.conversations?${params.join("&")}`;
    const res = UrlFetchApp.fetch(url, {
      "headers": { "Authorization": "Bearer " + SLACK_TOKEN }
    });

    const json = JSON.parse(res.getContentText());

    if (!json.ok) {
      Logger.log("Bot参加チャンネル一覧の取得に失敗しました: " + json.error);
      break;
    }

    channels.push(...json.channels);

    cursor = json.response_metadata && json.response_metadata.next_cursor
      ? json.response_metadata.next_cursor
      : "";

    Utilities.sleep(1200);

  } while (cursor);

  return channels;
}

// 10 チャンネルの過去ログを取得 ==========================================
// =======================================================================

function getAllChannelMessages(channelId, cursor) {
  const params = [
    `channel=${encodeURIComponent(channelId)}`,
    `limit=50`
  ];

  if (cursor) {
    params.push(`cursor=${encodeURIComponent(cursor)}`);
  }

  const url = `https://slack.com/api/conversations.history?${params.join("&")}`;
  const res = UrlFetchApp.fetch(url, {
    "headers": { "Authorization": "Bearer " + SLACK_TOKEN }
  });

  const json = JSON.parse(res.getContentText());

  if (!json.ok) {
    throw new Error("過去ログの取得に失敗しました: " + json.error);
  }

  const nextCursor = json.response_metadata && json.response_metadata.next_cursor
    ? json.response_metadata.next_cursor
    : "";

  Utilities.sleep(1200);

  return {
    messages: json.messages.reverse(),
    nextCursor: nextCursor
  };
}

// 11 チャンネル情報を取得 ================================================
// =======================================================================

function getChannelInfo(channelId) {
  if (!channelId) {
    return {
      id: "unknown",
      name: "unknown"
    };
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = `channel:${channelId}`;
  const cached = cache.get(cacheKey);

  if (cached) {
    return JSON.parse(cached);
  }

  try {
    const url = `https://slack.com/api/conversations.info?channel=${encodeURIComponent(channelId)}`;
    const res = UrlFetchApp.fetch(url, {
      "headers": {
        "Authorization": "Bearer " + SLACK_TOKEN
      }
    });

    const json = JSON.parse(res.getContentText());

    if (json.ok && json.channel) {
      const channel = normalizeConversationInfo_(json.channel);

      cache.put(cacheKey, JSON.stringify(channel), 21600); // 6時間
      return channel;
    }

  } catch (e) {
    Logger.log(`チャンネル情報の取得に失敗しました: ${channelId} / ${e}`);
  }

  return {
    id: channelId,
    name: channelId,
    type: getFallbackConversationType_(channelId)
  };
}

function normalizeConversationInfo_(conversation) {
  const channelId = conversation.id || "unknown";

  return {
    id: channelId,
    name: getConversationDisplayName_(conversation),
    type: getConversationType_(conversation)
  };
}

function getConversationDisplayName_(conversation) {
  const channelId = conversation.id || "unknown";

  if (isDirectMessageConversation_(conversation)) {
    if (conversation.user) {
      return `DM_${getUserName(conversation.user)}`;
    }

    return `DM_${conversation.name || conversation.name_normalized || channelId}`;
  }

  if (isGroupDirectMessageConversation_(conversation)) {
    const selfUserId = getSlackSelfUserId();
    const memberNames = getConversationMemberUserIds(channelId)
      .filter(userId => userId && userId !== selfUserId)
      .map(getUserName)
      .filter(Boolean);

    if (memberNames.length) {
      return `GroupDM_${memberNames.join("_")}`;
    }

    return `GroupDM_${conversation.name || conversation.name_normalized || channelId}`;
  }

  return conversation.name || conversation.name_normalized || channelId;
}

function getConversationType_(conversation) {
  if (isDirectMessageConversation_(conversation)) return "im";
  if (isGroupDirectMessageConversation_(conversation)) return "mpim";
  if (conversation.is_channel) return "public_channel";
  if (conversation.is_group) return "private_channel";
  return getFallbackConversationType_(conversation.id || "");
}

function isDirectMessageConversation_(conversation) {
  return Boolean(
    conversation &&
    !conversation.is_mpim &&
    (
      conversation.is_im ||
      conversation.type === "im" ||
      String(conversation.id || "").startsWith("D")
    )
  );
}

function isGroupDirectMessageConversation_(conversation) {
  return Boolean(conversation && (conversation.is_mpim || conversation.type === "mpim"));
}

function getFallbackConversationType_(channelId) {
  if (channelId.startsWith("C")) return "public_channel";
  if (channelId.startsWith("D")) return "im";
  if (channelId.startsWith("G")) return "private_channel";
  return "unknown";
}

function getConversationMemberUserIds(channelId) {
  const members = [];
  let cursor = "";

  do {
    const params = [
      `channel=${encodeURIComponent(channelId)}`,
      "limit=200"
    ];

    if (cursor) {
      params.push(`cursor=${encodeURIComponent(cursor)}`);
    }

    try {
      const url = `https://slack.com/api/conversations.members?${params.join("&")}`;
      const res = UrlFetchApp.fetch(url, {
        "headers": { "Authorization": "Bearer " + SLACK_TOKEN }
      });
      const json = JSON.parse(res.getContentText());

      if (!json.ok) {
        Logger.log(`会話メンバー一覧の取得に失敗しました: ${channelId} / ${json.error}`);
        break;
      }

      members.push(...json.members);

      cursor = json.response_metadata && json.response_metadata.next_cursor
        ? json.response_metadata.next_cursor
        : "";

      if (cursor) {
        Utilities.sleep(1200);
      }
    } catch (e) {
      Logger.log(`会話メンバー一覧の取得に失敗しました: ${channelId} / ${e}`);
      break;
    }
  } while (cursor);

  return members;
}

function getSlackSelfUserId() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(SLACK_SELF_USER_ID_CACHE_KEY);

  if (cached) {
    return cached;
  }

  try {
    const url = "https://slack.com/api/auth.test";
    const res = UrlFetchApp.fetch(url, {
      "headers": { "Authorization": "Bearer " + SLACK_TOKEN }
    });
    const json = JSON.parse(res.getContentText());

    if (json.ok && json.user_id) {
      cache.put(SLACK_SELF_USER_ID_CACHE_KEY, json.user_id, 21600);
      return json.user_id;
    }

    Logger.log(`Slack自己ユーザーIDの取得に失敗しました: ${json.error || "unknown error"}`);
  } catch (e) {
    Logger.log(`Slack自己ユーザーIDの取得に失敗しました: ${e}`);
  }

  return "";
}

// 12 処理済みメッセージ・過去ログリセット管理 ============================
// =======================================================================

// 12-1 過去ログ取得の進捗と処理済み記録をリセット
function resetImportPastMessages() {
  cleanupRuntimeProperties();
  clearProcessedMessageRecords();
  deleteImportPastMessagesTrigger();
  cleanupLegacyProcessedMessageProperties();
  Logger.log("過去ログインポートの進捗と処理済み記録をリセットし，importPastMessages トリガーを削除しました");
}

// 12-2 メッセージが処理済みか確認
function isMessageProcessed(key) {
  return getProcessedMessageKeySet().has(key);
}

// 12-3 メッセージを処理済みとして記録
function markMessageProcessed(key) {
  if (!key || isMessageProcessed(key)) return;

  getProcessedMessagesSheet().appendRow([key, new Date()]);
  getProcessedMessageKeySet().add(key);
}

// 12-4 処理済みメッセージキーをシートから読み込む
function getProcessedMessageKeySet() {
  if (processedMessageKeyCache) {
    return processedMessageKeyCache;
  }

  const sheet = getProcessedMessagesSheet();
  const lastRow = sheet.getLastRow();
  const keys = new Set();

  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

    values.forEach(row => {
      if (row[0]) {
        keys.add(String(row[0]));
      }
    });
  }

  processedMessageKeyCache = keys;
  return processedMessageKeyCache;
}

// 12-5 処理済み記録用シートを取得または作成
function getProcessedMessagesSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error("処理済み記録用のスプレッドシートが見つかりません");
  }

  let sheet = spreadsheet.getSheetByName(PROCESSED_MESSAGES_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(PROCESSED_MESSAGES_SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["message_key", "processed_at"]);
  }

  return sheet;
}

// 12-6 処理済みメッセージ記録をクリア
function clearProcessedMessageRecords() {
  const sheet = getProcessedMessagesSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 2).clearContent();
  }

  processedMessageKeyCache = null;
  Logger.log("処理済みメッセージ記録をクリアしました");
}

// 13 一時スクリプトプロパティ削除 ========================================
// =======================================================================

// 13-1 過去ログ取得用の一時スクリプトプロパティだけ削除
// SLACK_TOKEN と DOC_FOLDER_ID は消さない
function cleanupRuntimeProperties() {
  const props = PropertiesService.getScriptProperties();
  const keys = props.getKeys();

  keys.forEach(key => {
    const isRuntimeKey =
      key === "IMPORT_CHANNEL_INDEX" ||
      key === "IMPORT_CURSOR" ||
      key === "IMPORT_ACTIVE" ||
      key === IMPORT_TARGET_CHANNEL_PROP;

    if (isRuntimeKey) {
      props.deleteProperty(key);
    }
  });

  Logger.log("過去ログ取得用の一時プロパティを削除しました。SLACK_TOKEN と DOC_FOLDER_ID は残しています");
}

// 13-2 Slackイベントキュー用の一時スクリプトプロパティを削除
function cleanupSlackEventQueueProperties() {
  const props = PropertiesService.getScriptProperties();
  const keys = props.getKeys();

  keys.forEach(key => {
    if (
      key === SLACK_EVENT_QUEUE_PROP ||
      key === SLACK_EVENT_DONE_CACHE_KEYS_PROP ||
      key.startsWith("DONE_")
    ) {
      props.deleteProperty(key);
    }
  });

  Logger.log("Slackイベントキュー用の一時プロパティを削除しました");
}

// 13-3 旧形式の処理済みメッセージプロパティを削除
function cleanupLegacyProcessedMessageProperties() {
  const props = PropertiesService.getScriptProperties();
  const keys = props.getKeys();

  keys.forEach(key => {
    if (/^[CGD][A-Z0-9]+:\d+\.\d+$/.test(key)) {
      props.deleteProperty(key);
    }
  });

  Logger.log("旧形式の処理済みメッセージプロパティを削除しました");
}

// TEST ===================================================================
// =======================================================================

// TEST-1 時間主導トリガー作成テスト
function testCreateSimpleTrigger() {
  deleteImportPastMessagesTrigger();

  ScriptApp
    .newTrigger("testTriggerTarget")
    .timeBased()
    .after(60 * 1000)
    .create();

  Logger.log("test trigger created");
}

// TEST-2 testCreateSimpleTrigger 用のトリガー先
function testTriggerTarget() {
  Logger.log("test trigger fired");
}

// TEST-3 Bot参加チャンネルのIDとチャンネル名を確認する
function testGetBotJoinedChannelNames() {
  const channels = getBotJoinedChannels();

  if (!channels.length) {
    Logger.log("Bot参加チャンネルがありません");
    return;
  }

  channels.forEach(channel => {
    Logger.log(
      `getBotJoinedChannels: id=${channel.id}, name=${channel.name}, name_normalized=${channel.name_normalized}`
    );

    const info = getChannelInfo(channel.id);

    Logger.log(
      `getChannelInfo:        id=${info.id}, name=${info.name}`
    );

    Logger.log("------------------------------");
  });
}

// TEST-4 doPost 経由で親投稿と返信追記を確認する
function testDoPostAppendText() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty("IMPORT_ACTIVE") === "1") {
    throw new Error("IMPORT_ACTIVE=1 のため doPost追記テストを実行できません。過去ログ取得の完了後に再実行してください。");
  }

  const originalGetThreadMessages = getThreadMessages;
  const originalQueue = getSlackEventQueue_();
  const originalDoneCacheKeys = props.getProperty(SLACK_EVENT_DONE_CACHE_KEYS_PROP);
  const channel = {
    id: "C_DOPOST_APPEND_TEST",
    name: "doPost追記テスト"
  };
  const user = {
    id: "U_DOPOST_APPEND_TEST",
    name: "doPostテスト"
  };
  const now = new Date();
  const label = Utilities.formatDate(now, "JST", "yyyy/MM/dd HH:mm:ss");
  const parentTs = createTestSlackTimestamp_(now.getTime());
  const replyTs = createTestSlackTimestamp_(now.getTime() + 1000);
  const parentMsg = {
    type: "message",
    channel: channel.id,
    user: user.id,
    text: `[doPost追記テスト 親] ${label}`,
    ts: parentTs
  };
  const replyMsg = {
    type: "message",
    channel: channel.id,
    user: user.id,
    text: `[doPost追記テスト 返信] ${label}`,
    ts: replyTs,
    thread_ts: parentTs
  };

  try {
    const cache = CacheService.getScriptCache();
    cache.put(`channel:${channel.id}`, JSON.stringify(channel), 21600);
    cache.put(user.id, user.name, 21600);

    Logger.log(`doPost追記テスト: 既存Slackイベントキュー ${originalQueue.length} 件を一時退避します`);
    saveSlackEventQueue_([]);

    getThreadMessages = function(channelId, threadTs) {
      if (channelId === channel.id && threadTs === parentTs) {
        return [parentMsg, replyMsg];
      }

      return originalGetThreadMessages(channelId, threadTs);
    };

    Logger.log("doPost追記テスト: 親投稿を doPost に投入します");
    doPost(buildTestDoPostRequest_(parentMsg));
    Logger.log(`doPost追記テスト: 親投稿投入後のキュー件数=${getSlackEventQueue_().length}`);
    processSlackEventQueue();

    Logger.log("doPost追記テスト: 返信投稿を doPost に投入します");
    doPost(buildTestDoPostRequest_(replyMsg));
    Logger.log(`doPost追記テスト: 返信投稿投入後のキュー件数=${getSlackEventQueue_().length}`);
    processSlackEventQueue();

    const year = Utilities.formatDate(now, "JST", "yyyy");
    Logger.log(`doPost追記テスト完了: #${channel.name} / Slack_Log_${year} を確認してください`);
  } finally {
    getThreadMessages = originalGetThreadMessages;
    cleanupTestDoPostQueue_(channel.id, [parentTs, replyTs]);
    saveSlackEventQueue_(originalQueue);

    if (originalDoneCacheKeys === null) {
      props.deleteProperty(SLACK_EVENT_DONE_CACHE_KEYS_PROP);
    } else {
      props.setProperty(SLACK_EVENT_DONE_CACHE_KEYS_PROP, originalDoneCacheKeys);
    }

    if (originalQueue.length) {
      createSlackEventQueueTrigger();
    } else {
      deleteSlackEventQueueTrigger();
    }

    CacheService.getScriptCache().removeAll([`channel:${channel.id}`, user.id]);
    Logger.log(`doPost追記テスト: 既存Slackイベントキュー ${originalQueue.length} 件を復元しました`);
  }
}

// TEST-5 doPost が Slackイベントをキューに積めるかだけを確認する
function testDoPostEnqueueOnly() {
  const props = PropertiesService.getScriptProperties();
  const originalQueue = getSlackEventQueue_();
  const originalDoneCacheKeys = props.getProperty(SLACK_EVENT_DONE_CACHE_KEYS_PROP);
  const msg = {
    type: "message",
    channel: "C_DOPOST_QUEUE_TEST",
    user: "U_DOPOST_QUEUE_TEST",
    text: "doPostキュー投入テスト",
    ts: createTestSlackTimestamp_(new Date().getTime())
  };

  try {
    Logger.log(`doPostキュー投入テスト: 既存Slackイベントキュー ${originalQueue.length} 件を一時退避します`);
    saveSlackEventQueue_([]);

    doPost(buildTestDoPostRequest_(msg));

    const queue = getSlackEventQueue_();
    Logger.log(`doPostキュー投入テスト: doPost後のキュー件数=${queue.length}`);

    if (!queue.length) {
      throw new Error("doPost がテスト投稿をキューに追加しませんでした");
    }

    Logger.log(`doPostキュー投入テスト: 追加されたキュー=${JSON.stringify(queue[0])}`);
  } finally {
    cleanupTestDoPostQueue_(msg.channel, [msg.ts]);
    saveSlackEventQueue_(originalQueue);

    if (originalDoneCacheKeys === null) {
      props.deleteProperty(SLACK_EVENT_DONE_CACHE_KEYS_PROP);
    } else {
      props.setProperty(SLACK_EVENT_DONE_CACHE_KEYS_PROP, originalDoneCacheKeys);
    }

    if (originalQueue.length) {
      createSlackEventQueueTrigger();
    } else {
      deleteSlackEventQueueTrigger();
    }

    Logger.log(`doPostキュー投入テスト: 既存Slackイベントキュー ${originalQueue.length} 件を復元しました`);
  }
}

function buildTestDoPostRequest_(msg) {
  const event = {
    type: "message",
    channel: msg.channel,
    user: msg.user,
    text: msg.text,
    ts: msg.ts
  };

  if (msg.thread_ts) {
    event.thread_ts = msg.thread_ts;
  }

  return {
    postData: {
      contents: JSON.stringify({
        type: "event_callback",
        event: event
      })
    }
  };
}

function createTestSlackTimestamp_(millis) {
  const seconds = Math.floor(millis / 1000);
  const micros = (millis % 1000) * 1000;
  return `${seconds}.${String(micros).padStart(6, "0")}`;
}

function cleanupTestDoPostQueue_(channelId, targetTsList) {
  try {
    const targetTsSet = new Set(targetTsList);
    const queue = getSlackEventQueue_();
    const cleanedQueue = queue.filter(item => {
      return item.channelId !== channelId || !targetTsSet.has(item.ts);
    });

    if (cleanedQueue.length === queue.length) {
      return;
    }

    saveSlackEventQueue_(cleanedQueue);

    if (!cleanedQueue.length) {
      deleteSlackEventQueueTrigger();
    }
  } catch (e) {
    Logger.log(`doPost追記テストのキュー掃除に失敗しました: ${e}`);
  }
}

// TEST-6 Slackイベントキューの詰まり確認
function testLogSlackEventQueueStatus() {
  const props = PropertiesService.getScriptProperties();
  const queue = getSlackEventQueue_();

  Logger.log(`IMPORT_ACTIVE=${props.getProperty("IMPORT_ACTIVE") || "(未設定)"}`);
  Logger.log(`Slackイベントキュー件数=${queue.length}`);

  queue.slice(0, 20).forEach((item, index) => {
    Logger.log(
      [
        `#${index + 1}`,
        `key=${item.key}`,
        `channelId=${item.channelId}`,
        `ts=${item.ts}`,
        `threadTs=${item.threadTs}`
      ].join(" / ")
    );
  });

  if (queue.length > 20) {
    Logger.log(`残り ${queue.length - 20} 件は省略しました`);
  }
}
