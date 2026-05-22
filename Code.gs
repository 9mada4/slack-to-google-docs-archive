// ===============================================================
// 初期設定・復旧メモ
//
// 【今後のSlack投稿の自動保存】
// 最初に1回だけ 2-2 createSlackEventQueueTrigger() を手動実行する。
//
// この関数は，processSlackEventQueue() を1分ごとに実行する
// 時間主導トリガーをコードから作成するためのもの。
// 毎日実行する関数ではない。
//
// 自動保存がおかしくなった場合は，Apps Script左メニューの
// 「トリガー」から processSlackEventQueue のトリガーだけを削除し，
// createSlackEventQueueTrigger() をもう一度手動実行する。
//
// 【過去ログ取得】
// 過去ログを保存したい場合は，1-2 createImportPastMessagesTrigger() を
// 1回だけ手動実行する。
// これは importPastMessages() を5分ごとに実行するトリガーを作成する。
//
// 過去ログ取得を最初からやり直す場合は，resetImportPastMessages() を
// 手動実行して進捗をリセットしてから，
// createImportPastMessagesTrigger() を再実行する。
// ===============================================================
// 歯車>スクリプトプロパティに`SLACK_TOKEN`, `DOC_FOLDER_ID`を設定

// SlackAPI管理画面>OAuth & Permissions>OAuth Tokens
const SLACK_TOKEN = PropertiesService.getScriptProperties().getProperty('SLACK_TOKEN');
if (!SLACK_TOKEN) {
  throw new Error("Script Properties に SLACK_TOKEN が設定されていません");
}

// GoogleDriveのFolderのリンク>https://drive.google.com/drive/folders/ここがDOC_FOLDER_ID
const DOC_FOLDER_ID = PropertiesService.getScriptProperties().getProperty('DOC_FOLDER_ID');
if (!DOC_FOLDER_ID) {
  throw new Error("Script Properties に DOC_FOLDER_ID が設定されていません");
}
// ==============================================================

// 参考 ============
// https://zenn.dev/gemcook/articles/38beb65aa8371c
// ===============================================================


// 1-1 ===========================================================
function importPastMessages() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    Logger.log("importPastMessages は前回実行中のためスキップします");
    return;
  }

  try{
    const props = PropertiesService.getScriptProperties();

    const channels = getBotJoinedChannels();
    if (!channels.length) {
      Logger.log("Bot参加チャンネルがありません");
      props.deleteProperty("IMPORT_ACTIVE");
      deleteImportPastMessagesTrigger();
      return;
    }

    let channelIndex = Number(props.getProperty("IMPORT_CHANNEL_INDEX") || 0);
    let cursor = props.getProperty("IMPORT_CURSOR") || "";

    if (channelIndex >= channels.length) {
      Logger.log("Bot参加チャンネルの過去ログインポートが完了しました!");
      cleanupRuntimeProperties();
      deleteImportPastMessagesTrigger();
      return;
    }

    const channel = channels[channelIndex];
    const channelId = channel.id;
    const channelName = channel.name || channelId;

    Logger.log(`チャンネル処理中: ${channelName} (${channelId})`);

    const result = getAllChannelMessages(channelId, cursor);
    const messages = result.messages;

    Logger.log(`${channelName}: この回では ${messages.length} 件を処理します`);

    messages.forEach(msg => {
      const key = `${channelId}:${msg.ts}`;
      if (props.getProperty(key)) return;

      if (msg.reply_count && msg.reply_count > 0) {
        const threadMessages = getThreadMessages(channelId, msg.ts);

        if (threadMessages) {
          threadMessages.forEach(tMsg => {
            const threadKey = `${channelId}:${tMsg.ts}`;
            if (props.getProperty(threadKey)) return;

            writeMessageToDoc(tMsg, channel);
            props.setProperty(threadKey, "done");
          });
        }
      } else {
        writeMessageToDoc(msg, channel);
        props.setProperty(key, "done");
      }
    });

    if (result.nextCursor) {
      props.setProperty("IMPORT_CHANNEL_INDEX", String(channelIndex));
      props.setProperty("IMPORT_CURSOR", result.nextCursor);
      Logger.log("次回，同じチャンネルの続きを処理します");
      return;
    }

    props.setProperty("IMPORT_CHANNEL_INDEX", String(channelIndex + 1));
    props.deleteProperty("IMPORT_CURSOR");

    Logger.log("このチャンネルは完了しました。次回，次のチャンネルを処理します");
  } finally {
    lock.releaseLock();
  }
}

// 1-2 importPastMessages を5分ごとに自動実行するトリガーを作成
function createImportPastMessagesTrigger() {
  deleteImportPastMessagesTrigger();
  PropertiesService.getScriptProperties().setProperty("IMPORT_ACTIVE", "1");

  ScriptApp
    .newTrigger("importPastMessages")
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log("importPastMessages を5分ごとに実行するトリガーを作成しました");

  // 初回だけ待たずに実行する
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

// 2 =============================================================
function doPost(e) {
  const params = JSON.parse(e.postData.contents);

  if (params.type === 'url_verification') {
    return ContentService
      .createTextOutput(params.challenge)
      .setMimeType(ContentService.MimeType.TEXT);
  }

  const event = params.event;
  if (!event) {
    return ContentService.createTextOutput("ok");
  }

  // メッセージの新規投稿のみ処理対象
  if (event.type === 'message' && !event.subtype) {
    const channelId = event.channel || "unknown";
    const ts = event.ts;
    const threadTs = event.thread_ts || event.ts;

    if (!ts) {
      return ContentService.createTextOutput("ok");
    }

    const key = `${channelId}:${ts}`;
    const queueProp = "SLACK_EVENT_QUEUE";

    const lock = LockService.getScriptLock();

    try {
      if (!lock.tryLock(1000)) {
        throw new Error("Slack event queue lock timeout");
      }

      const props = PropertiesService.getScriptProperties();

      const cache = CacheService.getScriptCache();
      const doneKey = `DONE_${key}`;
      
      // 直近で処理済みなら何もしない
      if (cache.get(doneKey)) {
        return ContentService.createTextOutput("ok");
      }

      const queue = JSON.parse(props.getProperty(queueProp) || "[]");

      // すでにキュー済みなら重複追加しない
      const alreadyQueued = queue.some(item => item.key === key);
      if (!alreadyQueued) {
        queue.push({
          key: key,
          channelId: channelId,
          ts: ts,
          threadTs: threadTs
        });

        props.setProperty(queueProp, JSON.stringify(queue));

        // Slack投稿が来たときだけ，後処理トリガーを作成する
        createSlackEventQueueTrigger();
      }

    } finally {
      if (lock.hasLock()) {
        lock.releaseLock();
      }
    }
  }

  return ContentService.createTextOutput("ok");
}

// 2-1 doPostで積んだSlackイベントを後処理する
function processSlackEventQueue() {
  const queueProp = "SLACK_EVENT_QUEUE";
  const props = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();

  let targets = [];

  try {
    if (!lock.tryLock(1000)) {
      Logger.log("processSlackEventQueue: lock取得失敗");
      return;
    }

    if (props.getProperty("IMPORT_ACTIVE") === "1") {
      Logger.log("processSlackEventQueue: importPastMessages 実行中のためキュー処理を保留します");
      return;
    }

    const queue = JSON.parse(props.getProperty(queueProp) || "[]");

    if (!queue.length) {
      deleteSlackEventQueueTrigger();
      props.deleteProperty(queueProp);
      return;
    }

    // 1回で処理しすぎない
    targets = queue.slice(0, 5);
    const rest = queue.slice(5);

    props.setProperty(queueProp, JSON.stringify(rest));

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }

  const failed = [];

  targets.forEach(item => {
    try {
    const cache = CacheService.getScriptCache();
    const doneKey = `DONE_${item.key}`;

    // 直近で処理済みならスキップ
    if (props.getProperty(item.key) || cache.get(doneKey)) {
      return;
    }

      // Slackから該当メッセージを取り直す
      const threadMessages = getThreadMessages(item.channelId, item.threadTs);
      if (!threadMessages) {
        throw new Error("threadMessages is null");
      }

      const msg = threadMessages.find(m => m.ts === item.ts);
      if (!msg) {
        throw new Error(`message not found: ${item.key}`);
      }

      const channel = getChannelInfo(item.channelId);

      writeMessageToDoc(msg, channel);
      cache.put(doneKey, "done", 21600); // 6時間だけ重複防止
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

      const queue = JSON.parse(props.getProperty(queueProp) || "[]");
      props.setProperty(queueProp, JSON.stringify(failed.concat(queue)));

    } finally {
      if (lock.hasLock()) {
        lock.releaseLock();
      }
    }
  }

  // 未処理キューが残っていなければ，毎分トリガーを止める
  const remainingQueue = JSON.parse(props.getProperty(queueProp) || "[]");
  if (!remainingQueue.length) {
    props.deleteProperty(queueProp);
    deleteSlackEventQueueTrigger();
  }
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



// 3 ==========================================================
function writeMessageToDoc(msg, channel) {
  const threadTs = msg.thread_ts ? parseFloat(msg.thread_ts) : parseFloat(msg.ts);
  const year = Utilities.formatDate(new Date(threadTs * 1000), "JST", "yyyy");
  const doc = getOrCreateDocForYear(year, channel || {
    id: msg.channel || "unknown",
    name: msg.channel || "unknown"
  });
  const body = doc.getBody();

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
    body.appendParagraph(`\n============================`);
    body.appendParagraph(`${name} (${time})`).setBold(true);
    body.appendParagraph(text || "(テキストなし)");
  } else {
    const parentTime = Utilities.formatDate(
      new Date(parseFloat(msg.thread_ts) * 1000),
      "JST",
      "MM/dd HH:mm"
    );

    body.appendParagraph(
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

          const inlineImg = body.appendImage(imgBlob);
          const width = 300;
          const height = inlineImg.getHeight() * (width / inlineImg.getWidth());

          inlineImg.setWidth(width).setHeight(height);

          if (isReply) {
            inlineImg.getParent().asParagraph().setIndentStart(40);
          }

        } catch (e) {
          const label = file.title || file.name || file.id || "unknown";
          const p = body
            .appendParagraph(`  (画像の取得に失敗しました: ${label})`)
            .setItalic(true);

          if (isReply) {
            p.setIndentStart(40);
          }

          if (file.permalink) {
            const linkP = body
              .appendParagraph(`  Slackファイル: ${file.permalink}`)
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
}

// 3-1 ==============
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

// 4 ==================
function getThreadMessages(channel, ts) {
  const url = `https://slack.com/api/conversations.replies?channel=${channel}&ts=${ts}`;
  const res = UrlFetchApp.fetch(url, { "headers": { "Authorization": "Bearer " + SLACK_TOKEN } });
  const json = JSON.parse(res.getContentText());
  return json.ok ? json.messages : null;
}

// 5 ==========
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

// 6 ===============
function getOrCreateDocForYear(year, channel) {
  const channelId = channel && channel.id ? channel.id : "unknown";
  const channelName = channel && channel.name ? channel.name : channelId;

  const safeChannelName = String(channelName).replace(/[\\/:*?"<>|#]/g, "_");
  const channelFolderName = `${safeChannelName}_${channelId}`;

  const rootFolder = DriveApp.getFolderById(DOC_FOLDER_ID);
  const channelFolders = rootFolder.getFoldersByName(channelFolderName);

  const channelFolder = channelFolders.hasNext()
    ? channelFolders.next()
    : rootFolder.createFolder(channelFolderName);

  const fileName = `Slack_Log_${year}`;
  const files = channelFolder.getFilesByName(fileName);

  if (files.hasNext()) {
    return DocumentApp.openById(files.next().getId());
  }

  const doc = DocumentApp.create(fileName);
  DriveApp.getFileById(doc.getId()).moveTo(channelFolder);

  doc.getBody()
    .appendParagraph(`${year}年_Slack記録ドキュメント #${channelName}`)
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  return doc;
}

// 7 =============
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

// 8 =============
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
    Logger.log("過去ログの取得に失敗しました: " + json.error);
    return {
      messages: [],
      nextCursor: ""
    };
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

// 9 ============
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
      const channel = {
        id: json.channel.id,
        name: json.channel.name || json.channel.name_normalized || json.channel.id
      };

      cache.put(cacheKey, JSON.stringify(channel), 21600); // 6時間
      return channel;
    }

  } catch (e) {
    Logger.log(`チャンネル情報の取得に失敗しました: ${channelId} / ${e}`);
  }

  return {
    id: channelId,
    name: channelId
  };
}

// 10 ============
function resetImportPastMessages() {
  cleanupRuntimeProperties();
  deleteImportPastMessagesTrigger();
  Logger.log("過去ログインポートの進捗と処理済み記録をリセットし，トリガーも削除しました");
}

// 11 ============================================================
// 一時スクリプトプロパティだけ削除する
// SLACK_TOKEN と DOC_FOLDER_ID は消さない
function cleanupRuntimeProperties() {
  const props = PropertiesService.getScriptProperties();
  const keys = props.getKeys();

  keys.forEach(key => {
    const isRuntimeKey =
      key === "IMPORT_CHANNEL_INDEX" ||
      key === "IMPORT_CURSOR" ||
      key === "IMPORT_ACTIVE" ||
      key.startsWith("DONE_") ||
      /^[CGD][A-Z0-9]+:\d+\.\d+$/.test(key); // 旧形式: Cxxxx:171...

    if (isRuntimeKey) {
      props.deleteProperty(key);
    }
  });

  Logger.log("一時プロパティを削除しました。SLACK_TOKEN と DOC_FOLDER_ID は残しています");
}

// TEST =========================
function testCreateSimpleTrigger() {
  deleteImportPastMessagesTrigger();

  ScriptApp
    .newTrigger("testTriggerTarget")
    .timeBased()
    .after(60 * 1000)
    .create();

  Logger.log("test trigger created");
}

function testTriggerTarget() {
  Logger.log("test trigger fired");
}

// Bot参加チャンネルのIDとチャンネル名を確認する
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
