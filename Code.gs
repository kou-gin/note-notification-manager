// ===== 定数定義 =====
var SPREADSHEET_ID = "1xlMXt4nKDxmWaWw1FX40wNOIL7YimOnT_1WvNmv3jLM";
var MAIN_SHEET_NAME = "通知一覧";
var META_SHEET_NAME = "_meta";
var GIN_EMAIL = "kou.ainote@gmail.com";
var KOU_EMAIL = "sung.ksg@gmail.com";

// ===== doGet エントリーポイント =====
function doGet(e) {
  var action   = (e && e.parameter && e.parameter.action)   ? e.parameter.action   : "";
  var callback = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : "";
  var result;

  try {
    if (action === "run") {
      result = runSync();
    } else if (action === "getData") {
      result = { status: "ok", data: getData() };
    } else {
      result = { status: "error", message: "不明なactionです: " + action };
    }
  } catch (err) {
    result = { status: "error", message: err.message };
  }

  var json = JSON.stringify(result);

  if (callback) {
    // JSONP形式で返却
    var output = ContentService.createTextOutput(callback + "(" + json + ")");
    output.setMimeType(ContentService.MimeType.JAVASCRIPT);
    return output;
  }

  // callbackなしの場合はJSON
  var output = ContentService.createTextOutput(json);
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ===== メイン同期処理 =====
function runSync() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // _metaシートの準備
  var metaSheet = ss.getSheetByName(META_SHEET_NAME);
  if (!metaSheet) {
    metaSheet = ss.insertSheet(META_SHEET_NAME);
  }

  // 最終実行時刻の取得
  var lastRunValue = metaSheet.getRange("A1").getValue();
  var startDate;
  if (lastRunValue) {
    startDate = new Date(lastRunValue);
  } else {
    startDate = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
  }

  // Gmail検索（UNIXタイム秒）
  var afterUnix = Math.floor(startDate.getTime() / 1000);
  var query = "from:noreply@note.com after:" + afterUnix;
  var threads = GmailApp.search(query);

  // メインシートの既存データ取得（重複チェック用）
  var mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!mainSheet) {
    mainSheet = ss.insertSheet(MAIN_SHEET_NAME);
  }

  var existingUrls = getExistingUrls(mainSheet);
  var newRows = [];

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var msg = messages[j];
      var row = parseMessage(msg, existingUrls);
      if (row) {
        newRows.push(row);
        existingUrls[row[5]] = true; // F列のnoteURL
      }
    }
  }

  // 新規行を追記
  if (newRows.length > 0) {
    var lastRow = mainSheet.getLastRow();
    mainSheet.getRange(lastRow + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
  }

  // _metaシートに現在時刻を保存
  metaSheet.getRange("A1").setValue(new Date().toISOString());

  return { status: "ok", added: newRows.length, message: newRows.length + "件追記しました" };
}

// ===== メッセージ解析 =====
function parseMessage(msg, existingUrls) {
  var subject = msg.getSubject();

  // 種別判定
  var type = "";
  if (subject.indexOf("スキ") !== -1) {
    type = "スキ";
  } else if (subject.indexOf("フォロー") !== -1) {
    type = "フォロー";
  } else {
    return null;
  }

  // アカウント判定（To ヘッダーから）
  var to = msg.getTo();
  var account = "";
  if (to.indexOf(GIN_EMAIL) !== -1) {
    account = "Gin";
  } else if (to.indexOf(KOU_EMAIL) !== -1) {
    if (subject.indexOf("KOU") !== -1) {
      account = "KOU";
    } else {
      account = KOU_EMAIL;
    }
  } else {
    // Toが取得できない場合はスキップ
    return null;
  }

  // noteURL抽出（note.com%2F 以降をデコード）
  var body = msg.getBody();
  var noteUrl = extractNoteUrl(body);
  if (!noteUrl) {
    return null;
  }

  // 重複チェック
  if (existingUrls[noteUrl]) {
    return null;
  }

  // 件名から相手名を抽出
  var senderName = extractSenderName(subject, type);

  // 受信日時をJST文字列に変換して保存
  var rawDate = msg.getDate();
  var jstDate = Utilities.formatDate(rawDate, "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss");

  return [jstDate, type, senderName, account, noteUrl, "未対応"];
}

// ===== noteURL抽出（相手のユーザーページURL） =====
var OWN_ACCOUNTS = [
  "gin_ainote", "kou_gi_io",
  "unsubscribe_mails",
  "n", "tags", "search", "login", "signup", "contact"
];

function extractNoteUrl(body) {
  // ユーザーページURL: https://note.com/[username]（/n/ を含まないもの）
  var userPagePattern = /https:\/\/note\.com\/([^\/\s"'<>?#]+)/g;
  var candidates = [];
  var i, m;

  // エンコードされたURLを収集してデコード
  var encodedFound = body.match(/https?:\/\/[^"'\s]*note\.com%2F[^"'\s]*/g);
  if (encodedFound) {
    for (i = 0; i < encodedFound.length; i++) {
      try { candidates.push(decodeURIComponent(encodedFound[i])); } catch (e) {}
    }
  }

  // デコード済みnote.com URLを直接収集
  var directFound = body.match(/https?:\/\/note\.com\/[^\s"'<>]*/g);
  if (directFound) {
    for (i = 0; i < directFound.length; i++) {
      candidates.push(directFound[i]);
    }
  }

  // 全候補からユーザーページURLを抽出し、自分のアカウントを除外
  for (i = 0; i < candidates.length; i++) {
    userPagePattern.lastIndex = 0;
    m = userPagePattern.exec(candidates[i]);
    if (!m) continue;

    var username = m[1];
    // 自分のアカウント・システムパスはスキップ
    if (OWN_ACCOUNTS.indexOf(username) !== -1) continue;

    return "https://note.com/" + username;
  }

  return null;
}

// ===== 件名から送信者名を抽出 =====
function extractSenderName(subject, type) {
  // 例: "○○さんがあなたの記事にスキしました"
  //     "○○さんがあなたをフォローしました"
  var match = subject.match(/^(.+?)さんが/);
  if (match) return match[1];
  // さん抜きパターン
  match = subject.match(/^(.+?)が/);
  if (match) return match[1];
  return subject;
}

// ===== 件名から記事タイトルを抽出 =====
function extractArticleTitle(subject) {
  // 例: "○○さんがあなたの記事「記事タイトル」にスキしました"
  var match = subject.match(/「(.+?)」/);
  if (match) return match[1];
  // 角括弧パターン
  match = subject.match(/【(.+?)】/);
  if (match) return match[1];
  return "";
}

// ===== 既存URLを取得（重複チェック用） =====
function getExistingUrls(sheet) {
  var urls = {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return urls;

  // E列（5列目）がnoteURL
  var values = sheet.getRange(2, 5, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0]) {
      urls[values[i][0]] = true;
    }
  }
  return urls;
}

// ===== データ取得（フロントエンド用） =====
function getData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var lastCol = 6; // A〜F列
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  return values.map(function(row) {
    return {
      date:    row[0] ? (row[0] instanceof Date
               ? Utilities.formatDate(row[0], "Asia/Tokyo", "yyyy-MM-dd HH:mm")
               : String(row[0]).substring(0, 16)) : "",
      type:    row[1],
      name:    row[2],
      account: row[3],
      url:     row[4],
      status:  row[5]
    };
  });
}

// ===== 定期実行トリガー設定 =====
function setTrigger() {
  // 既存のrunSyncトリガーを削除してから再登録
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "runSync") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger("runSync")
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log("トリガーを設定しました：1時間ごとにrunSyncを実行");
}

// ===== 既存データのURL修正 =====
function fixExistingUrls() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(MAIN_SHEET_NAME);
  if (!sheet) {
    Logger.log("シートが見つかりません: " + MAIN_SHEET_NAME);
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("対象データなし");
    return;
  }

  // ユーザーページURL形式: https://note.com/[username]
  var userPagePattern = /https:\/\/note\.com\/([^\/?\s"'<>#]+)/;

  var urlRange = sheet.getRange(2, 6, lastRow - 1, 1);
  var values   = urlRange.getValues();
  var updated  = 0;
  var skipped  = 0;

  for (var i = 0; i < values.length; i++) {
    var original = String(values[i][0] || "").trim();
    if (!original) continue;

    // /n/ を含む場合はメール本文がないため変換不可→スキップ
    if (original.indexOf("/n/") !== -1) {
      skipped++;
      continue;
    }

    // /n/ を含まない場合はユーザーページURLに切り詰める
    var match = original.match(userPagePattern);
    if (!match) continue;

    var fixed = "https://note.com/" + match[1];
    if (fixed !== original) {
      values[i][0] = fixed;
      updated++;
    }
  }

  if (updated > 0) {
    urlRange.setValues(values);
  }
  Logger.log(updated + "件修正、" + skipped + "件スキップ（/n/ 含む）");
}
