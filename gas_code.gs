/**
 * 「応援の循環」キックオフセミナー（2026年10月2日（金）20:00〜22:00 JST）
 * ── 申込管理シート連携 + 自動返信メール + リマインド
 * ── 既存生向けと一般向けの2つのLPから同じGASへ送信され、type で振り分けて別タブに記録する
 *
 * 【リマインドは2系統】
 *   ①Flodesk（本命）：申込のたびに、まだ送信時刻が来ていない回のセグメント「応援の循環_R1〜R6」へ自動追加。
 *      Flodesk側のワークフロー（開始条件＝セグメントに追加されたとき→指定日時まで待機→メール）が送る。
 *      スクリプトプロパティ FLODESK_API_KEY が必要（なつこさんが登録。AIは触らない）。
 *   ②GmailのGAS送信（保険）：setupTriggers で6つの時間トリガーを作る。Flodeskの完成・テスト確認後に
 *      ?action=stopGmailReminders で止める（重複送信を避けるため）。
 *   R1 9/29 20:00 ／ R2 9/30 20:00 ／ R3 10/1 20:00 ／ R4 10/2 08:00 ／ R5 10/2 19:00 ／ R6 10/2 20:00（日本時間）
 *
 * 🚨有料化・決済導線は一切なし。個人情報をAIに返さない：inspect等は件数・見出しだけを返す
 */

const CONFIG = {
  SENDER_NAME: "IOBオーガニックスクール事務局",
  SENDER_EMAIL: "school@iob.bio",
  NOTIFY_TO: "school@iob.bio"
};

const MEET = "https://meet.google.com/rjz-buhr-eyo";
const WHEN = "10月2日（金）20:00〜22:00（日本時間）";
const BAND_URL = "https://band.us/n/a0a6b5M8hfTc8";
const ADMIN_KEY = "zICLZ91dIQA5ZjFN200sAjhk";
const TEST_NAME = "テスト（アーニャ）";
const TEST_EMAIL = "organiclifeingermany@gmail.com";

/** リマインド6回（key, 送信日時, 見出し列名） */
const REMINDERS = [
  { key: "d3", at: "2026-09-29T20:00:00+09:00", label: "リマインド9/29(3日前)" },
  { key: "d2", at: "2026-09-30T20:00:00+09:00", label: "リマインド9/30(2日前)" },
  { key: "d1", at: "2026-10-01T20:00:00+09:00", label: "リマインド10/1(前日)" },
  { key: "m",  at: "2026-10-02T08:00:00+09:00", label: "リマインド10/2朝" },
  { key: "h1", at: "2026-10-02T19:00:00+09:00", label: "リマインド10/2 19時" },
  { key: "s",  at: "2026-10-02T20:00:00+09:00", label: "リマインド10/2 20時" }
];

/** 既存生向け・一般向けのシート設定。BAND参加状況の列は既存生向けだけにある */
const EVENTS = {
  "応援の循環キックオフ": {
    sheet: "応援の循環キックオフ申込",
    base: ["申込日時", "お名前", "メールアドレス", "BAND参加状況", "高田さんへの質問", "通知結果", "返信結果"],
    hasBand: true
  },
  "応援の循環キックオフ_一般": {
    sheet: "応援の循環キックオフ申込_一般",
    base: ["申込日時", "お名前", "メールアドレス", "高田さんへの質問", "通知結果", "返信結果"],
    hasBand: false
  }
};
function headersOf_(cfg) { return cfg.base.concat(REMINDERS.map(function (r) { return r.label; }), ["Flodesk追加"]); }
function flodeskCol_(cfg) { return cfg.base.length + REMINDERS.length + 1; }
function reminderCol_(cfg, key) {
  for (var i = 0; i < REMINDERS.length; i++) if (REMINDERS[i].key === key) return cfg.base.length + 1 + i;
  return 0;
}

function eventSheet_(ss, cfg) {
  var headers = headersOf_(cfg);
  var sh = ss.getSheetByName(cfg.sheet);
  if (!sh) {
    sh = ss.insertSheet(cfg.sheet);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  } else if (sh.getRange(1, headers.length).getValue() !== headers[headers.length - 1]) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);   // 見出しを最新（リマインド6列）に更新
    sh.setFrozenRows(1);
  }
  sh.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  return sh;
}

function signature_() {
  return "──────\n" +
         "Institut für Organic Business GmbH\n" +
         "オーガニックビジネス研究所 スクール事務局\n" +
         "「オーガニックが、あたりまえ。」な社会へ\n" +
         "営業時間：日本時間 平日10〜16時／お問い合わせは3営業日以内にお答えします\n" +
         "──────";
}

/* ============ メール送信（Gmail） ============ */
function gmailOpts_() {
  var opts = { name: CONFIG.SENDER_NAME };
  try { if (GmailApp.getAliases().indexOf(CONFIG.SENDER_EMAIL) !== -1) opts.from = CONFIG.SENDER_EMAIL; } catch (e) {}
  return opts;
}
function sendMail_(to, subject, body) {
  GmailApp.sendEmail(to, subject, body, gmailOpts_());
  return "gmail";
}

/* ============ Flodesk連携（申込者をリマインド用セグメントへ自動追加） ============ */
const FLODESK_TEST_SEGMENT = "応援の循環_テスト";
function segName_(i) { return "応援の循環_R" + (i + 1); }
function flodeskKey_() { return PropertiesService.getScriptProperties().getProperty("FLODESK_API_KEY"); }
function flodeskFetch_(method, path, payload) {
  var opt = { method: method, contentType: "application/json", muteHttpExceptions: true,
    headers: { Authorization: "Basic " + Utilities.base64Encode(flodeskKey_() + ":") } };
  if (payload) opt.payload = JSON.stringify(payload);
  return UrlFetchApp.fetch("https://api.flodesk.com/v1" + path, opt);
}
/** セグメント名→IDの対応（Flodeskから取得。6時間キャッシュ） */
function segmentIds_() {
  var cache = CacheService.getScriptCache(), hit = cache.get("segids");
  if (hit) return JSON.parse(hit);
  var map = {};
  for (var page = 1; page <= 10; page++) {
    var res = flodeskFetch_("get", "/segments?per_page=100&page=" + page);
    if (res.getResponseCode() !== 200) throw new Error("Flodesk segments HTTP " + res.getResponseCode());
    var data = JSON.parse(res.getContentText());
    var list = data.data || data.segments || [];
    list.forEach(function (sg) { map[sg.name] = sg.id; });
    if (list.length < 100) break;
  }
  cache.put("segids", JSON.stringify(map), 21600);
  return map;
}
/** 申込者を、まだ送信時刻が来ていない回のセグメントへ追加。戻り値は状態の短い文字列 */
function flodeskAdd_(email, name, testSegment) {
  if (!flodeskKey_()) return "Flodesk未設定";
  var ids = segmentIds_(), want = [], missing = [];
  if (testSegment) { want = [FLODESK_TEST_SEGMENT]; }
  else {
    var limit = Date.now() + 5 * 60 * 1000;                       // 5分以内に迫った回は間に合わないので追加しない
    REMINDERS.forEach(function (r, i) { if (new Date(r.at).getTime() > limit) want.push(segName_(i)); });
  }
  if (!want.length) return "追加対象なし（全て送信時刻を過ぎている）";
  var segIds = [];
  want.forEach(function (n) { if (ids[n]) segIds.push(ids[n]); else missing.push(n); });
  if (missing.length) { CacheService.getScriptCache().remove("segids"); return "ERR セグメント未作成: " + missing.join(","); }
  var res = flodeskFetch_("post", "/subscribers", { email: email, first_name: name, segment_ids: segIds });
  var code = res.getResponseCode();
  return (code >= 200 && code < 300) ? "追加OK " + segIds.length + "件" : "ERR HTTP " + code;
}
/** エディタから実行：Flodeskへの接続確認（権限の承認も兼ねる）。作成済みセグメントの有無だけを表示する */
function checkFlodesk() {
  if (!flodeskKey_()) { Logger.log("FLODESK_API_KEY が未登録です"); return; }
  CacheService.getScriptCache().remove("segids");
  var ids = segmentIds_(), names = REMINDERS.map(function (r, i) { return segName_(i); }).concat([FLODESK_TEST_SEGMENT]);
  names.forEach(function (n) { Logger.log(n + "：" + (ids[n] ? "あり" : "まだ作成されていません")); });
  Logger.log("接続OK");
}
/** 既存の申込者を一括でセグメントへ追加（件数だけ返す。氏名・メールは返さない） */
function backfillFlodesk_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), out = { added: 0, failed: 0, skipped: 0, noKey: !flodeskKey_() };
  if (out.noKey) return out;
  for (var type in EVENTS) {
    var cfg = EVENTS[type], sh = eventSheet_(ss, cfg), col = flodeskCol_(cfg), width = headersOf_(cfg).length;
    for (var r = 2; r <= sh.getLastRow(); r++) {
      var v = sh.getRange(r, 1, 1, width).getValues()[0], name = String(v[1] || ""), email = String(v[2] || "").trim();
      if (!email || (name === TEST_NAME && email === TEST_EMAIL)) { out.skipped++; continue; }
      if (String(v[col - 1]).indexOf("追加OK") === 0) { out.skipped++; continue; }
      var st = flodeskAdd_(email, name, false);
      sh.getRange(r, col).setValue(st);
      if (st.indexOf("追加OK") === 0) out.added++; else out.failed++;
    }
  }
  return out;
}

/* ============ 本文 ============ */
function applyReplyBody(cfg, d) {
  var bandBlock = "";
  if (cfg.hasBand && d.bandStatus !== "参加済み") {
    bandBlock =
      "■ この2時間の続きは、BANDで\n\n" +
      "「IOB 応援の循環 2026」は、IOBで学んだ仲間が近況を知り、つながり、応援を渡し合っていくための場所です。\n" +
      "よかったら、この機会にBANDにも加わってください。\n" +
      BAND_URL + "\n\n\n";
  }
  var recapBlock = "■ お申し込み内容（控え）\n" +
    (cfg.hasBand ? "BAND参加状況：" + (d.bandStatus || "") + "\n" : "") +
    "高田さんへの質問：" + (d.question || "（なし）") + "\n\n";
  return d.name + " 様\n\n" +
    "IOBオーガニックスクール事務局です。\n" +
    "「応援の循環」キックオフセミナーにお申し込みいただき、ありがとうございます。\n\n" +
    "■ 日時\n" + WHEN + "\n\n" +
    "■ 参加用リンク（Google Meet）\n" + MEET + "\n" +
    "時間になったら、このリンクを開いてください。\n" +
    "カメラは切ったままでも、聞いているだけでも大丈夫です。\n\n" +
    "■ 当日やること（2時間）\n" +
    "・オープニング\n" +
    "・Connect（自分自身の「応援の共体験」を思い出すワーク）\n" +
    "・高田洋平さんトーク\n" +
    "・Commonワーク（実際にみんなで考え、言葉にする）\n" +
    "・共有\n" +
    "・クロージング\n\n" +
    bandBlock +
    "■ リマインド\n" +
    "開催日までのあいだ、このメールアドレス宛にリマインドをお送りします。\n\n" +
    recapBlock +
    "何も持ってこなくて大丈夫です。\n" +
    "当日、画面の向こうでお会いできたら嬉しいです。\n\n" +
    "ご不明な点は school@iob.bio までご連絡ください。\n\n" +
    signature_();
}

function reminderSubject_(key) {
  return {
    d3: "あと3日｜10月2日（金）20時から「応援の循環」キックオフセミナー",
    d2: "あと2日｜「応援の循環」キックオフセミナー",
    d1: "明日20時です｜「応援の循環」キックオフセミナーのご案内",
    m:  "今夜20時からです｜「応援の循環」キックオフセミナー（Google Meetリンク）",
    h1: "あと1時間です｜「応援の循環」キックオフセミナー（Google Meetリンク）",
    s:  "始まりました｜「応援の循環」キックオフセミナー（Google Meetリンク）"
  }[key];
}
function reminderBody_(name, key) {
  var head = name + " 様\n\nIOBオーガニックスクール事務局です。\n\n";
  var link = "▼参加用リンク（Google Meet）\n" + MEET + "\n\n";
  var tail = signature_();
  var b = {
    d3: "「応援の循環」キックオフセミナーまで、あと3日になりました。\n\n■ 日時\n" + WHEN + "\n\n" + link +
        "当日は、高田洋平さんと一緒に「応援が巡る関係」を実際にワークしながら体験します。\nカメラは切ったままでも、聞いているだけでも大丈夫です。\n\n",
    d2: "「応援の循環」キックオフセミナーまで、あと2日です。\n\n" +
        "当日のはじめに、「これまでに誰かに応援してもらったこと、誰かを応援したこと」を思い出す時間があります。\n" +
        "いまのうちに、ぼんやり思い浮かべておいてください。書き留めておく必要はありません。\n\n■ 日時\n" + WHEN + "\n\n" + link,
    d1: "「応援の循環」キックオフセミナーは、明日10月2日（金）20:00からです（日本時間・22:00ごろまで）。\n\n" + link +
        "時間になったら、このリンクを開いてください。\n何も持ってこなくて大丈夫です。\n\n明日の夜、画面の向こうでお会いできたら嬉しいです。\n\n",
    m:  "今夜、「応援の循環」キックオフセミナーです。\n\n■ 日時\n本日 " + WHEN.replace("10月2日（金）", "") + "\n\n" + link +
        "何も持ってこなくて大丈夫です。カメラは切ったままでも、聞いているだけでも。\n\n",
    h1: "「応援の循環」キックオフセミナーは、あと1時間、20:00からです。\n\n" + link +
        "時間になったら、このリンクを開いてください。\n\n",
    s:  "「応援の循環」キックオフセミナーが、いま始まりました。\n\n" + link + "こちらのリンクから入室してください。\n\n"
  }[key];
  return head + b + tail;
}

/* ============ 申込受付 ============ */
function applyEvent_(ss, cfg, d) {
  var sheet = eventSheet_(ss, cfg);
  // 二重登録防止：同じメールアドレスが直近10分以内にあれば、何もせず成功を返す（LP側の再送対策）
  var lr = sheet.getLastRow();
  if (lr >= 2 && d.email) {
    var from = Math.max(2, lr - 29), rows = sheet.getRange(from, 1, lr - from + 1, 3).getDisplayValues();
    for (var q = rows.length - 1; q >= 0; q--) {
      var t = Date.parse(String(rows[q][0]).replace(/\//g, "-").replace(" ", "T") + "+09:00");
      if (String(rows[q][2]).trim().toLowerCase() === String(d.email).trim().toLowerCase() && Date.now() - t < 10 * 60 * 1000) return { result: "ok", dup: true };
    }
  }
  var ts = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
  var vals = cfg.hasBand
    ? [ts, d.name || "", d.email || "", d.bandStatus || "", d.question || "", "", ""]
    : [ts, d.name || "", d.email || "", d.question || "", "", ""];
  for (var i = 0; i < REMINDERS.length + 1; i++) vals.push("");
  sheet.appendRow(vals);
  var rowNum = sheet.getLastRow();
  var notifyCol = cfg.base.length - 1, replyCol = cfg.base.length;
  var notifyStatus = "", replyStatus = "";

  try {
    MailApp.sendEmail({
      to: CONFIG.NOTIFY_TO,
      subject: "【応援の循環キックオフ申込】" + d.name + " 様" + (cfg.hasBand ? "" : "（一般）"),
      body: "「応援の循環」キックオフセミナーの申込が入りました。\n\n" +
        "お名前：" + d.name + "\nメール：" + d.email + "\n" +
        (cfg.hasBand ? "BAND参加状況：" + (d.bandStatus || "") + "\n" : "") +
        "高田さんへの質問：" + (d.question || "（なし）") + "\n日時：" + ts
    });
    notifyStatus = "通知OK " + ts;
  } catch (e1) { notifyStatus = "通知ERR: " + e1; }
  sheet.getRange(rowNum, notifyCol).setValue(notifyStatus);

  if (d.email) {
    try {
      var via = sendMail_(d.email, "【お申し込みありがとうございます】「応援の循環」キックオフセミナーのご案内（Google Meetリンク）", applyReplyBody(cfg, d));
      replyStatus = "返信OK " + ts + "（" + via + "）";
    } catch (e2) { replyStatus = "返信ERR: " + e2; }
    sheet.getRange(rowNum, replyCol).setValue(replyStatus);
    try { sheet.getRange(rowNum, flodeskCol_(cfg)).setValue(flodeskAdd_(d.email, d.name || "", false)); }
    catch (e3) { sheet.getRange(rowNum, flodeskCol_(cfg)).setValue("ERR " + e3); }
  }
  return { result: "ok" };
}

function doPost(e) { return handle(e); }
function doGet(e)  { return handle(e); }

function handle(e) {
  try {
    if (e && e.parameter && e.parameter.action) {
      return ContentService.createTextOutput(JSON.stringify(admin_(e.parameter))).setMimeType(ContentService.MimeType.JSON);
    }
    // 🚨 e.parameter.payload を先に見ること（form-urlencoded送信では postData.contents が生文字列になるため）
    var raw = "";
    if (e && e.parameter && e.parameter.payload) raw = e.parameter.payload;
    else if (e && e.postData && e.postData.contents) raw = e.postData.contents;
    if (!raw) return ContentService.createTextOutput(JSON.stringify({ result: "no-data" })).setMimeType(ContentService.MimeType.JSON);

    const d = JSON.parse(raw);
    const cfg = EVENTS[d.type];
    if (!cfg) return ContentService.createTextOutput(JSON.stringify({ result: "unknown-type" })).setMimeType(ContentService.MimeType.JSON);
    return ContentService.createTextOutput(JSON.stringify(applyEvent_(SpreadsheetApp.getActiveSpreadsheet(), cfg, d))).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ result: "error", message: String(err) })).setMimeType(ContentService.MimeType.JSON);
  }
}

/* ============ リマインド ============ */
function sendReminderD3() { return sendReminderAll_("d3", false); }
function sendReminderD2() { return sendReminderAll_("d2", false); }
function sendReminderD1() { return sendReminderAll_("d1", false); }
function sendReminderM()  { return sendReminderAll_("m", false); }
function sendReminderH1() { return sendReminderAll_("h1", false); }
function sendReminderS()  { return sendReminderAll_("s", false); }

function sendReminderAll_(key, testOnly) {
  var results = {};
  for (var type in EVENTS) results[type] = sendReminder_(EVENTS[type], key, testOnly);
  return results;
}

/** testOnly=true ならテスト行（TEST_NAME＋TEST_EMAIL）だけに送る */
function sendReminder_(cfg, key, testOnly) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = eventSheet_(ss, cfg);
  var col = reminderCol_(cfg, key);
  if (!col) return { result: "unknown-key" };
  var width = headersOf_(cfg).length;
  var lastRow = sheet.getLastRow();
  var pending = [], seen = {}, skippedSent = 0, skippedDup = 0;
  for (var r = 2; r <= lastRow; r++) {
    var v = sheet.getRange(r, 1, 1, width).getValues()[0];
    var name = String(v[1] || ""), email = String(v[2] || "").trim();
    if (!email) continue;
    var isTest = (name === TEST_NAME && email === TEST_EMAIL);
    if (testOnly && !isTest) continue;
    if (!testOnly && isTest) continue;
    if (String(v[col - 1]).indexOf("送信OK") === 0) { skippedSent++; continue; }
    var k = email.toLowerCase();
    if (seen[k]) { sheet.getRange(r, col).setValue("重複（同じ宛先に送信済み）"); skippedDup++; continue; }
    seen[k] = true;
    pending.push({ row: r, name: name, email: email });
  }

  var remaining = MailApp.getRemainingDailyQuota();
  if (pending.length > remaining) {
    console.error("リマインド中止：送信枠不足 sheet=" + cfg.sheet + " pending=" + pending.length + " remaining=" + remaining);
    return { result: "abort", reason: "送信枠が足りないので送らなかった", key: key, pending: pending.length, remaining: remaining };
  }

  var sent = 0, failed = 0, subject = reminderSubject_(key);
  for (var i = 0; i < pending.length; i++) {
    var p = pending[i];
    try {
      var via = sendMail_(p.email, subject, reminderBody_(p.name, key));
      sheet.getRange(p.row, col).setValue("送信OK " + Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss") + "（" + via + "）");
      sent++;
    } catch (err) {
      sheet.getRange(p.row, col).setValue("送信ERR: " + err);
      failed++;
      console.error("リマインド打ち切り sheet=" + cfg.sheet + " row=" + p.row + " " + err);
      break;
    }
  }
  return { result: "ok", key: key, testOnly: !!testOnly, sent: sent, failed: failed, skippedSent: skippedSent, skippedDup: skippedDup, remainingBefore: remaining };
}

/** 6つの時間トリガーを作る（旧2本・同名は削除して作り直す＝重複しない） */
function setupTriggers() {
  var handlers = { d3: "sendReminderD3", d2: "sendReminderD2", d1: "sendReminderD1", m: "sendReminderM", h1: "sendReminderH1", s: "sendReminderS" };
  var targets = Object.keys(handlers).map(function (k) { return handlers[k]; }).concat(["sendReminderDayBefore", "sendReminderSameDay"]);
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (targets.indexOf(t.getHandlerFunction()) !== -1) { ScriptApp.deleteTrigger(t); removed++; }
  });
  var created = [];
  REMINDERS.forEach(function (r) {
    ScriptApp.newTrigger(handlers[r.key]).timeBased().at(new Date(r.at)).create();
    created.push(handlers[r.key] + " " + r.at);
  });
  return { result: "ok", removedOld: removed, created: created };
}

/** Gmail(GAS)のリマインドを止める：6つの時間トリガーを削除。Flodeskのテスト完了後に1回だけ実行 */
function stopGmailReminders() {
  var names = ["sendReminderD3", "sendReminderD2", "sendReminderD1", "sendReminderM", "sendReminderH1", "sendReminderS", "sendReminderDayBefore", "sendReminderSameDay"];
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) { if (names.indexOf(t.getHandlerFunction()) !== -1) { ScriptApp.deleteTrigger(t); removed++; } });
  return { result: "ok", removedTriggers: removed };
}

/* ============ 保守用GET（個人情報は返さない） ============ */
function admin_(p) {
  var action = String(p.action || "");
  if (action === "setupTriggers") { if (String(p.key || "") !== ADMIN_KEY) return { result: "denied" }; return setupTriggers(); }
  if (action === "stopGmailReminders") { if (String(p.key || "") !== ADMIN_KEY) return { result: "denied" }; return stopGmailReminders(); }
  if (action === "backfillFlodesk") { if (String(p.key || "") !== ADMIN_KEY) return { result: "denied" }; return backfillFlodesk_(); }
  if (action === "flodeskTest") { if (String(p.key || "") !== ADMIN_KEY) return { result: "denied" }; try { return { result: flodeskAdd_(TEST_EMAIL, TEST_NAME, true) }; } catch (e) { return { result: "ERR " + e }; } }
  var cfg = EVENTS[p.type] || EVENTS["応援の循環キックオフ"];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = eventSheet_(ss, cfg);
  var lastRow = sheet.getLastRow();
  var width = headersOf_(cfg).length;
  if (action === "inspect") {
    var last = null, sent = {};
    REMINDERS.forEach(function (r) { sent[r.key] = 0; });
    var tests = 0;
    if (lastRow >= 2) {
      var st = sheet.getRange(lastRow, cfg.base.length - 1, 1, 2 + REMINDERS.length).getValues()[0];
      last = { notify: String(st[0]).slice(0, 40), reply: String(st[1]).slice(0, 40) };
    }
    for (var r = 2; r <= lastRow; r++) {
      var v = sheet.getRange(r, 1, 1, width).getValues()[0];
      if (String(v[1]) === TEST_NAME && String(v[2]) === TEST_EMAIL) tests++;
      REMINDERS.forEach(function (rm) { if (String(v[reminderCol_(cfg, rm.key) - 1]).indexOf("送信OK") === 0) sent[rm.key]++; });
    }
    return { sheet: cfg.sheet, dataRows: Math.max(lastRow - 1, 0), headers: sheet.getRange(1, 1, 1, width).getValues()[0],
             tz: Session.getScriptTimeZone(), gmailRemaining: MailApp.getRemainingDailyQuota(), flodeskKey: !!flodeskKey_(),
             triggers: ScriptApp.getProjectTriggers().length, sent: sent, tests: tests, lastRowStatus: last };
  }
  if (action === "clearTest") {
    var removed = 0;
    for (var r2 = lastRow; r2 >= 2; r2--) {
      var bc = sheet.getRange(r2, 2, 1, 2).getValues()[0];
      if (bc[0] === TEST_NAME && bc[1] === TEST_EMAIL) { sheet.deleteRow(r2); removed++; }
    }
    return { result: "ok", removed: removed };
  }
  if (action === "remind") {
    if (String(p.key || "") !== ADMIN_KEY) return { result: "denied" };
    return sendReminder_(cfg, String(p.which || ""), String(p.test || "") === "1");
  }
  return { result: "unknown-action" };
}
