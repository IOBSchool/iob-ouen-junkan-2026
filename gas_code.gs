/**
 * 「応援の循環」キックオフセミナー（2026年10月2日（金）20:00〜22:00 JST）
 * ── 申込管理シート連携 + 自動返信メール + リマインド6通
 * ── 既存生向け（IOB専門家コース・コスメ専門家コース）と一般向けの2つのLPから同じGASへ送信され、
 *    type で振り分けて別タブに記録する
 *
 * 【リマインド（日本時間・6回）】
 *   9/29(火)20:00 あと3日 ／ 9/30(水)20:00 あと2日 ／ 10/1(木)20:00 明日
 *   10/2(金)08:00 今夜 ／ 19:00 あと1時間 ／ 20:00 始まりました
 *   → setupTriggers を1回実行すると6つの時間トリガーが作られる（何度実行しても重複しない）
 *
 * 【送信経路】
 *   スクリプトプロパティ BREVO_API_KEY と BREVO_ENABLED=true があればBrevo(1日300通)で送り、
 *   なければ／失敗したらGmail(1日100通)で送る。キーの登録はなつこさん本人が行う（AIは触らない）。
 *   事務局宛の申込通知だけは常にGmail直送（Brevo→iob.bio宛が届かなかった前例があるため）。
 *
 * 🚨有料化・決済導線は一切なし（ライブもアーカイブも、既存生向けも一般向けも無料）
 * 🚨一斉送信の4原則：①送る前に残数確認 ②1人ずつ送信済みを記録し再実行時は未送信だけ
 *   ③例外が出たらその回は打ち切る ④件数が多い場合はBrevoへ
 * 🚨個人情報をAIに返さない：inspect は件数・見出し・最終行の送信結果だけを返す
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
function headersOf_(cfg) { return cfg.base.concat(REMINDERS.map(function (r) { return r.label; })); }
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

/* ============ メール送信（Brevo優先・失敗時Gmail） ============ */
function brevoOn_() {
  var p = PropertiesService.getScriptProperties();
  return !!p.getProperty("BREVO_API_KEY") && p.getProperty("BREVO_ENABLED") === "true";
}
function gmailOpts_() {
  var opts = { name: CONFIG.SENDER_NAME };
  try { if (GmailApp.getAliases().indexOf(CONFIG.SENDER_EMAIL) !== -1) opts.from = CONFIG.SENDER_EMAIL; } catch (e) {}
  return opts;
}
/** 戻り値は "brevo" か "gmail"。🚨Brevoは201でも送信元未認証だと裏で弾く → 成功判定はBrevo Logsで */
function sendMail_(to, subject, body, forceBrevo) {
  var key = PropertiesService.getScriptProperties().getProperty("BREVO_API_KEY");
  if (key && (forceBrevo || brevoOn_())) {
    try {
      var res = UrlFetchApp.fetch("https://api.brevo.com/v3/smtp/email", {
        method: "post", contentType: "application/json",
        headers: { "api-key": key, accept: "application/json" },
        payload: JSON.stringify({ sender: { email: CONFIG.SENDER_EMAIL, name: CONFIG.SENDER_NAME }, to: [{ email: to }], subject: subject, textContent: body }),
        muteHttpExceptions: true
      });
      var code = res.getResponseCode();
      if (code === 201 || code === 202) return "brevo";
      console.error("Brevo失敗 HTTP " + code + " → Gmailで送信");
    } catch (err) { console.error("Brevoエラー " + err + " → Gmailで送信"); }
  }
  GmailApp.sendEmail(to, subject, body, gmailOpts_());
  return "gmail";
}
/** エディタから実行：Brevo経由で自分宛にテスト送信（BREVO_API_KEY登録後） */
function testBrevoSend() {
  Logger.log("送信経路: " + sendMail_(TEST_EMAIL, "【テスト】Brevo経由の送信確認（応援の循環GAS）", "Brevo経由の送信テストです。\n" + new Date(), true));
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
  var ts = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
  var vals = cfg.hasBand
    ? [ts, d.name || "", d.email || "", d.bandStatus || "", d.question || "", "", ""]
    : [ts, d.name || "", d.email || "", d.question || "", "", ""];
  for (var i = 0; i < REMINDERS.length; i++) vals.push("");
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

  var brevo = brevoOn_();
  var remaining = brevo ? 300 : MailApp.getRemainingDailyQuota();
  if (pending.length > remaining) {
    console.error("リマインド中止：送信枠不足 sheet=" + cfg.sheet + " pending=" + pending.length + " remaining=" + remaining);
    return { result: "abort", reason: "送信枠が足りないので送らなかった", key: key, pending: pending.length, remaining: remaining, brevo: brevo };
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
  return { result: "ok", key: key, testOnly: !!testOnly, sent: sent, failed: failed, skippedSent: skippedSent, skippedDup: skippedDup, remainingBefore: remaining, brevo: brevo };
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

/** 「Flodesk除外用」タブ：両タブの申込者メールアドレスを1列にまとめる（CSVでダウンロード→Flodeskの申込済みセグメントへ取り込み） */
function setupExportTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tab = ss.getSheetByName("Flodesk除外用") || ss.insertSheet("Flodesk除外用");
  tab.clear();
  tab.getRange("A1").setValue("email");
  var a = "'" + EVENTS["応援の循環キックオフ"].sheet + "'!C2:C";
  var b = "'" + EVENTS["応援の循環キックオフ_一般"].sheet + "'!C2:C";
  tab.getRange("A2").setFormula('=IFERROR(QUERY({' + a + ';' + b + '},"select * where Col1 <> \'\'",0),"")');
  tab.getRange("C1").setValue("使い方：ファイル→ダウンロード→CSV（このタブが開いている状態で）。Flodeskの「応援の循環_申込済み」セグメントに取り込み、告知の配信で「除外」に指定する。新しい申込が入るたびに、この一覧も自動で増える。");
  return { result: "ok", tab: "Flodesk除外用" };
}

/* ============ 保守用GET（個人情報は返さない） ============ */
function admin_(p) {
  var action = String(p.action || "");
  if (action === "setupTriggers") { if (String(p.key || "") !== ADMIN_KEY) return { result: "denied" }; return setupTriggers(); }
  if (action === "setupExportTab") { if (String(p.key || "") !== ADMIN_KEY) return { result: "denied" }; return setupExportTab(); }
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
             tz: Session.getScriptTimeZone(), gmailRemaining: MailApp.getRemainingDailyQuota(), brevo: brevoOn_(),
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
