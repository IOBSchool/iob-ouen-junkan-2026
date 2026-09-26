/**
 * 「応援の循環」キックオフセミナー（2026年10月2日（金）20:00〜22:00 JST）
 * ── 申込管理シート連携 + 自動返信メール + リマインド
 * ── 既存生向け（IOB専門家コース・コスメ専門家コース）と一般向け（オーガニックをつくる人・届ける人・選ぶ人）の
 *    2つのLPから同じGASへ送信され、type で振り分けて別タブに記録する
 *
 * 【セットアップ手順】
 * 1. Googleスプレッドシートを新規作成する（例：「応援の循環キックオフ_申込管理」）
 * 2. そのシートで「拡張機能 → Apps Script」を開き、このコードを全部貼り付ける
 * 3. 「デプロイ → 新しいデプロイ → 種類：ウェブアプリ」
 *      - 次のユーザーとして実行：自分
 *      - アクセスできるユーザー：全員
 * 4. 発行された「ウェブアプリURL」をコピーし、index.html と general.html 両方の
 *      const GAS_URL = "PASTE_GAS_WEBAPP_URL_HERE";
 *    に貼り付けて再デプロイする
 * 5. Apps Script の「トリガー」画面で「時間主導型 → 特定の日時」を2つ設定する
 *      - sendReminderDayBefore … 10月1日（木）20:00頃（日本時間）
 *      - sendReminderSameDay   … 10月2日（金）17:00頃（日本時間）
 *    （トリガーのタイムゾーンは Apps Script の既定タイムゾーンに従う。
 *      ?action=inspect の tz で確認できる。この2本は既存生向け・一般向け両方のシートに送る）
 *
 * 【注意】再デプロイのときは必ず「新バージョン」を選ぶこと。
 *         「アクセスできるユーザー：全員」でないとLPから叩けない。
 *
 * 🚨有料化・決済導線は一切なし（2026-09-26 なつこさん最終確認：ライブもアーカイブも、
 *    既存生向けも一般向けも、すべて無料。「シェアしたら無料」等の条件付けもしない）
 * 🚨一斉送信の4原則（過去の事故の再発防止）：
 *   ①送る前に残数確認 ②1人ずつ送信済みを記録し、再実行時は未送信だけ送る
 *   ③例外が出たらその回は打ち切る ④件数が多い場合は別途相談する
 * 🚨個人情報をAIに返さない：inspect は件数・見出し・最終行の送信結果だけを返す
 */

const CONFIG = {
  SENDER_NAME: "IOBオーガニックスクール事務局",
  SENDER_EMAIL: "school@iob.bio",   // Gmailの送信元エイリアスに登録済みの場合のみ反映される
  NOTIFY_TO: "school@iob.bio"       // 申込通知の受信先（カンマ区切りで複数可）
};

const MEET = "https://meet.google.com/rjz-buhr-eyo";
const WHEN = "10月2日（金）20:00〜22:00（日本時間）";
const BAND_URL = "https://band.us/n/a0a6b5M8hfTc8";
const ADMIN_KEY = "zICLZ91dIQA5ZjFN200sAjhk";   // GETからリマインドを手動実行するときの合言葉（トリガー実行には不要）
const TEST_NAME = "テスト（アーニャ）";
const TEST_EMAIL = "organiclifeingermany@gmail.com";

/** 既存生向け・一般向けのシート設定。BAND参加状況の列は既存生向けだけにある */
const EVENTS = {
  "応援の循環キックオフ": {
    sheet: "応援の循環キックオフ申込",
    headers: ["申込日時", "お名前", "メールアドレス", "BAND参加状況", "高田さんへの質問", "通知結果", "返信結果", "前日リマインド", "当日リマインド"],
    hasBand: true,
    colDayBefore: 8, // H列
    colSameDay: 9    // I列
  },
  "応援の循環キックオフ_一般": {
    sheet: "応援の循環キックオフ申込_一般",
    headers: ["申込日時", "お名前", "メールアドレス", "高田さんへの質問", "通知結果", "返信結果", "前日リマインド", "当日リマインド"],
    hasBand: false,
    colDayBefore: 7, // G列
    colSameDay: 8    // H列
  }
};

function eventSheet_(ss, cfg) {
  var sh = ss.getSheetByName(cfg.sheet);
  if (!sh) {
    sh = ss.insertSheet(cfg.sheet);
    sh.appendRow(cfg.headers);
    sh.getRange(1, 1, 1, cfg.headers.length).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
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

/** 送信オプション（差出人名＋school@iob.bio がGmailのエイリアスにあれば from も設定） */
function mailOpts_() {
  var opts = { name: CONFIG.SENDER_NAME };
  try {
    if (GmailApp.getAliases().indexOf(CONFIG.SENDER_EMAIL) !== -1) opts.from = CONFIG.SENDER_EMAIL;
  } catch (e) {}
  return opts;
}

/** 申込直後の自動返信（Meetリンク・日時）。既存生向けはBAND未参加者にBAND案内を添える。一般向けは既存生限定BANDへは案内しない */
function applyReplyBody(cfg, d) {
  var bandBlock = "";
  if (cfg.hasBand && d.bandStatus !== "参加済み") {
    bandBlock =
      "■ この2時間の続きは、BANDで\n\n" +
      "「IOB 応援の循環 2026」は、IOBで学んだ仲間が近況を知り、つながり、応援を渡し合っていくための場所です。\n" +
      "よかったら、この機会にBANDにも加わってください。\n" +
      BAND_URL + "\n\n\n";
  }
  var recapBlock = cfg.hasBand
    ? "■ お申し込み内容（控え）\n" +
      "BAND参加状況：" + (d.bandStatus || "") + "\n" +
      "高田さんへの質問：" + (d.question || "（なし）") + "\n\n"
    : "■ お申し込み内容（控え）\n" +
      "高田さんへの質問：" + (d.question || "（なし）") + "\n\n";

  return d.name + " 様\n\n" +
    "IOBオーガニックスクール事務局です。\n" +
    "「応援の循環」キックオフセミナーにお申し込みいただき、ありがとうございます。\n\n" +
    "■ 日時\n" +
    WHEN + "\n\n" +
    "■ 参加用リンク（Google Meet）\n" +
    MEET + "\n" +
    "時間になったら、このリンクを開いてください。\n" +
    "カメラは切ったままでも、聞いているだけでも大丈夫です。\n\n" +
    "■ 当日やること（2時間）\n" +
    "・オープニング\n" +
    "・Connect（自分自身の「応援の共体験」を思い出すワーク）\n" +
    "・高田洋平さんトーク\n" +
    "・Commonワーク（実際にみんなで考え、言葉にする）\n" +
    "・共有\n" +
    "・クロージング\n\n" +
    "物販・商品のご案内は一切ありません。\n\n\n" +
    bandBlock +
    "■ リマインド\n" +
    "前日と当日にも、このメールアドレス宛にリマインドをお送りします。\n\n" +
    recapBlock +
    "何も持ってこなくて大丈夫です。\n" +
    "当日、画面の向こうでお会いできたら嬉しいです。\n\n" +
    "ご不明な点は school@iob.bio までご連絡ください。\n\n" +
    signature_();
}

/** リマインドの件名・本文（which = "dayBefore" | "sameDay"）*/
function reminderSubject_(which) {
  return which === "sameDay"
    ? "このあと20時からです｜「応援の循環」キックオフセミナー（Google Meetリンク）"
    : "明日20時です｜「応援の循環」キックオフセミナーのご案内";
}
function reminderBody_(name, which) {
  if (which === "sameDay") {
    return name + " 様\n\n" +
      "IOBオーガニックスクール事務局です。\n\n" +
      "「応援の循環」キックオフセミナーは、このあと20:00からです（日本時間・22:00ごろまで）。\n\n" +
      "▼参加用リンク（Google Meet）\n" +
      MEET + "\n\n" +
      "何も持ってこなくて大丈夫です。カメラは切ったままでも、聞いているだけでも。\n" +
      "物販・商品のご案内は一切ありません。\n\n" +
      "のちほど、画面の向こうでお会いしましょう。\n\n" +
      signature_();
  }
  return name + " 様\n\n" +
    "IOBオーガニックスクール事務局です。\n\n" +
    "「応援の循環」キックオフセミナーは、明日10月2日（金）20:00からです（日本時間・22:00ごろまで）。\n\n" +
    "▼参加用リンク（Google Meet）\n" +
    MEET + "\n\n" +
    "時間になったら、このリンクを開いてください。\n" +
    "カメラは切ったままでも、聞いているだけでも大丈夫です。\n\n" +
    "明日は、高田洋平さんと一緒に「応援が巡る関係」を実際にワークしながら体験します。\n" +
    "物販・商品のご案内は一切ありません。\n\n" +
    "明日の夜、画面の向こうでお会いできたら嬉しいです。\n\n" +
    signature_();
}

/** LPからの申込：シート追記 → 事務局へ通知 → 申込者へ自動返信 */
function applyEvent_(ss, cfg, d) {
  var sheet = eventSheet_(ss, cfg);
  var ts = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
  var row = cfg.hasBand
    ? [ts, d.name || "", d.email || "", d.bandStatus || "", d.question || "", "", "", "", ""]
    : [ts, d.name || "", d.email || "", d.question || "", "", "", "", ""];
  sheet.appendRow(row);
  var rowNum = sheet.getLastRow();
  var notifyCol = cfg.hasBand ? 6 : 5;
  var replyCol = cfg.hasBand ? 7 : 6;
  var notifyStatus = "", replyStatus = "";

  if (CONFIG.NOTIFY_TO) {
    try {
      MailApp.sendEmail({
        to: CONFIG.NOTIFY_TO,
        subject: "【応援の循環キックオフ申込】" + d.name + " 様" + (cfg.hasBand ? "" : "（一般）"),
        body: "「応援の循環」キックオフセミナーの申込が入りました。\n\n" +
          "お名前：" + d.name + "\n" +
          "メール：" + d.email + "\n" +
          (cfg.hasBand ? "BAND参加状況：" + (d.bandStatus || "") + "\n" : "") +
          "高田さんへの質問：" + (d.question || "（なし）") + "\n" +
          "日時：" + ts
      });
      notifyStatus = "通知OK " + ts;
    } catch (e1) {
      notifyStatus = "通知ERR: " + e1;
    }
    sheet.getRange(rowNum, notifyCol).setValue(notifyStatus);
  }

  if (d.email) {
    try {
      var opts = mailOpts_();
      GmailApp.sendEmail(d.email, "【お申し込みありがとうございます】「応援の循環」キックオフセミナーのご案内（Google Meetリンク）", applyReplyBody(cfg, d), opts);
      replyStatus = "返信OK " + ts;
    } catch (e2) {
      replyStatus = "返信ERR: " + e2;
    }
    sheet.getRange(rowNum, replyCol).setValue(replyStatus);
  }
  return { result: "ok" };
}

/** POST・GETどちらで来ても同じ処理に流す（302リダイレクト対策） */
function doPost(e) { return handle(e); }
function doGet(e)  { return handle(e); }

function handle(e) {
  try {
    // 保守用GET（個人情報は返さない）
    if (e && e.parameter && e.parameter.action) {
      return ContentService.createTextOutput(JSON.stringify(admin_(e.parameter)))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // POSTボディ・GETパラメータの両対応
    // 🚨 e.parameter.payload を先に見ること。
    // LP側は Content-Type: application/x-www-form-urlencoded で送るため、
    // その場合 e.postData.contents は "payload=%7B...%7D" という生の文字列になり、
    // 先に読むと JSON.parse が必ず失敗して申込が保存されない。
    var raw = "";
    if (e && e.parameter && e.parameter.payload) {
      raw = e.parameter.payload;
    } else if (e && e.postData && e.postData.contents) {
      raw = e.postData.contents;
    }
    if (!raw) {
      return ContentService.createTextOutput(JSON.stringify({ result: "no-data" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const d = JSON.parse(raw);
    const cfg = EVENTS[d.type];
    if (!cfg) {
      return ContentService.createTextOutput(JSON.stringify({ result: "unknown-type" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return ContentService.createTextOutput(JSON.stringify(applyEvent_(ss, cfg, d)))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ result: "error", message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/** トリガーから呼ぶ2本（引数なし）。既存生向け・一般向け両方のシートに送る */
function sendReminderDayBefore() { return sendReminderAll_("dayBefore", false); }
function sendReminderSameDay()   { return sendReminderAll_("sameDay", false); }

function sendReminderAll_(which, testOnly) {
  var results = {};
  for (var type in EVENTS) {
    results[type] = sendReminder_(EVENTS[type], which, testOnly);
  }
  return results;
}

/** リマインド送信本体。testOnly=true ならテスト行（TEST_NAME＋TEST_EMAIL）だけに送る */
function sendReminder_(cfg, which, testOnly) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = eventSheet_(ss, cfg);
  var col = (which === "sameDay") ? cfg.colSameDay : cfg.colDayBefore;
  var lastRow = sheet.getLastRow();
  var pending = [], seen = {}, skippedSent = 0, skippedDup = 0;
  for (var r = 2; r <= lastRow; r++) {
    var v = sheet.getRange(r, 1, 1, cfg.headers.length).getValues()[0];
    var name = String(v[1] || ""), email = String(v[2] || "").trim();
    if (!email) continue;
    var isTest = (name === TEST_NAME && email === TEST_EMAIL);
    if (testOnly && !isTest) continue;
    if (!testOnly && isTest) continue;                      // 本番ではテスト行に送らない
    if (String(v[col - 1]).indexOf("送信OK") === 0) { skippedSent++; continue; }   // ② 送信済みは飛ばす
    var k = email.toLowerCase();
    if (seen[k]) { sheet.getRange(r, col).setValue("重複（同じ宛先に送信済み） " + Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm")); skippedDup++; continue; }
    seen[k] = true;
    pending.push({ row: r, name: name, email: email });
  }

  // ① 残数確認
  var remaining = MailApp.getRemainingDailyQuota();
  if (pending.length > remaining) {
    console.error("応援の循環リマインド中止：送信枠不足 sheet=" + cfg.sheet + " pending=" + pending.length + " remaining=" + remaining);
    return { result: "abort", reason: "送信枠が足りないので送らなかった", which: which, pending: pending.length, remaining: remaining };
  }

  var sent = 0, failed = 0;
  var subject = reminderSubject_(which);
  var opts = mailOpts_();
  for (var i = 0; i < pending.length; i++) {
    var p = pending[i];
    try {
      GmailApp.sendEmail(p.email, subject, reminderBody_(p.name, which), opts);
      sheet.getRange(p.row, col).setValue("送信OK " + Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss"));
      sent++;
    } catch (err) {
      sheet.getRange(p.row, col).setValue("送信ERR: " + err);
      failed++;
      console.error("応援の循環リマインド打ち切り sheet=" + cfg.sheet + " row=" + p.row + " " + err);
      break;                                                 // ③ 例外が出たらその回は打ち切る（次回実行で未送信だけ送る）
    }
  }
  return { result: "ok", which: which, testOnly: !!testOnly, sent: sent, failed: failed, skippedSent: skippedSent, skippedDup: skippedDup, remainingBefore: remaining };
}

/** 保守用GET（個人情報は返さない）。type で既存生向け・一般向けを切り替える（省略時は既存生向け） */
function admin_(p) {
  var cfg = EVENTS[p.type] || EVENTS["応援の循環キックオフ"];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = eventSheet_(ss, cfg);
  var lastRow = sheet.getLastRow();
  var action = String(p.action || "");
  var notifyCol = cfg.hasBand ? 6 : 5;
  if (action === "inspect") {
    var last = null;
    if (lastRow >= 2) {
      var st = sheet.getRange(lastRow, notifyCol, 1, 4).getValues()[0];
      last = { notify: String(st[0]).slice(0, 40), reply: String(st[1]).slice(0, 40), dayBefore: String(st[2]).slice(0, 40), sameDay: String(st[3]).slice(0, 40) };
    }
    var cnt = { sentDayBefore: 0, sentSameDay: 0, tests: 0 };
    for (var r = 2; r <= lastRow; r++) {
      var v = sheet.getRange(r, 1, 1, cfg.headers.length).getValues()[0];
      if (String(v[1]) === TEST_NAME && String(v[2]) === TEST_EMAIL) cnt.tests++;
      if (String(v[cfg.colDayBefore - 1]).indexOf("送信OK") === 0) cnt.sentDayBefore++;
      if (String(v[cfg.colSameDay - 1]).indexOf("送信OK") === 0) cnt.sentSameDay++;
    }
    return { sheet: cfg.sheet, dataRows: Math.max(lastRow - 1, 0), headers: sheet.getRange(1, 1, 1, cfg.headers.length).getValues()[0],
             tz: Session.getScriptTimeZone(), remaining: MailApp.getRemainingDailyQuota(), counts: cnt, lastRowStatus: last };
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
    var which = (p.which === "sameDay") ? "sameDay" : "dayBefore";
    return sendReminder_(cfg, which, String(p.test || "") === "1");
  }
  if (action === "setupTriggers") {
    if (String(p.key || "") !== ADMIN_KEY) return { result: "denied" };
    return setupTriggers();
  }
  return { result: "unknown-action" };
}

/** 前日10/1(木)20:00・当日10/2(金)17:00(日本時間)の時間主導型トリガーを作成する。
 *  同名トリガーが既にあれば一旦削除してから作り直す(重複実行防止)。
 *  実行方法：Apps Scriptエディタで関数選択→setupTriggers→実行、
 *  または ?action=setupTriggers&key=(ADMIN_KEY) をGETで叩く */
function setupTriggers() {
  var targets = ["sendReminderDayBefore", "sendReminderSameDay"];
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (targets.indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  ScriptApp.newTrigger("sendReminderDayBefore")
    .timeBased()
    .at(new Date("2026-10-01T20:00:00+09:00"))
    .create();
  ScriptApp.newTrigger("sendReminderSameDay")
    .timeBased()
    .at(new Date("2026-10-02T17:00:00+09:00"))
    .create();
  return { result: "ok", removedOld: removed, created: targets };
}
