/**
 * 「応援の循環」キックオフセミナー（2026年10月2日（金）20:00〜21:30 JST）
 * ── 申込管理シート連携 + 自動返信メール + リマインド
 *
 * 【セットアップ手順】
 * 1. Googleスプレッドシートを新規作成する（例：「応援の循環キックオフ_申込管理」）
 * 2. そのシートで「拡張機能 → Apps Script」を開き、このコードを全部貼り付ける
 * 3. 「デプロイ → 新しいデプロイ → 種類：ウェブアプリ」
 *      - 次のユーザーとして実行：自分
 *      - アクセスできるユーザー：全員
 * 4. 発行された「ウェブアプリURL」をコピーし、index.html 内の
 *      const GAS_URL = "PASTE_GAS_WEBAPP_URL_HERE";
 *    に貼り付けて再デプロイする
 * 5. Apps Script の「トリガー」画面で「時間主導型 → 特定の日時」を2つ設定する
 *      - sendReminderDayBefore … 10月1日（木）20:00頃（日本時間）
 *      - sendReminderSameDay   … 10月2日（金）17:00頃（日本時間）
 *    （トリガーのタイムゾーンは Apps Script の既定タイムゾーンに従う。
 *      ?action=inspect の tz で確認できる）
 *
 * 【注意】再デプロイのときは必ず「新バージョン」を選ぶこと。
 *         「アクセスできるユーザー：全員」でないとLPから叩けない。
 *
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

const EVENT = {
  SHEET: "応援の循環キックオフ申込",
  HEADERS: ["申込日時", "お名前", "メールアドレス", "BAND参加状況", "高田さんへの質問", "通知結果", "返信結果", "前日リマインド", "当日リマインド"],
  MEET: "https://meet.google.com/rjz-buhr-eyo",
  WHEN: "10月2日（金）20:00〜21:30（日本時間）",
  BAND_URL: "https://band.us/n/a0a6b5M8hfTc8",
  ADMIN_KEY: "zICLZ91dIQA5ZjFN200sAjhk"   // GETからリマインドを手動実行するときの合言葉（トリガー実行には不要）
};
const COL_DAYBEFORE = 8;  // H列
const COL_SAMEDAY = 9;    // I列
const TEST_NAME = "テスト（アーニャ）";
const TEST_EMAIL = "organiclifeingermany@gmail.com";

function eventSheet_(ss) {
  var sh = ss.getSheetByName(EVENT.SHEET);
  if (!sh) {
    sh = ss.insertSheet(EVENT.SHEET);
    sh.appendRow(EVENT.HEADERS);
    sh.getRange(1, 1, 1, EVENT.HEADERS.length).setFontWeight("bold");
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

/** 申込直後の自動返信（Meetリンク・日時。BAND未参加の場合はBAND案内も添える） */
function applyReplyBody(d) {
  var bandBlock = "";
  if (d.bandStatus !== "参加済み") {
    bandBlock =
      "■ この90分の続きは、BANDで\n\n" +
      "「IOB 応援の循環 2026」は、IOBで学んだ仲間が近況を知り、つながり、応援を渡し合っていくための場所です。\n" +
      "よかったら、この機会にBANDにも加わってください。\n" +
      EVENT.BAND_URL + "\n\n\n";
  }
  return d.name + " 様\n\n" +
    "IOBオーガニックスクール事務局です。\n" +
    "「応援の循環」キックオフセミナーにお申し込みいただき、ありがとうございます。\n\n" +
    "■ 日時\n" +
    EVENT.WHEN + "\n\n" +
    "■ 参加用リンク（Google Meet）\n" +
    EVENT.MEET + "\n" +
    "時間になったら、このリンクを開いてください。\n" +
    "カメラは切ったままでも、聞いているだけでも大丈夫です。\n\n" +
    "■ 当日やること（90分）\n" +
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
    "■ お申し込み内容（控え）\n" +
    "BAND参加状況：" + (d.bandStatus || "") + "\n" +
    "高田さんへの質問：" + (d.question || "（なし）") + "\n\n" +
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
      "「応援の循環」キックオフセミナーは、このあと20:00からです（日本時間・21:30ごろまで）。\n\n" +
      "▼参加用リンク（Google Meet）\n" +
      EVENT.MEET + "\n\n" +
      "何も持ってこなくて大丈夫です。カメラは切ったままでも、聞いているだけでも。\n" +
      "物販・商品のご案内は一切ありません。\n\n" +
      "のちほど、画面の向こうでお会いしましょう。\n\n" +
      signature_();
  }
  return name + " 様\n\n" +
    "IOBオーガニックスクール事務局です。\n\n" +
    "「応援の循環」キックオフセミナーは、明日10月2日（金）20:00からです（日本時間・21:30ごろまで）。\n\n" +
    "▼参加用リンク（Google Meet）\n" +
    EVENT.MEET + "\n\n" +
    "時間になったら、このリンクを開いてください。\n" +
    "カメラは切ったままでも、聞いているだけでも大丈夫です。\n\n" +
    "明日は、高田洋平さんと一緒に「応援が巡る関係」を実際にワークしながら体験します。\n" +
    "物販・商品のご案内は一切ありません。\n\n" +
    "明日の夜、画面の向こうでお会いできたら嬉しいです。\n\n" +
    signature_();
}

/** LPからの申込：シート追記 → 事務局へ通知 → 申込者へ自動返信 */
function applyEvent_(ss, d) {
  var sheet = eventSheet_(ss);
  var ts = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
  sheet.appendRow([ts, d.name || "", d.email || "", d.bandStatus || "", d.question || "", "", "", "", ""]);
  var row = sheet.getLastRow();
  var notifyStatus = "", replyStatus = "";

  if (CONFIG.NOTIFY_TO) {
    try {
      MailApp.sendEmail({
        to: CONFIG.NOTIFY_TO,
        subject: "【応援の循環キックオフ申込】" + d.name + " 様",
        body: "「応援の循環」キックオフセミナーの申込が入りました。\n\n" +
          "お名前：" + d.name + "\n" +
          "メール：" + d.email + "\n" +
          "BAND参加状況：" + (d.bandStatus || "") + "\n" +
          "高田さんへの質問：" + (d.question || "（なし）") + "\n" +
          "日時：" + ts
      });
      notifyStatus = "通知OK " + ts;
    } catch (e1) {
      notifyStatus = "通知ERR: " + e1;
    }
    sheet.getRange(row, 6).setValue(notifyStatus);
  }

  if (d.email) {
    try {
      var opts = mailOpts_();
      GmailApp.sendEmail(d.email, "【お申し込みありがとうございます】「応援の循環」キックオフセミナーのご案内（Google Meetリンク）", applyReplyBody(d), opts);
      replyStatus = "返信OK " + ts;
    } catch (e2) {
      replyStatus = "返信ERR: " + e2;
    }
    sheet.getRange(row, 7).setValue(replyStatus);
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return ContentService.createTextOutput(JSON.stringify(applyEvent_(ss, d)))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ result: "error", message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/** トリガーから呼ぶ2本（引数なし） */
function sendReminderDayBefore() { return sendReminder_("dayBefore", false); }
function sendReminderSameDay()   { return sendReminder_("sameDay", false); }

/** リマインド送信本体。testOnly=true ならテスト行（TEST_NAME＋TEST_EMAIL）だけに送る */
function sendReminder_(which, testOnly) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = eventSheet_(ss);
  var col = (which === "sameDay") ? COL_SAMEDAY : COL_DAYBEFORE;
  var lastRow = sheet.getLastRow();
  var pending = [], seen = {}, skippedSent = 0, skippedDup = 0;
  for (var r = 2; r <= lastRow; r++) {
    var v = sheet.getRange(r, 1, 1, EVENT.HEADERS.length).getValues()[0];
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
    console.error("応援の循環リマインド中止：送信枠不足 pending=" + pending.length + " remaining=" + remaining);
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
      console.error("応援の循環リマインド打ち切り row=" + p.row + " " + err);
      break;                                                 // ③ 例外が出たらその回は打ち切る（次回実行で未送信だけ送る）
    }
  }
  return { result: "ok", which: which, testOnly: !!testOnly, sent: sent, failed: failed, skippedSent: skippedSent, skippedDup: skippedDup, remainingBefore: remaining };
}

/** 保守用GET（個人情報は返さない） */
function admin_(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = eventSheet_(ss);
  var lastRow = sheet.getLastRow();
  var action = String(p.action || "");
  if (action === "inspect") {
    var last = null;
    if (lastRow >= 2) {
      var st = sheet.getRange(lastRow, 6, 1, 4).getValues()[0];
      last = { notify: String(st[0]).slice(0, 40), reply: String(st[1]).slice(0, 40), dayBefore: String(st[2]).slice(0, 40), sameDay: String(st[3]).slice(0, 40) };
    }
    var cnt = { sentDayBefore: 0, sentSameDay: 0, tests: 0 };
    for (var r = 2; r <= lastRow; r++) {
      var v = sheet.getRange(r, 1, 1, EVENT.HEADERS.length).getValues()[0];
      if (String(v[1]) === TEST_NAME && String(v[2]) === TEST_EMAIL) cnt.tests++;
      if (String(v[7]).indexOf("送信OK") === 0) cnt.sentDayBefore++;
      if (String(v[8]).indexOf("送信OK") === 0) cnt.sentSameDay++;
    }
    return { sheet: EVENT.SHEET, dataRows: Math.max(lastRow - 1, 0), headers: sheet.getRange(1, 1, 1, EVENT.HEADERS.length).getValues()[0],
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
    if (String(p.key || "") !== EVENT.ADMIN_KEY) return { result: "denied" };
    var which = (p.which === "sameDay") ? "sameDay" : "dayBefore";
    return sendReminder_(which, String(p.test || "") === "1");
  }
  return { result: "unknown-action" };
}
