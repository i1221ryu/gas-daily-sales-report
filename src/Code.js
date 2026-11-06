/**
 * 日次売上レポート自動送信ツール(Google Apps Script)
 *
 * 「売上」シートのデータを日次で集計し、前日比をつけて毎朝メールで送信します。
 * 集計結果は「日次集計」シートにも記録されます。
 */

const SHEET_SALES = '売上';
const SHEET_DAILY = '日次集計';
const SHEET_SETTINGS = '設定';
const TRIGGER_HANDLER = 'sendDailyReport';

// 「売上」シートの列(1始まり)
const COL = { DATE: 1, PRODUCT: 2, CATEGORY: 3, QUANTITY: 4, PRICE: 5, AMOUNT: 6 };

// ---------------------------------------------------------------------------
// メニュー
// ---------------------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 売上レポート')
    .addItem('① 初期設定(シート作成・サンプルデータ投入)', 'setup')
    .addItem('② 今すぐレポートを送信(テスト)', 'sendReportNow')
    .addSeparator()
    .addItem('③ 毎朝8時の自動送信をオンにする', 'installTrigger')
    .addItem('自動送信をオフにする', 'removeTrigger')
    .addToUi();
}

// ---------------------------------------------------------------------------
// メイン処理(トリガーから毎朝実行)
// ---------------------------------------------------------------------------
function sendDailyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = readSettings_(ss);

  try {
    const tz = ss.getSpreadsheetTimeZone();
    const now = new Date();
    const targetKey = dateKey_(addDays_(now, -1), tz); // 昨日
    const prevKey = dateKey_(addDays_(now, -2), tz);   // 一昨日

    const rows = readSalesRows_(ss, tz);
    const summary = buildSummary(rows, targetKey, prevKey);
    console.log('集計完了: %s 売上=%s 件数=%s', targetKey, summary.total, summary.count);

    writeDailyRow_(ss, summary);

    if (summary.count === 0 && !settings.sendWhenEmpty) {
      console.log('%s の売上データが0件のため、メール送信をスキップしました', targetKey);
      return summary;
    }

    MailApp.sendEmail({
      to: settings.to,
      subject: buildSubject(summary, settings.subjectPrefix),
      body: buildTextBody(summary),
      htmlBody: buildHtmlBody(summary, ss.getUrl()),
    });
    console.log('レポートを送信しました: %s', settings.to);
    return summary;
  } catch (e) {
    console.error('レポート作成に失敗しました: %s\n%s', e.message, e.stack);
    notifyError_(settings.to, e, ss.getUrl());
    throw e; // 実行ログにも「失敗」として残す
  }
}

/** メニューからの手動送信 */
function sendReportNow() {
  const ui = SpreadsheetApp.getUi();
  try {
    const summary = sendDailyReport();
    const to = readSettings_(SpreadsheetApp.getActiveSpreadsheet()).to;
    ui.alert(summary.count === 0
      ? summary.date + ' の売上データが0件だったため、送信しませんでした。'
      : to + ' にレポートを送信しました(' + summary.date + ' / ' + formatYen(summary.total) + ')。');
  } catch (e) {
    ui.alert('エラーが発生しました: ' + e.message);
  }
}

// ---------------------------------------------------------------------------
// トリガー設定
// ---------------------------------------------------------------------------
function installTrigger() {
  removeTriggerSilently_();
  ScriptApp.newTrigger(TRIGGER_HANDLER).timeBased().everyDays(1).atHour(8).create();
  SpreadsheetApp.getUi().alert('毎朝8時台の自動送信をオンにしました。');
}

function removeTrigger() {
  const n = removeTriggerSilently_();
  SpreadsheetApp.getUi().alert(n > 0 ? '自動送信をオフにしました。' : '自動送信は設定されていません。');
}

function removeTriggerSilently_() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(t);
      n++;
    }
  });
  return n;
}

// ---------------------------------------------------------------------------
// シートの読み書き
// ---------------------------------------------------------------------------
function readSalesRows_(ss, tz) {
  const sheet = ss.getSheetByName(SHEET_SALES);
  if (!sheet) throw new Error('「' + SHEET_SALES + '」シートが見つかりません。メニューの「初期設定」を実行してください。');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, COL.AMOUNT).getValues();
  const rows = [];
  values.forEach(function (v, i) {
    const date = v[COL.DATE - 1];
    if (date === '' || date === null) return; // 空行
    if (!(date instanceof Date)) {
      console.warn('%s行目: 日付の形式が正しくないためスキップしました (%s)', i + 2, date);
      return;
    }
    const quantity = Number(v[COL.QUANTITY - 1]) || 0;
    const price = Number(v[COL.PRICE - 1]) || 0;
    // 金額列が空なら 数量×単価 で計算
    const amount = v[COL.AMOUNT - 1] === '' ? quantity * price : Number(v[COL.AMOUNT - 1]) || 0;
    rows.push({
      date: dateKey_(date, tz),
      product: String(v[COL.PRODUCT - 1]),
      category: String(v[COL.CATEGORY - 1]),
      quantity: quantity,
      amount: amount,
    });
  });
  return rows;
}

/** 日次集計シートに1行記録(同じ日付があれば上書き) */
function writeDailyRow_(ss, s) {
  const sheet = ss.getSheetByName(SHEET_DAILY) || createDailySheet_(ss);
  const row = [s.date, s.total, s.prevTotal, s.diff, s.rate === null ? '' : s.rate / 100, s.count, s.quantity, new Date()];

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const keys = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().map(function (r) { return r[0]; });
    const idx = keys.indexOf(s.date);
    if (idx >= 0) {
      sheet.getRange(idx + 2, 1, 1, row.length).setValues([row]);
      return;
    }
  }
  sheet.appendRow(row);
}

function readSettings_(ss) {
  const defaults = {
    to: Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail(),
    subjectPrefix: '【日次売上レポート】',
    sendWhenEmpty: false,
  };
  const sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet) return defaults;

  const map = {};
  sheet.getRange(1, 1, sheet.getLastRow() || 1, 2).getValues().forEach(function (r) { map[r[0]] = r[1]; });
  return {
    to: String(map['送信先メールアドレス'] || defaults.to).trim(),
    subjectPrefix: map['件名の先頭'] || defaults.subjectPrefix,
    sendWhenEmpty: map['売上0件でも送信する'] === true || map['売上0件でも送信する'] === 'はい',
  };
}

function notifyError_(to, e, url) {
  try {
    MailApp.sendEmail({
      to: to,
      subject: '【エラー】日次売上レポートの作成に失敗しました',
      body: 'レポート作成中にエラーが発生しました。\n\nエラー内容: ' + e.message + '\n\nスプレッドシート: ' + url,
    });
  } catch (mailError) {
    console.error('エラー通知メールの送信にも失敗しました: %s', mailError.message);
  }
}

// ---------------------------------------------------------------------------
// 初期設定
// ---------------------------------------------------------------------------
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  if (ss.getSheetByName(SHEET_SALES)) {
    const res = ui.alert('「売上」シートはすでにあります。サンプルデータで上書きしますか?', ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) return;
  }

  createSalesSheet_(ss);
  if (!ss.getSheetByName(SHEET_DAILY)) createDailySheet_(ss);
  createSettingsSheet_(ss);
  ui.alert('初期設定が完了しました。\n「設定」シートで送信先を確認し、メニューの「② 今すぐレポートを送信」で試してください。');
}

function createSalesSheet_(ss) {
  const sheet = ss.getSheetByName(SHEET_SALES) || ss.insertSheet(SHEET_SALES, 0);
  sheet.clear();
  sheet.getRange(1, 1, 1, 6).setValues([['日付', '商品名', 'カテゴリ', '数量', '単価', '金額']])
    .setFontWeight('bold').setBackground('#e8f0fe');

  const products = [
    ['ブレンドコーヒー豆 200g', 'コーヒー豆', 1200],
    ['エチオピア シングルオリジン 200g', 'コーヒー豆', 1800],
    ['ドリップバッグ 10袋', 'コーヒー豆', 980],
    ['ハンドドリップセット', '器具', 4500],
    ['電動コーヒーミル', '器具', 8800],
    ['ペーパーフィルター 100枚', '器具', 450],
    ['焼き菓子アソート', 'スイーツ', 1500],
    ['チーズケーキ', 'スイーツ', 2200],
    ['ギフトボックス', 'ギフト', 3500],
  ];

  const tz = ss.getSpreadsheetTimeZone();
  const rows = [];
  for (let d = 30; d >= 1; d--) {
    const day = new Date(dateKey_(addDays_(new Date(), -d), tz) + 'T00:00:00');
    const n = 5 + Math.floor(Math.random() * 11); // 1日5〜15件
    for (let i = 0; i < n; i++) {
      const p = products[Math.floor(Math.random() * products.length)];
      const qty = 1 + Math.floor(Math.random() * 3);
      rows.push([day, p[0], p[1], qty, p[2], null]);
    }
  }
  const start = 2;
  sheet.getRange(start, 1, rows.length, 6).setValues(rows);
  // 金額列は数式(数量×単価)
  sheet.getRange(start, 6, rows.length, 1).setFormulaR1C1('=R[0]C[-2]*R[0]C[-1]');
  sheet.getRange(start, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(start, 5, rows.length, 2).setNumberFormat('¥#,##0');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 6);
}

function createDailySheet_(ss) {
  const sheet = ss.insertSheet(SHEET_DAILY);
  sheet.getRange(1, 1, 1, 8)
    .setValues([['日付', '売上合計', '前日売上', '前日差', '前日比', '件数', '数量', '記録日時']])
    .setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange('B:D').setNumberFormat('¥#,##0');
  sheet.getRange('E:E').setNumberFormat('+0.0%;-0.0%;0.0%');
  sheet.getRange('H:H').setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.setFrozenRows(1);
  return sheet;
}

function createSettingsSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (sheet) return; // 既存の設定は残す
  sheet = ss.insertSheet(SHEET_SETTINGS);
  const me = Session.getActiveUser().getEmail();
  sheet.getRange(1, 1, 3, 3).setValues([
    ['送信先メールアドレス', me, '複数の場合はカンマ区切り'],
    ['件名の先頭', '【日次売上レポート】', ''],
    ['売上0件でも送信する', 'いいえ', '「はい」にすると0件の日も送信'],
  ]);
  sheet.getRange('A1:A3').setFontWeight('bold').setBackground('#f3f4f6');
  sheet.getRange('C1:C3').setFontColor('#6b7280');
  sheet.autoResizeColumns(1, 3);
}

// ---------------------------------------------------------------------------
// ユーティリティ
// ---------------------------------------------------------------------------
function dateKey_(date, tz) {
  return Utilities.formatDate(date, tz, 'yyyy-MM-dd');
}

function addDays_(date, days) {
  const d = new Date(date.getTime());
  d.setDate(d.getDate() + days);
  return d;
}
