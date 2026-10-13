/**
 * 【課題3 解答例】問い合わせ対応 週次レポート自動配信
 *
 * 「メールログ」シート（email_classifier.gs が出力する形式）から
 * 直近7日間のログを集計し、集計値を Claude API に渡して週次サマリー文章を
 * 生成させ、内訳とともに Slack へ自動配信する。
 * 毎週金曜17時に自動実行するトリガーも登録できる。
 *
 * スクリプトプロパティ:
 *   - CLAUDE_API_KEY    : Claude API キー
 *   - SLACK_WEBHOOK_URL : Slack Incoming Webhook の URL
 *
 * ※ 本ファイルは課題3の解答例です。定数名は他ファイルとの衝突を避けるため
 *   接頭辞 WR_ を付けています。実運用時は単独の GAS プロジェクトで使用してください。
 */

// ===== 定数定義 =====
const WR_LOG_SHEET_NAME = 'メールログ'; // 入力元シート名

// メールログシートの列インデックス（0始まり）：受信日時・送信者・件名・分類・要約
const WR_COL_RECEIVED_AT = 0;
const WR_COL_CATEGORY = 3;
const WR_NUM_COLS = 5;

// 集計対象期間（日数）
const WR_PERIOD_DAYS = 7;

// 集計対象とする分類の一覧（表示順）
const WR_CATEGORIES = ['クレーム', '質問', '注文', 'その他'];

const WR_PROP_CLAUDE_API_KEY = 'CLAUDE_API_KEY';
const WR_PROP_SLACK_WEBHOOK_URL = 'SLACK_WEBHOOK_URL';

// Claude API 設定
const WR_CLAUDE_API_URL = 'https://api.anthropic.com/v1/messages';
const WR_CLAUDE_API_VERSION = '2023-06-01';
const WR_CLAUDE_MODEL = 'claude-haiku-4-5';
const WR_CLAUDE_MAX_TOKENS = 1024;

/**
 * メインの処理。
 * 直近7日間のメールログを集計し、Claude で文章化して Slack へ配信する。
 */
function sendWeeklyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(WR_LOG_SHEET_NAME);
  if (!sheet) {
    throw new Error('「' + WR_LOG_SHEET_NAME + '」シートが見つかりません。');
  }

  // 直近7日間のログを集計する
  const stats = aggregateRecentLogs(sheet);

  // 対象ログが0件の場合は、その旨を通知して終了する
  if (stats.total === 0) {
    notifyWeeklySlack('今週（直近' + WR_PERIOD_DAYS + '日間）の問い合わせはありませんでした。', stats);
    return;
  }

  // 集計値（数値のみ）を Claude に渡して週次サマリー文章を生成させる
  const reportText = generateReportWithClaude(stats);

  // Slack へ配信する
  notifyWeeklySlack(reportText, stats);
}

/**
 * メールログシートから直近 WR_PERIOD_DAYS 日間のログを集計する。
 * 個別の本文は読み込まず、件数のみを数える。
 *
 * @param {Sheet} sheet メールログシート
 * @return {{total: number, byCategory: Object}} 集計結果
 */
function aggregateRecentLogs(sheet) {
  const byCategory = {};
  WR_CATEGORIES.forEach(function (category) {
    byCategory[category] = 0;
  });

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { total: 0, byCategory: byCategory };
  }

  // データ範囲を一括取得する（実行時間短縮のため）
  const values = sheet.getRange(2, 1, lastRow - 1, WR_NUM_COLS).getValues();

  // 集計の下限となる日時（現在時刻の7日前）を求める
  const threshold = new Date();
  threshold.setDate(threshold.getDate() - WR_PERIOD_DAYS);

  let total = 0;
  values.forEach(function (row) {
    const receivedAt = row[WR_COL_RECEIVED_AT];
    // 受信日時が日付として読めない行はスキップする
    if (!(receivedAt instanceof Date)) {
      return;
    }
    // しきい値より前（7日より古い）のログはスキップする
    if (receivedAt < threshold) {
      return;
    }

    // 想定外の分類は「その他」に寄せる
    let category = row[WR_COL_CATEGORY];
    if (WR_CATEGORIES.indexOf(category) === -1) {
      category = 'その他';
    }
    byCategory[category] += 1;
    total += 1;
  });

  return { total: total, byCategory: byCategory };
}

/**
 * 集計値を Claude API に渡し、日本語の週次サマリー文章を生成させる。
 * ※ 個別メールの本文・要約は渡さず、集計済みの数値のみを渡す
 *   （トークン節約とプライバシー配慮のため）。
 *
 * @param {{total: number, byCategory: Object}} stats 集計結果
 * @return {string} 生成された週次サマリー文章
 */
function generateReportWithClaude(stats) {
  const apiKey = PropertiesService.getScriptProperties().getProperty(WR_PROP_CLAUDE_API_KEY);
  if (!apiKey) {
    throw new Error('スクリプトプロパティ「' + WR_PROP_CLAUDE_API_KEY + '」が設定されていません。');
  }

  // Claude へ渡す集計値の文字列を組み立てる
  const breakdownLines = WR_CATEGORIES.map(function (category) {
    return '・' + category + ': ' + stats.byCategory[category] + '件';
  }).join('\n');

  const systemPrompt =
    'あなたはカスタマーサポートの責任者です。' +
    '渡された問い合わせ件数の集計値をもとに、社内向けの週次サマリーを' +
    '日本語で3〜4文にまとめてください。件数の多い分類や特徴に触れ、' +
    '簡潔で読みやすい文章にしてください。箇条書きではなく文章で出力してください。';

  const userContent =
    '直近' + WR_PERIOD_DAYS + '日間の問い合わせ集計です。\n' +
    '総件数: ' + stats.total + '件\n' +
    '分類別内訳:\n' + breakdownLines;

  const payload = {
    model: WR_CLAUDE_MODEL,
    max_tokens: WR_CLAUDE_MAX_TOKENS,
    system: systemPrompt,
    messages: [
      { role: 'user', content: userContent }
    ]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': WR_CLAUDE_API_VERSION
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(WR_CLAUDE_API_URL, options);
  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('Claude API エラー (HTTP ' + statusCode + '): ' + responseText);
  }

  const json = JSON.parse(responseText);
  return extractWeeklyText(json);
}

/**
 * Claude API のレスポンスから最初のテキストブロックを取り出す。
 *
 * @param {Object} json パース済みレスポンス
 * @return {string} テキスト本文
 */
function extractWeeklyText(json) {
  if (!json || !Array.isArray(json.content)) {
    throw new Error('Claude API のレスポンス形式が不正です。');
  }
  for (let i = 0; i < json.content.length; i++) {
    const block = json.content[i];
    if (block && block.type === 'text' && typeof block.text === 'string') {
      return block.text;
    }
  }
  throw new Error('Claude API のレスポンスにテキストが含まれていません。');
}

/**
 * 週次レポートを Slack へ投稿する（生成文章＋分類別内訳）。
 *
 * @param {string} reportText 本文（Claude の生成文章、または0件の案内）
 * @param {{total: number, byCategory: Object}} stats 集計結果
 */
function notifyWeeklySlack(reportText, stats) {
  const webhookUrl = PropertiesService.getScriptProperties().getProperty(WR_PROP_SLACK_WEBHOOK_URL);
  if (!webhookUrl) {
    throw new Error('スクリプトプロパティ「' + WR_PROP_SLACK_WEBHOOK_URL + '」が設定されていません。');
  }

  // 分類別内訳の行を組み立てる
  const breakdownLines = WR_CATEGORIES.map(function (category) {
    return '• ' + category + ': ' + stats.byCategory[category] + '件';
  }).join('\n');

  const messageText =
    '*【週次レポート】問い合わせ対応サマリー（直近' + WR_PERIOD_DAYS + '日間）*\n' +
    reportText + '\n\n' +
    '*内訳（総件数: ' + stats.total + '件）*\n' +
    breakdownLines;

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ text: messageText }),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(webhookUrl, options);
  const statusCode = response.getResponseCode();
  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('Slack 通知エラー (HTTP ' + statusCode + '): ' + response.getContentText());
  }
}

/**
 * 毎週金曜17時に sendWeeklyReport を自動実行するトリガーを登録する。
 * 重複登録を防ぐため、既存の同名トリガーを削除してから登録する。
 */
function setWeeklyReportTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'sendWeeklyReport') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 毎週金曜の17時台に実行する時間主導型トリガーを新規登録する
  ScriptApp.newTrigger('sendWeeklyReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(17)
    .create();
}
