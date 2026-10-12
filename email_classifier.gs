/**
 * メール自動分類・通知システム
 *
 * Gmail の「要処理」ラベルが付いた未読メールを取得し、
 * Claude API で「クレーム」「質問」「注文」「その他」に分類する。
 * 分類結果をスプレッドシートの「メールログ」シートへ記録し、
 * Slack（Incoming Webhook）で担当者へ通知する。
 * 処理が完了したメールには「処理済み」ラベルを付与し「要処理」ラベルを外す。
 *
 * API キーなどの機密情報はスクリプトプロパティから取得する
 *   - CLAUDE_API_KEY    : Claude API キー
 *   - SLACK_WEBHOOK_URL : Slack Incoming Webhook の URL
 */

// ===== 定数定義 =====
const TARGET_LABEL_NAME = '要処理';   // 処理対象を示すラベル名
const DONE_LABEL_NAME = '処理済み';   // 処理完了を示すラベル名

const LOG_SHEET_NAME = 'メールログ';  // 分類結果の記録先シート名
const ERROR_SHEET_NAME = 'エラーログ'; // エラー内容の記録先シート名

// メールログシートの列見出し（受信日時・送信者・件名・分類・要約の順）
const LOG_HEADERS = ['受信日時', '送信者', '件名', '分類', '要約'];
// エラーログシートの列見出し
const ERROR_HEADERS = ['発生日時', '件名', 'エラー内容'];

// スクリプトプロパティのキー名
const PROP_CLAUDE_API_KEY = 'CLAUDE_API_KEY';
const PROP_SLACK_WEBHOOK_URL = 'SLACK_WEBHOOK_URL';

// Claude API 設定
const CLAUDE_API_URL = 'https://api.anthropic.com/v1/messages';
const CLAUDE_API_VERSION = '2023-06-01';
const CLAUDE_MODEL = 'claude-haiku-4-5'; // コスト重視のため最新の Haiku を使用
const CLAUDE_MAX_TOKENS = 1024;

// 1回の実行で処理する最大スレッド数（実行時間の上限対策）
const MAX_THREADS_PER_RUN = 10;

// 分類カテゴリの一覧
const CATEGORIES = ['クレーム', '質問', '注文', 'その他'];

/**
 * メインの処理。
 * 「要処理」ラベルの未読メールを分類・記録・通知し、ラベルを付け替える。
 */
function classifyEmails() {
  // 「要処理」ラベルを取得（存在しなければ処理を中断）
  const targetLabel = GmailApp.getUserLabelByName(TARGET_LABEL_NAME);
  if (!targetLabel) {
    throw new Error('「' + TARGET_LABEL_NAME + '」ラベルが見つかりません。');
  }

  // 「処理済み」ラベルを取得（存在しなければ新規作成）
  let doneLabel = GmailApp.getUserLabelByName(DONE_LABEL_NAME);
  if (!doneLabel) {
    doneLabel = GmailApp.createLabel(DONE_LABEL_NAME);
  }

  // 「要処理」ラベルが付いた未読スレッドを取得する
  // search で未読かつ対象ラベルのスレッドに絞り込む
  const query = 'label:"' + TARGET_LABEL_NAME + '" is:unread';
  const threads = GmailApp.search(query, 0, MAX_THREADS_PER_RUN);

  threads.forEach(function (thread) {
    // スレッド内の未読メッセージを対象に処理する
    thread.getMessages().forEach(function (message) {
      if (!message.isUnread()) {
        return;
      }

      const subject = message.getSubject();
      try {
        // メール本文を Claude API で分類・要約する
        const result = classifyWithClaude(message.getPlainBody());

        // スプレッドシートの「メールログ」シートへ記録する
        appendMailLog(
          message.getDate(),
          message.getFrom(),
          subject,
          result.category,
          result.summary
        );

        // Slack へ担当者に通知する（件名・分類・要約を含める）
        notifySlack(subject, result.category, result.summary);

        // 処理済みとして扱うため既読にする
        message.markRead();
      } catch (e) {
        // エラー内容を「エラーログ」シートへ記録する
        appendErrorLog(subject, e);
      }
    });

    // スレッド単位でラベルを付け替える
    // 「処理済み」を付与し「要処理」を外す
    thread.addLabel(doneLabel);
    thread.removeLabel(targetLabel);
  });
}

/**
 * Claude API にメール本文を送り、分類と要約を取得する。
 *
 * @param {string} body メール本文（プレーンテキスト）
 * @return {{category: string, summary: string}} 分類結果と要約
 */
function classifyWithClaude(body) {
  // スクリプトプロパティから API キーを取得する
  const apiKey = PropertiesService.getScriptProperties().getProperty(PROP_CLAUDE_API_KEY);
  if (!apiKey) {
    throw new Error('スクリプトプロパティ「' + PROP_CLAUDE_API_KEY + '」が設定されていません。');
  }

  // Claude へ渡す指示（JSON 形式での回答を求める）
  const systemPrompt =
    'あなたはカスタマーサポートのメール分類担当です。' +
    '受け取ったメール本文を読み、次の4つのカテゴリのいずれか1つに分類してください：' +
    CATEGORIES.join('、') + '。' +
    'また、本文を日本語で1〜2文に要約してください。' +
    '回答は必ず次の JSON 形式のみで出力してください（前後に説明文を付けないこと）：' +
    '{"category": "<カテゴリ名>", "summary": "<要約>"}';

  const payload = {
    model: CLAUDE_MODEL,
    max_tokens: CLAUDE_MAX_TOKENS,
    system: systemPrompt,
    messages: [
      { role: 'user', content: '次のメール本文を分類・要約してください。\n\n' + body }
    ]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': CLAUDE_API_VERSION
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true // エラー時もレスポンス本文を取得して原因を確認する
  };

  // API を呼び出す
  const response = UrlFetchApp.fetch(CLAUDE_API_URL, options);
  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();

  // 2xx 以外はエラーとして扱う
  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('Claude API エラー (HTTP ' + statusCode + '): ' + responseText);
  }

  // レスポンス JSON からテキスト本文を取り出す
  const json = JSON.parse(responseText);
  const text = extractTextFromResponse(json);

  // テキスト中の JSON 部分を解析して分類結果を得る
  return parseClassificationResult(text);
}

/**
 * Claude API のレスポンスから最初のテキストブロックを取り出す。
 *
 * @param {Object} json パース済みのレスポンス JSON
 * @return {string} テキスト本文
 */
function extractTextFromResponse(json) {
  if (!json || !Array.isArray(json.content)) {
    throw new Error('Claude API のレスポンス形式が不正です。');
  }

  // content 配列から type が 'text' のブロックを探す
  for (let i = 0; i < json.content.length; i++) {
    const block = json.content[i];
    if (block && block.type === 'text' && typeof block.text === 'string') {
      return block.text;
    }
  }

  throw new Error('Claude API のレスポンスにテキストが含まれていません。');
}

/**
 * Claude が返したテキストから JSON を抽出し、分類結果を組み立てる。
 * JSON の前後に余計な文字があっても最初の { 〜 最後の } を取り出して解析する。
 *
 * @param {string} text Claude の回答テキスト
 * @return {{category: string, summary: string}} 分類結果と要約
 */
function parseClassificationResult(text) {
  const start = text.indexOf('{');
  const end = text.lastIndexOf('}');
  if (start === -1 || end === -1 || end < start) {
    throw new Error('分類結果の JSON を解析できませんでした: ' + text);
  }

  const jsonText = text.substring(start, end + 1);
  const parsed = JSON.parse(jsonText);

  // カテゴリが想定外の場合は「その他」に寄せる
  let category = parsed.category;
  if (CATEGORIES.indexOf(category) === -1) {
    category = 'その他';
  }

  const summary = parsed.summary || '';
  return { category: category, summary: summary };
}

/**
 * 分類結果を「メールログ」シートへ1行追記する。
 * シートが無ければ新規作成し、ヘッダー行を用意する。
 *
 * @param {Date} receivedAt 受信日時
 * @param {string} from 送信者
 * @param {string} subject 件名
 * @param {string} category 分類
 * @param {string} summary 要約
 */
function appendMailLog(receivedAt, from, subject, category, summary) {
  const sheet = getOrCreateSheet(LOG_SHEET_NAME, LOG_HEADERS);
  // 受信日時・送信者・件名・分類・要約の順で1行追記する
  sheet.appendRow([receivedAt, from, subject, category, summary]);
}

/**
 * エラー内容を「エラーログ」シートへ1行追記する。
 * シートが無ければ新規作成し、ヘッダー行を用意する。
 *
 * @param {string} subject 対象メールの件名
 * @param {Error} error 発生したエラー
 */
function appendErrorLog(subject, error) {
  const sheet = getOrCreateSheet(ERROR_SHEET_NAME, ERROR_HEADERS);
  const message = error && error.message ? error.message : String(error);
  // 発生日時・件名・エラー内容の順で1行追記する
  sheet.appendRow([new Date(), subject, message]);
}

/**
 * 指定名のシートを取得する。存在しなければ新規作成してヘッダー行を書き込む。
 *
 * @param {string} sheetName シート名
 * @param {Array<string>} headers ヘッダー行の見出し配列
 * @return {Sheet} 対象シート
 */
function getOrCreateSheet(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    // ヘッダー行を一括で書き込む
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

/**
 * Slack の Incoming Webhook へ通知を送る（件名・分類・要約を含める）。
 *
 * @param {string} subject 件名
 * @param {string} category 分類
 * @param {string} summary 要約
 */
function notifySlack(subject, category, summary) {
  // スクリプトプロパティから Webhook URL を取得する
  const webhookUrl = PropertiesService.getScriptProperties().getProperty(PROP_SLACK_WEBHOOK_URL);
  if (!webhookUrl) {
    throw new Error('スクリプトプロパティ「' + PROP_SLACK_WEBHOOK_URL + '」が設定されていません。');
  }

  // Slack へ送るメッセージを組み立てる
  const messageText =
    '*新着メールを分類しました*\n' +
    '• 件名: ' + subject + '\n' +
    '• 分類: ' + category + '\n' +
    '• 要約: ' + summary;

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
 * 5分おきに classifyEmails を自動実行するトリガーを登録する。
 * 重複登録を防ぐため、既存の同名トリガーを削除してから登録する。
 */
function setClassifyTrigger() {
  // 既存の classifyEmails 用トリガーを削除する
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'classifyEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 5分おきに実行する時間主導型トリガーを新規登録する
  ScriptApp.newTrigger('classifyEmails')
    .timeBased()
    .everyMinutes(5)
    .create();
}
