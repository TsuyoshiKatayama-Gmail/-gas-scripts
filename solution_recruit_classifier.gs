/**
 * 【課題1 解答例】採用応募メールの自動仕分け・通知システム
 *
 * Gmail の「応募」ラベルが付いた未読メールを取得し、Claude API で
 * 応募区分（新卒/中途/アルバイト/その他）・希望職種・要約を判定する。
 * 結果を「応募ログ」シートへ記録し、Slack へ通知する。
 * 処理済みメールには「選考中」ラベルを付け「応募」ラベルを外す。
 *
 * スクリプトプロパティ:
 *   - CLAUDE_API_KEY    : Claude API キー
 *   - SLACK_WEBHOOK_URL : Slack Incoming Webhook の URL
 *
 * ※ 本ファイルは課題1の解答例です。定数名は他ファイルとの衝突を避けるため
 *   接頭辞 RC_ を付けています。実運用時は単独の GAS プロジェクトで使用してください。
 */

// ===== 定数定義 =====
const RC_TARGET_LABEL_NAME = '応募';     // 処理対象ラベル
const RC_DONE_LABEL_NAME = '選考中';     // 処理完了ラベル

const RC_LOG_SHEET_NAME = '応募ログ';
const RC_ERROR_SHEET_NAME = 'エラーログ';

// 応募ログシートの列見出し（受信日時・応募者メール・件名・応募区分・希望職種・要約）
const RC_LOG_HEADERS = ['受信日時', '応募者メール', '件名', '応募区分', '希望職種', '要約'];
const RC_ERROR_HEADERS = ['発生日時', '件名', 'エラー内容'];

const RC_PROP_CLAUDE_API_KEY = 'CLAUDE_API_KEY';
const RC_PROP_SLACK_WEBHOOK_URL = 'SLACK_WEBHOOK_URL';

// Claude API 設定
const RC_CLAUDE_API_URL = 'https://api.anthropic.com/v1/messages';
const RC_CLAUDE_API_VERSION = '2023-06-01';
const RC_CLAUDE_MODEL = 'claude-haiku-4-5'; // コスト重視で最新の Haiku を使用
const RC_CLAUDE_MAX_TOKENS = 1024;

// 1回の実行で処理する最大スレッド数（実行時間の上限対策）
const RC_MAX_THREADS_PER_RUN = 10;

// 応募区分の一覧
const RC_CATEGORIES = ['新卒', '中途', 'アルバイト', 'その他'];

/**
 * メインの処理。
 * 「応募」ラベルの未読メールを分類・記録・通知し、ラベルを付け替える。
 */
function classifyRecruitEmails() {
  // 「応募」ラベルを取得（存在しなければ処理を中断）
  const targetLabel = GmailApp.getUserLabelByName(RC_TARGET_LABEL_NAME);
  if (!targetLabel) {
    throw new Error('「' + RC_TARGET_LABEL_NAME + '」ラベルが見つかりません。');
  }

  // 「選考中」ラベルを取得（存在しなければ新規作成）
  let doneLabel = GmailApp.getUserLabelByName(RC_DONE_LABEL_NAME);
  if (!doneLabel) {
    doneLabel = GmailApp.createLabel(RC_DONE_LABEL_NAME);
  }

  // 「応募」ラベルが付いた未読スレッドを取得する
  const query = 'label:"' + RC_TARGET_LABEL_NAME + '" is:unread';
  const threads = GmailApp.search(query, 0, RC_MAX_THREADS_PER_RUN);

  threads.forEach(function (thread) {
    // スレッド内の未読メッセージを対象に処理する
    thread.getMessages().forEach(function (message) {
      if (!message.isUnread()) {
        return;
      }

      const subject = message.getSubject();
      try {
        // メール本文を Claude API で分類・抽出・要約する
        const result = classifyRecruitWithClaude(message.getPlainBody());

        // 「応募ログ」シートへ記録する
        appendRecruitLog(
          message.getDate(),
          message.getFrom(),
          subject,
          result.category,
          result.position,
          result.summary
        );

        // Slack へ通知する
        notifyRecruitSlack(subject, result.category, result.position, result.summary);

        // 処理済みとして既読にする
        message.markRead();
      } catch (e) {
        // 1通の失敗で全体を止めず、エラーログへ記録して継続する
        appendRecruitErrorLog(subject, e);
      }
    });

    // スレッド単位でラベルを付け替える
    thread.addLabel(doneLabel);
    thread.removeLabel(targetLabel);
  });
}

/**
 * Claude API にメール本文を送り、応募区分・希望職種・要約を取得する。
 *
 * @param {string} body メール本文（プレーンテキスト）
 * @return {{category: string, position: string, summary: string}} 判定結果
 */
function classifyRecruitWithClaude(body) {
  const apiKey = PropertiesService.getScriptProperties().getProperty(RC_PROP_CLAUDE_API_KEY);
  if (!apiKey) {
    throw new Error('スクリプトプロパティ「' + RC_PROP_CLAUDE_API_KEY + '」が設定されていません。');
  }

  // Claude へ渡す指示（JSON 形式のみでの回答を求める）
  const systemPrompt =
    'あなたは採用担当のメール仕分け担当です。' +
    '受け取った応募メール本文を読み、次の3つの情報を判定してください。' +
    '(1) 応募区分：次のいずれか1つ（' + RC_CATEGORIES.join('、') + '）。' +
    '(2) 希望職種：本文から読み取れる希望職種。読み取れない場合は「不明」。' +
    '(3) 要約：本文を日本語で1〜2文に要約。' +
    '回答は必ず次の JSON 形式のみで出力してください（前後に説明文を付けないこと）：' +
    '{"category": "<応募区分>", "position": "<希望職種>", "summary": "<要約>"}';

  const payload = {
    model: RC_CLAUDE_MODEL,
    max_tokens: RC_CLAUDE_MAX_TOKENS,
    system: systemPrompt,
    messages: [
      { role: 'user', content: '次の応募メール本文を判定してください。\n\n' + body }
    ]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': RC_CLAUDE_API_VERSION
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true // エラー時も本文を取得して原因を確認する
  };

  const response = UrlFetchApp.fetch(RC_CLAUDE_API_URL, options);
  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('Claude API エラー (HTTP ' + statusCode + '): ' + responseText);
  }

  const json = JSON.parse(responseText);
  const text = extractRecruitText(json);
  return parseRecruitResult(text);
}

/**
 * Claude API のレスポンスから最初のテキストブロックを取り出す。
 *
 * @param {Object} json パース済みレスポンス
 * @return {string} テキスト本文
 */
function extractRecruitText(json) {
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
 * Claude の回答テキストから JSON を抽出し、判定結果を組み立てる。
 * 想定外の応募区分は「その他」に寄せる（フォールバック）。
 *
 * @param {string} text Claude の回答テキスト
 * @return {{category: string, position: string, summary: string}} 判定結果
 */
function parseRecruitResult(text) {
  const start = text.indexOf('{');
  const end = text.lastIndexOf('}');
  if (start === -1 || end === -1 || end < start) {
    throw new Error('判定結果の JSON を解析できませんでした: ' + text);
  }

  const parsed = JSON.parse(text.substring(start, end + 1));

  // 応募区分が想定外の場合は「その他」に寄せる
  let category = parsed.category;
  if (RC_CATEGORIES.indexOf(category) === -1) {
    category = 'その他';
  }

  const position = parsed.position || '不明';
  const summary = parsed.summary || '';
  return { category: category, position: position, summary: summary };
}

/**
 * 判定結果を「応募ログ」シートへ1行追記する。
 *
 * @param {Date} receivedAt 受信日時
 * @param {string} from 応募者メール
 * @param {string} subject 件名
 * @param {string} category 応募区分
 * @param {string} position 希望職種
 * @param {string} summary 要約
 */
function appendRecruitLog(receivedAt, from, subject, category, position, summary) {
  const sheet = getOrCreateRecruitSheet(RC_LOG_SHEET_NAME, RC_LOG_HEADERS);
  sheet.appendRow([receivedAt, from, subject, category, position, summary]);
}

/**
 * エラー内容を「エラーログ」シートへ1行追記する。
 *
 * @param {string} subject 対象メールの件名
 * @param {Error} error 発生したエラー
 */
function appendRecruitErrorLog(subject, error) {
  const sheet = getOrCreateRecruitSheet(RC_ERROR_SHEET_NAME, RC_ERROR_HEADERS);
  const message = error && error.message ? error.message : String(error);
  sheet.appendRow([new Date(), subject, message]);
}

/**
 * 指定名のシートを取得する。無ければ新規作成してヘッダー行を書き込む。
 *
 * @param {string} sheetName シート名
 * @param {Array<string>} headers ヘッダー見出し配列
 * @return {Sheet} 対象シート
 */
function getOrCreateRecruitSheet(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

/**
 * Slack の Incoming Webhook へ通知を送る。
 *
 * @param {string} subject 件名
 * @param {string} category 応募区分
 * @param {string} position 希望職種
 * @param {string} summary 要約
 */
function notifyRecruitSlack(subject, category, position, summary) {
  const webhookUrl = PropertiesService.getScriptProperties().getProperty(RC_PROP_SLACK_WEBHOOK_URL);
  if (!webhookUrl) {
    throw new Error('スクリプトプロパティ「' + RC_PROP_SLACK_WEBHOOK_URL + '」が設定されていません。');
  }

  const messageText =
    '*新着の応募メールを仕分けしました*\n' +
    '• 件名: ' + subject + '\n' +
    '• 応募区分: ' + category + '\n' +
    '• 希望職種: ' + position + '\n' +
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
 * 5分おきに classifyRecruitEmails を自動実行するトリガーを登録する。
 * 重複登録を防ぐため、既存の同名トリガーを削除してから登録する。
 */
function setRecruitClassifyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'classifyRecruitEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('classifyRecruitEmails')
    .timeBased()
    .everyMinutes(5)
    .create();
}
