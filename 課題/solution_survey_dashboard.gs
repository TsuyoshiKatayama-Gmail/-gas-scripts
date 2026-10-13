/**
 * 【課題2 解答例】研修アンケート集計ダッシュボード
 *
 * 「アンケート回答」シートを読み込み、研修名ごとに回答件数・満足度平均・
 * 理解度平均を集計して「研修サマリー」シートへ書き込み、
 * 研修ごとの平均満足度を棒グラフで可視化する。
 * 毎週月曜9時に自動実行するトリガーも登録できる。
 *
 * ※ 本ファイルは課題2の解答例です。定数名は他ファイルとの衝突を避けるため
 *   接頭辞 SV_ を付けています。実運用時は単独の GAS プロジェクトで使用してください。
 */

// ===== 定数定義 =====
const SV_SOURCE_SHEET_NAME = 'アンケート回答'; // 入力元シート名
const SV_SUMMARY_SHEET_NAME = '研修サマリー';   // 出力先シート名

// アンケート回答シートの列インデックス（0始まり）
const SV_COL_DATE = 0;         // A列：回答日時
const SV_COL_TRAINING = 1;     // B列：研修名
const SV_COL_SATISFACTION = 2; // C列：満足度(1〜5)
const SV_COL_UNDERSTANDING = 3; // D列：理解度(1〜5)
// E列：自由記述（集計には使用しない）

// 読み込む列数（A〜E）
const SV_NUM_COLS = 5;

/**
 * メインの集計処理。
 * アンケート回答を研修名ごとに集計し、サマリーシートへ書き込んで棒グラフを作成する。
 */
function runSurveyDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SV_SOURCE_SHEET_NAME);

  // 入力元シートが存在しない場合は処理を中断する
  if (!sourceSheet) {
    throw new Error('「' + SV_SOURCE_SHEET_NAME + '」シートが見つかりません。');
  }

  // 研修名ごとの集計結果を取得する
  const trainingMap = aggregateByTraining(sourceSheet);

  // 出力先シートを取得（存在しなければ新規作成）
  let summarySheet = ss.getSheetByName(SV_SUMMARY_SHEET_NAME);
  if (!summarySheet) {
    summarySheet = ss.insertSheet(SV_SUMMARY_SHEET_NAME);
  }

  // サマリーシートへ書き込み、棒グラフを作成する
  writeSurveySummary(summarySheet, trainingMap);
}

/**
 * アンケート回答シートを読み込み、研修名ごとに件数・満足度合計・理解度合計を集計する。
 *
 * @param {Sheet} sourceSheet アンケート回答シート
 * @return {Object} 研修名をキーとした集計結果マップ
 *   例: { 'ビジネスマナー研修': { count: 10, satisfactionSum: 42, understandingSum: 40 } }
 */
function aggregateByTraining(sourceSheet) {
  const trainingMap = {};

  // ヘッダー行（1行目）を除いたデータ範囲を取得
  const lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) {
    // データが存在しない場合は空のマップを返す
    return trainingMap;
  }

  // A列〜E列をまとめて一括取得（実行時間短縮のため）
  const values = sourceSheet.getRange(2, 1, lastRow - 1, SV_NUM_COLS).getValues();

  values.forEach(function (row) {
    const training = row[SV_COL_TRAINING];
    const satisfaction = row[SV_COL_SATISFACTION];
    const understanding = row[SV_COL_UNDERSTANDING];

    // 研修名が空の行はスキップする
    if (!training) {
      return;
    }
    // 満足度・理解度が数値でない行はスキップする（不正データ対策）
    if (typeof satisfaction !== 'number' || isNaN(satisfaction)) {
      return;
    }
    if (typeof understanding !== 'number' || isNaN(understanding)) {
      return;
    }

    if (!trainingMap[training]) {
      trainingMap[training] = { count: 0, satisfactionSum: 0, understandingSum: 0 };
    }
    trainingMap[training].count += 1;
    trainingMap[training].satisfactionSum += satisfaction;
    trainingMap[training].understandingSum += understanding;
  });

  return trainingMap;
}

/**
 * 集計結果をサマリーシートへ書き込み、研修ごとの平均満足度の棒グラフを作成する。
 * シートは毎回クリアしてから書き直す。
 *
 * @param {Sheet} summarySheet 研修サマリーシート
 * @param {Object} trainingMap 研修名ごとの集計結果マップ
 */
function writeSurveySummary(summarySheet, trainingMap) {
  // 既存の内容と既存グラフを毎回クリアする
  summarySheet.clear();
  summarySheet.getCharts().forEach(function (chart) {
    summarySheet.removeChart(chart);
  });

  // 研修名を名前順に並べる
  const trainingNames = Object.keys(trainingMap).sort();

  // 書き込み用の二次元配列を組み立てる（1行目はヘッダー）
  const output = [['研修名', '回答件数', '満足度平均', '理解度平均']];
  trainingNames.forEach(function (name) {
    const data = trainingMap[name];
    // 平均を小数第1位に丸める
    const avgSatisfaction = roundTo1(data.satisfactionSum / data.count);
    const avgUnderstanding = roundTo1(data.understandingSum / data.count);
    output.push([name, data.count, avgSatisfaction, avgUnderstanding]);
  });

  // 一括書き込み（実行時間短縮のため）
  summarySheet.getRange(1, 1, output.length, 4).setValues(output);

  // ヘッダー行を太字にする
  summarySheet.getRange(1, 1, 1, 4).setFontWeight('bold');

  // データが無い場合はグラフを作成せずに終了する
  if (trainingNames.length === 0) {
    return;
  }

  // 研修名（A列）と満足度平均（C列）を使って棒グラフを作成する
  const chart = summarySheet.newChart()
    .setChartType(Charts.ChartType.COLUMN) // 棒グラフ
    .addRange(summarySheet.getRange(1, 1, output.length, 1)) // A列：研修名
    .addRange(summarySheet.getRange(1, 3, output.length, 1)) // C列：満足度平均
    .setPosition(2, 6, 0, 0) // F列あたりに配置
    .setOption('title', '研修別 平均満足度')
    .setOption('legend', { position: 'none' })
    .build();

  summarySheet.insertChart(chart);
}

/**
 * 数値を小数第1位に丸める。
 *
 * @param {number} value 対象の数値
 * @return {number} 小数第1位に丸めた数値
 */
function roundTo1(value) {
  return Math.round(value * 10) / 10;
}

/**
 * 毎週月曜9時に runSurveyDashboard を自動実行するトリガーを登録する。
 * 重複登録を防ぐため、既存の同名トリガーを削除してから登録する。
 */
function setSurveyWeeklyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'runSurveyDashboard') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 毎週月曜の9時台に実行する時間主導型トリガーを新規登録する
  ScriptApp.newTrigger('runSurveyDashboard')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();
}
