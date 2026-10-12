/**
 * 売上データ集計ダッシュボード
 *
 * 「売上データ」シートの全行を読み込み、月ごとに集計して
 * 「月次サマリー」シートへ書き込み、月次推移を棒グラフで可視化する。
 * 毎朝9時に自動実行するトリガーも登録できる。
 */

// ===== 定数定義 =====
const SOURCE_SHEET_NAME = '売上データ';   // 入力元シート名
const SUMMARY_SHEET_NAME = '月次サマリー'; // 出力先シート名

// 売上データシートの列インデックス（0始まり）
const COL_DATE = 0;   // A列：日付
const COL_PERSON = 1; // B列：担当者名
const COL_PRODUCT = 2; // C列：商品名
const COL_AMOUNT = 3; // D列：金額

/**
 * メインの集計処理。
 * 売上データを月ごとに集計し、サマリーシートへ書き込んで棒グラフを作成する。
 */
function runDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);

  // 入力元シートが存在しない場合は処理を中断する
  if (!sourceSheet) {
    throw new Error('「' + SOURCE_SHEET_NAME + '」シートが見つかりません。');
  }

  // 月ごとの集計結果を保持するマップ
  // 例: { '2026年1月': { total: 150000, count: 3 }, ... }
  const monthlyMap = aggregateMonthly(sourceSheet);

  // 出力先シートを取得（存在しなければ新規作成）
  let summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
  if (!summarySheet) {
    summarySheet = ss.insertSheet(SUMMARY_SHEET_NAME);
  }

  // サマリーシートへ書き込み、棒グラフを作成する
  writeSummary(summarySheet, monthlyMap);
}

/**
 * 売上データシートを読み込み、月ごとに合計売上と件数を集計する。
 *
 * @param {Sheet} sourceSheet 売上データシート
 * @return {Object} 月（例：2026年1月）をキーとした集計結果マップ
 */
function aggregateMonthly(sourceSheet) {
  const monthlyMap = {};

  // ヘッダー行（1行目）を除いたデータ範囲を取得
  const lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) {
    // データが存在しない場合は空のマップを返す
    return monthlyMap;
  }

  // A列〜D列をまとめて一括取得（実行時間短縮のため）
  const values = sourceSheet.getRange(2, 1, lastRow - 1, 4).getValues();

  values.forEach(function (row) {
    const dateValue = row[COL_DATE];
    const amount = row[COL_AMOUNT];

    // 日付または金額が不正な行はスキップする
    if (!(dateValue instanceof Date)) {
      return;
    }
    if (typeof amount !== 'number' || isNaN(amount)) {
      return;
    }

    // 「2026年1月」形式の月キーを作成
    const monthKey = dateValue.getFullYear() + '年' + (dateValue.getMonth() + 1) + '月';

    // 月ごとに合計売上と件数を加算する
    if (!monthlyMap[monthKey]) {
      monthlyMap[monthKey] = { total: 0, count: 0 };
    }
    monthlyMap[monthKey].total += amount;
    monthlyMap[monthKey].count += 1;
  });

  return monthlyMap;
}

/**
 * 集計結果をサマリーシートへ書き込み、月次推移の棒グラフを作成する。
 * シートは毎回クリアしてから書き直す。
 *
 * @param {Sheet} summarySheet 月次サマリーシート
 * @param {Object} monthlyMap 月ごとの集計結果マップ
 */
function writeSummary(summarySheet, monthlyMap) {
  // 既存の内容と既存グラフを毎回クリアする
  summarySheet.clear();
  summarySheet.getCharts().forEach(function (chart) {
    summarySheet.removeChart(chart);
  });

  // 月キーを年月順に昇順ソートする
  const monthKeys = Object.keys(monthlyMap).sort(function (a, b) {
    return parseMonthKey(a) - parseMonthKey(b);
  });

  // 書き込み用の二次元配列を組み立てる（1行目はヘッダー）
  const output = [['月', '合計売上', '件数']];
  monthKeys.forEach(function (monthKey) {
    const data = monthlyMap[monthKey];
    output.push([monthKey, data.total, data.count]);
  });

  // 一括書き込み（実行時間短縮のため）
  summarySheet.getRange(1, 1, output.length, 3).setValues(output);

  // ヘッダー行を太字にする
  summarySheet.getRange(1, 1, 1, 3).setFontWeight('bold');

  // データが無い場合はグラフを作成せずに終了する
  if (monthKeys.length === 0) {
    return;
  }

  // 月（A列）と合計売上（B列）を使って棒グラフを作成する
  const chartRange = summarySheet.getRange(1, 1, output.length, 2);
  const chart = summarySheet.newChart()
    .setChartType(Charts.ChartType.COLUMN) // 棒グラフ
    .addRange(chartRange)
    .setPosition(2, 5, 0, 0) // E列あたりに配置
    .setOption('title', '月次売上推移')
    .setOption('legend', { position: 'none' })
    .build();

  summarySheet.insertChart(chart);
}

/**
 * 「2026年1月」形式の月キーを、比較用の数値（例：202601）へ変換する。
 *
 * @param {string} monthKey 月キー
 * @return {number} 年月を表す数値
 */
function parseMonthKey(monthKey) {
  const matched = monthKey.match(/(\d+)年(\d+)月/);
  if (!matched) {
    return 0;
  }
  const year = parseInt(matched[1], 10);
  const month = parseInt(matched[2], 10);
  return year * 100 + month;
}

/**
 * 毎朝9時に runDashboard を自動実行するトリガーを登録する。
 * 重複登録を防ぐため、既存の同名トリガーを削除してから登録する。
 */
function setDailyTrigger() {
  // 既存の runDashboard 用トリガーを削除する
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'runDashboard') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 毎日午前9時台に実行する時間主導型トリガーを新規登録する
  ScriptApp.newTrigger('runDashboard')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
}
