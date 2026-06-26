/**
 * ============================================================
 *  スクリプト名 : 作業ログ自動集計
 *  目的         : 作業ログを日別・クライアント別に自動集計し、
 *                 集計結果を別タブに書き出す
 *  トリガー     : 手動実行（メニュー「作業ログ集計」から）
 *  対象スプシ   : （新規作成したスプシのURLをここに記載）
 *  連携先       : なし（スプシ内で完結）
 *  必要なプロパティ: なし
 *  実行アカウント: 共有アカウント xxx@example.com
 *  管理者/連絡先 : 担当者名（Slack @xxx）
 *  作成/最終更新 : 2026-06-26 / Claude
 *  止め方       : メニューからの手動実行のため、トリガー削除は不要
 * ============================================================
 */

// ── R-2: 設定値を先頭に集約 ──
const CONFIG = {
  LOG_SHEET_GID: 0,                // 作業ログタブのGID
  SUMMARY_SHEET_NAME: '集計結果',   // 集計結果タブの名前（自動作成する）
  LOCK_WAIT_MS: 10000,             // ロック待機時間
  RETRY_MAX: 3,                    // リトライ上限
  TZ: 'Asia/Tokyo',                // タイムゾーン
};

// ── R-1: タブをGIDで取得する関数 ──
/** GIDでシートを取得する（タブ名変更に強い） */
function getSheetByGid_(gid) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .find(s => s.getSheetId() === gid);
  if (!sheet) throw new Error('GID ' + gid + ' のシートが見つかりません');
  return sheet;
}

// ── R-1: ヘッダー名で列を認識するユーティリティ ──
/** シートを一括読み込みし、ヘッダーマップ付きで返す */
function readSheet_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length === 0) throw new Error('シートにデータがありません');
  const header = {};
  values[0].forEach((name, i) => { header[String(name).trim()] = i; });
  return { rows: values.slice(1), header };
}

/** ヘッダー名から列indexを取得する */
function col_(header, name) {
  if (!(name in header)) throw new Error('列「' + name + '」が見つかりません');
  return header[name];
}

// ── R-4: リトライ関数 ──
/** 一時的な失敗を指数バックオフでリトライする */
function withRetry_(fn) {
  for (let i = 0; i < CONFIG.RETRY_MAX; i++) {
    try { return fn(); }
    catch (e) {
      if (i === CONFIG.RETRY_MAX - 1) throw e;
      Utilities.sleep(Math.pow(2, i) * 1000 + Math.floor(Math.random() * 500));
    }
  }
}

// ── R-7: 日付フォーマット（タイムゾーン明示） ──
/** 日付を yyyy/MM/dd 形式の文字列にする */
function formatDate_(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, CONFIG.TZ, 'yyyy/MM/dd');
}

// ── メニュー登録 ──
/** スプシを開いたときにメニューを追加する */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('作業ログ集計')
    .addItem('集計を実行', 'runSummary')
    .addItem('初回セットアップ', 'setupLogSheet')
    .addToUi();
}

// ── 初回セットアップ ──
/** 作業ログタブにヘッダーと書式を設定する */
function setupLogSheet() {
  const sheet = getSheetByGid_(CONFIG.LOG_SHEET_GID);
  const headers = ['日付', 'クライアント名', '作業内容', '作業時間（h）'];

  // ヘッダー書き込み（R-3: 一括書き込み）
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#cfe2f3')
    .setHorizontalAlignment('center');

  // 列幅
  sheet.setColumnWidth(1, 110);  // 日付
  sheet.setColumnWidth(2, 160);  // クライアント名
  sheet.setColumnWidth(3, 300);  // 作業内容
  sheet.setColumnWidth(4, 120);  // 作業時間

  // 日付列のフォーマット
  sheet.getRange('A2:A').setNumberFormat('yyyy/mm/dd');
  // 作業時間列のフォーマット
  sheet.getRange('D2:D').setNumberFormat('0.0');

  SpreadsheetApp.getUi().alert('セットアップ完了！\nA列〜D列にデータを入力してください。');
}

// ── メイン処理 ──
/** 作業ログを集計して結果タブに書き出す */
function runSummary() {
  // ── R-5: 排他制御 ──
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(CONFIG.LOCK_WAIT_MS)) {
    SpreadsheetApp.getUi().alert('別の集計が実行中です。少し待ってから再度お試しください。');
    return;
  }

  try {
    // ── R-1: GIDでシート取得 ──
    const logSheet = getSheetByGid_(CONFIG.LOG_SHEET_GID);
    const { rows, header } = readSheet_(logSheet);

    // ── R-4: バリデーション ──
    if (rows.length === 0) {
      SpreadsheetApp.getUi().alert('作業ログにデータがありません。');
      return;
    }

    // ── R-1: ヘッダー名で列を認識 ──
    const iDate = col_(header, '日付');
    const iClient = col_(header, 'クライアント名');
    const iTask = col_(header, '作業内容');
    const iHours = col_(header, '作業時間（h）');

    // ── 集計ロジック ──
    const byDate = {};      // { '2026/06/26': totalHours }
    const byClient = {};    // { 'クライアントA': { hours, tasks } }

    for (const row of rows) {
      const dateVal = row[iDate];
      const client = String(row[iClient] || '').trim();
      const task = String(row[iTask] || '').trim();
      const hours = Number(row[iHours]) || 0;

      // 空行はスキップ
      if (!client && !task && hours === 0) continue;

      // 日別集計
      const dateKey = formatDate_(dateVal) || '（日付なし）';
      byDate[dateKey] = (byDate[dateKey] || 0) + hours;

      // クライアント別集計
      if (client) {
        if (!byClient[client]) byClient[client] = { hours: 0, tasks: [] };
        byClient[client].hours += hours;
        if (task) byClient[client].tasks.push(task);
      }
    }

    // ── 集計結果タブの準備 ──
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sumSheet = ss.getSheetByName(CONFIG.SUMMARY_SHEET_NAME);
    if (sumSheet) {
      sumSheet.clear();
    } else {
      sumSheet = ss.insertSheet(CONFIG.SUMMARY_SHEET_NAME);
    }

    // ── R-3: 集計結果を一括書き込み ──
    const output = [];

    // 日別集計セクション
    output.push(['【日別集計】', '', '']);
    output.push(['日付', '合計時間（h）', '']);
    const sortedDates = Object.keys(byDate).sort();
    let totalHours = 0;
    for (const d of sortedDates) {
      output.push([d, byDate[d], '']);
      totalHours += byDate[d];
    }
    output.push(['合計', totalHours, '']);
    output.push(['', '', '']);

    // クライアント別集計セクション
    output.push(['【クライアント別集計】', '', '']);
    output.push(['クライアント名', '合計時間（h）', '作業内容']);
    const sortedClients = Object.keys(byClient).sort();
    for (const c of sortedClients) {
      const info = byClient[c];
      const uniqueTasks = [...new Set(info.tasks)];
      output.push([c, info.hours, uniqueTasks.join(' / ')]);
    }

    // 一括書き込み
    if (output.length > 0) {
      withRetry_(() => {
        sumSheet.getRange(1, 1, output.length, 3).setValues(output);
      });
    }

    // ── 書式設定 ──
    // セクションタイトル
    sumSheet.getRange(1, 1).setFontWeight('bold').setFontSize(12);
    sumSheet.getRange(sortedDates.length + 5, 1).setFontWeight('bold').setFontSize(12);

    // ヘッダー行
    sumSheet.getRange(2, 1, 1, 2).setFontWeight('bold').setBackground('#cfe2f3');
    sumSheet.getRange(sortedDates.length + 6, 1, 1, 3).setFontWeight('bold').setBackground('#cfe2f3');

    // 合計行
    sumSheet.getRange(sortedDates.length + 3, 1, 1, 2).setFontWeight('bold').setBackground('#fce5cd');

    // 列幅
    sumSheet.setColumnWidth(1, 160);
    sumSheet.setColumnWidth(2, 120);
    sumSheet.setColumnWidth(3, 400);

    // 時間列のフォーマット
    sumSheet.getRange(3, 2, sortedDates.length + 1, 1).setNumberFormat('0.0');

    SpreadsheetApp.flush();

    // ── R-7: 最終更新日時をタイムゾーン明示で記録 ──
    const now = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy/MM/dd HH:mm:ss');
    const footerRow = output.length + 2;
    sumSheet.getRange(footerRow, 1).setValue('最終集計: ' + now)
      .setFontSize(10).setFontColor('#999999');

    SpreadsheetApp.getUi().alert('集計完了！\n「' + CONFIG.SUMMARY_SHEET_NAME + '」タブを確認してください。');

  } catch (e) {
    console.error('runSummary failed: ' + e.message, e.stack);
    SpreadsheetApp.getUi().alert('エラーが発生しました:\n' + e.message);
    throw e;
  } finally {
    lock.releaseLock();
  }
}
