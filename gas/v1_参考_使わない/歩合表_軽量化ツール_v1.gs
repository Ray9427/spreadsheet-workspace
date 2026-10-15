/**
 * 歩合表 軽量化ツール (Google Apps Script)
 *
 *   ① 空白行列の削除 …… データの最終行・最終列より外側にある空白の行と列を削除する
 *   ② 列全体参照の置換 …… '編明細'!$R:$R のような参照を '編明細'!$R$1:$R$8617 のような範囲に置き換える
 *   ③ 外部連携用シート …… 「001」の A1:AS35 を「外部連携用」シートへ値だけで転記する(自動更新つき)
 *
 * 使い方: 保存してスプレッドシートを開き直すと、メニューに「🛠 軽量化ツール」が出ます。
 *         削除・置換は、先に「確認」メニューで内容を見てから実行してください。
 */

const CONFIG = {
  // ① 空白行列を削除するシート
  TRIM_SHEETS: ['001', '編明細', '編打刻'],
  // データの下に残す空白行の数。書いていないシートは 0(完全に削除)
  //   編明細・編打刻はデータが増えたときや、数式の結果が下に広がるときのために余白を残す
  KEEP_EMPTY_ROWS: { '編明細': 500, '編打刻': 500 },
  // 「人別シート」としてまとめて処理するシート名(3桁の数字: 000〜350 など)
  PERSON_SHEET_PATTERN: /^\d{3}$/,

  // ② 列全体参照を置き換えるとき、参照先シートのデータ最終行に足す余白の行数
  REF_HEADROOM_ROWS: 500,

  // ③ 外部連携用シート
  EXPORT: {
    SOURCE_SHEET: '001',
    SOURCE_RANGE: 'A1:AS35',
    TARGET_SHEET: '外部連携用',
    SYNC_EVERY_MINUTES: 10, // 1 / 5 / 10 / 15 / 30 のどれか
    // (任意) IMPORTRANGE をやめて、読み込む側のファイルへ直接書き込むときだけ設定する
    //   例: PUSH_SPREADSHEET_ID: '1AbC...xyz'(相手ファイルの URL の /d/ と /edit の間)
    PUSH_SPREADSHEET_ID: '',
    PUSH_SHEET_NAME: '歩合表_001',
  },

  // Apps Script は1回6分で強制終了するので、その手前で止める(続きは同じメニューをもう一度)
  TIME_LIMIT_MS: 4.5 * 60 * 1000,
};

// ============================================================
// メニュー
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛠 軽量化ツール')
    .addItem('① 空白行列を確認する(削除しない)', 'previewTrim')
    .addItem('① 空白行列を削除する', 'runTrim')
    .addSeparator()
    .addItem('①+ 人別シート(000〜350)の空白行列を確認する', 'previewTrimPersonSheets')
    .addItem('①+ 人別シート(000〜350)の空白行列を削除する', 'runTrimPersonSheets')
    .addSeparator()
    .addItem('② 列全体参照を確認する(変更しない)', 'previewColumnRefs')
    .addItem('② 列全体参照を範囲指定に置き換える', 'runColumnRefs')
    .addSeparator()
    .addItem('③ 外部連携用シートを今すぐ更新する', 'syncExternalSheet')
    .addItem('③ 自動更新をオンにする', 'enableAutoSync')
    .addItem('③ 自動更新をオフにする', 'disableAutoSync')
    .addToUi();
}

// ============================================================
// ① 空白行列の削除
// ============================================================

function previewTrim() {
  trimSheets_(CONFIG.TRIM_SHEETS, false);
}

function runTrim() {
  const names = CONFIG.TRIM_SHEETS.join('」「');
  if (!confirm_('「' + names + '」の空白の行と列を削除します。\nバックアップは取りましたか?')) return;
  trimSheets_(CONFIG.TRIM_SHEETS, true);
}

function previewTrimPersonSheets() {
  trimSheets_(personSheetNames_(), false);
}

function runTrimPersonSheets() {
  const names = personSheetNames_();
  if (!confirm_('人別シート ' + names.length + ' 枚の空白の行と列を削除します。\nバックアップは取りましたか?')) return;
  trimSheets_(names, true);
}

function personSheetNames_() {
  return SpreadsheetApp.getActive()
    .getSheets()
    .map((s) => s.getName())
    .filter((n) => CONFIG.PERSON_SHEET_PATTERN.test(n));
}

function trimSheets_(names, apply) {
  const started = Date.now();
  const ss = SpreadsheetApp.getActive();
  toast_(apply ? '空白の行と列を削除しています…' : '空白の行と列を数えています…');

  const lines = [];
  let handled = 0;
  let delRows = 0;
  let delCols = 0;
  let cellsBefore = 0;
  let cellsAfter = 0;

  for (const name of names) {
    if (Date.now() - started > CONFIG.TIME_LIMIT_MS) break;
    handled++;
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      lines.push('「' + name + '」: シートが見つかりません');
      continue;
    }
    const p = trimPlan_(sheet, CONFIG.KEEP_EMPTY_ROWS[name] || 0);
    cellsBefore += p.maxRows * p.maxCols;
    cellsAfter += p.keepRows * p.keepCols;
    if (p.delRows === 0 && p.delCols === 0) {
      lines.push('「' + name + '」: 削除する行・列はありません');
      continue;
    }
    if (apply) {
      if (p.delRows > 0) sheet.deleteRows(p.keepRows + 1, p.delRows);
      if (p.delCols > 0) sheet.deleteColumns(p.keepCols + 1, p.delCols);
    }
    delRows += p.delRows;
    delCols += p.delCols;
    lines.push(
      '「' + name + '」: ' + p.maxRows + '行×' + p.maxCols + '列 → ' + p.keepRows + '行×' + p.keepCols + '列'
    );
  }

  const head = [
    apply ? '削除しました。' : '【確認だけです。まだ何も削除していません】',
    '対象 ' + handled + ' / ' + names.length + ' シート: 行 ' + delRows + ' 本・列 ' + delCols + ' 本' +
      (apply ? 'を削除' : 'が削除の対象'),
    'セルの数: ' + cellsBefore.toLocaleString() + ' → ' + cellsAfter.toLocaleString(),
  ];
  if (handled < names.length) {
    head.push(
      '時間切れのため、' + (names.length - handled) + ' シートが残っています。同じメニューをもう一度実行すると続きから処理します。'
    );
  }
  notify_('① 空白行列の削除', head.concat([''], lines));
}

/** 何行目・何列目まで残すかを決める(データ・結合セル・外部連携の範囲・固定行列を考慮) */
function trimPlan_(sheet, keepEmptyRows) {
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  let lastRow = Math.max(sheet.getLastRow(), 1);
  let lastCol = Math.max(sheet.getLastColumn(), 1);

  // データのある範囲から始まる結合セルが外へはみ出しているときは、結合の端まで残す
  sheet
    .getRange(1, 1, maxRows, maxCols)
    .getMergedRanges()
    .forEach((m) => {
      if (m.getRow() <= lastRow && m.getColumn() <= lastCol) {
        lastRow = Math.max(lastRow, m.getLastRow());
        lastCol = Math.max(lastCol, m.getLastColumn());
      }
    });

  // ③ で転記する範囲(A1:AS35)は、空白でも消さない
  if (sheet.getName() === CONFIG.EXPORT.SOURCE_SHEET) {
    const r = sheet.getRange(CONFIG.EXPORT.SOURCE_RANGE);
    lastRow = Math.max(lastRow, r.getLastRow());
    lastCol = Math.max(lastCol, r.getLastColumn());
  }

  // 固定した行・列しか残らない状態にはできないので、それより1つ多く残す
  const keepRows = Math.min(maxRows, Math.max(lastRow + keepEmptyRows, sheet.getFrozenRows() + 1));
  const keepCols = Math.min(maxCols, Math.max(lastCol, sheet.getFrozenColumns() + 1));
  return { maxRows, maxCols, keepRows, keepCols, delRows: maxRows - keepRows, delCols: maxCols - keepCols };
}

// ============================================================
// ② 列全体参照の置換
// ============================================================

// 列全体参照(例: '編明細'!$R:$R、編明細!R:R、$A:$C)を見つける正規表現
//   1: シート名と「!」  2,4: 「$」  3,5: 列の英字
const COLUMN_REF_RE =
  /('(?:[^']|'')+'!|[^\s'!(),;:&=+\-*\/<>^"{}\[\]\u0001]+!)?(?<![A-Za-z0-9_$.])(\$?)([A-Z]{1,3}):(\$?)([A-Z]{1,3})(?![A-Za-z0-9_(.])/g;

function previewColumnRefs() {
  columnRefs_(false);
}

function runColumnRefs() {
  if (!confirm_('数式の列全体参照(例 $R:$R)を、範囲指定(例 $R$1:$R$8617)に置き換えます。\nバックアップは取りましたか?')) return;
  columnRefs_(true);
}

function columnRefs_(apply) {
  const started = Date.now();
  const ss = SpreadsheetApp.getActive();
  toast_('数式を調べています…(数分かかることがあります)');

  const scan = collectFormulas_(ss, started);
  if (!scan) {
    notify_('② 列全体参照', ['時間内にすべての数式を調べられませんでした。何も変更していません。']);
    return;
  }
  const analysis = analyzeColumnRefs_(scan.formulas);

  // 参照先シートごとの最終行(データの最終行 + 余白)
  const limits = {};
  const missing = [];
  for (const ref of analysis.refs) {
    if (limits[ref.sheetName] !== undefined || missing.indexOf(ref.sheetName) >= 0) continue;
    const target = ss.getSheetByName(ref.sheetName);
    if (target) limits[ref.sheetName] = Math.max(target.getLastRow(), 1) + CONFIG.REF_HEADROOM_ROWS;
    else missing.push(ref.sheetName);
  }

  const safe = analysis.refs.filter((r) => r.safe && limits[r.sheetName] !== undefined);
  const unsafe = analysis.refs.filter((r) => !r.safe);
  safe.forEach((r) => (r.newText = boundedRef_(r, limits[r.sheetName])));

  const lines = [];
  Object.keys(limits).forEach((name) => lines.push('参照先「' + name + '」: 1〜' + limits[name] + ' 行目に限定'));
  lines.push('');
  safe.forEach((r) => lines.push(r.count + ' か所  ' + r.text + '  →  ' + r.newText));
  if (unsafe.length) {
    lines.push('', '【自動では置き換えないもの(文字列の中や、より長い参照の一部にも出てくるため)】');
    unsafe.forEach((r) => lines.push(r.count + ' か所  ' + r.text + '  (例: ' + r.example + ')'));
  }
  if (missing.length) lines.push('', '【参照先シートが見つからないもの】 ' + missing.join(', '));
  if (analysis.noSheetCount) {
    lines.push('', '【シート名のない列全体参照 ' + analysis.noSheetCount + ' か所(手作業で確認してください)】');
    analysis.noSheetExamples.forEach((e) => lines.push(e));
  }

  if (!apply) {
    const total = safe.reduce((s, r) => s + r.count, 0);
    notify_('② 列全体参照の確認', ['【確認だけです。まだ何も変更していません】', '置き換える参照: ' + safe.length + ' 種類 / 合計 ' + total + ' か所', ''].concat(lines));
    return;
  }

  // 参照先シートの行数が足りなければ、範囲の終わりまで行を足しておく
  Object.keys(limits).forEach((name) => {
    const target = ss.getSheetByName(name);
    const max = target.getMaxRows();
    if (max < limits[name]) target.insertRowsAfter(max, limits[name] - max);
  });

  let replaced = 0;
  let done = 0;
  for (const r of safe) {
    if (Date.now() - started > CONFIG.TIME_LIMIT_MS) break;
    replaced += ss.createTextFinder(r.text).matchCase(true).matchFormulaText(true).replaceAllWith(r.newText);
    done++;
  }
  const head = ['置き換えました: ' + done + ' / ' + safe.length + ' 種類、合計 ' + replaced + ' か所'];
  if (done < safe.length) head.push('時間切れのため途中で止めました。同じメニューをもう一度実行すると続きから処理します。');
  notify_('② 列全体参照の置換', head.concat([''], lines));
}

/** 全シートの数式を集める。同じ数式はまとめて数える(人別シートは同じ数式が多いため) */
function collectFormulas_(ss, started) {
  const formulas = new Map(); // 数式 → { count, example }
  for (const sheet of ss.getSheets()) {
    if (Date.now() - started > CONFIG.TIME_LIMIT_MS) return null;
    if (sheet.getName() === CONFIG.EXPORT.TARGET_SHEET || sheet.getLastRow() === 0) continue;
    const grid = sheet.getDataRange().getFormulas();
    for (let r = 0; r < grid.length; r++) {
      for (let c = 0; c < grid[r].length; c++) {
        const f = grid[r][c];
        if (!f || f.indexOf(':') < 0) continue;
        const entry = formulas.get(f);
        if (entry) entry.count++;
        else formulas.set(f, { count: 1, example: sheet.getName() + '!' + columnLetter_(c + 1) + (r + 1) });
      }
    }
  }
  return { formulas };
}

/**
 * 列全体参照を集計する。
 * safe = 文字列の中やより長い参照の一部には出てこない(TextFinder で一括置換しても他を壊さない)
 */
function analyzeColumnRefs_(formulas) {
  const refs = new Map();
  let noSheetCount = 0;
  const noSheetExamples = [];

  formulas.forEach((info, f) => {
    for (const m of findColumnRefs_(f)) {
      if (!m.sheetName) {
        noSheetCount += info.count;
        if (noSheetExamples.length < 10) noSheetExamples.push(info.example + '  ' + m.text);
        continue;
      }
      const ref = refs.get(m.text) || Object.assign({ count: 0, example: info.example }, m);
      ref.count += info.count;
      refs.set(m.text, ref);
    }
  });

  // 同じ文字列が「正しい参照」として数えた回数より多く出てくるなら、一括置換は危ない
  refs.forEach((ref) => {
    let raw = 0;
    formulas.forEach((info, f) => {
      if (f.indexOf(ref.text) >= 0) raw += countOccurrences_(f, ref.text) * info.count;
    });
    ref.safe = raw === ref.count;
  });

  return { refs: Array.from(refs.values()), noSheetCount, noSheetExamples };
}

function findColumnRefs_(formula) {
  const masked = maskStrings_(formula);
  const re = new RegExp(COLUMN_REF_RE.source, 'g');
  const out = [];
  let m;
  while ((m = re.exec(masked)) !== null) {
    const prefix = m[1] || '';
    out.push({
      text: m[0],
      prefix,
      sheetName: prefix ? unquoteSheetName_(prefix.slice(0, -1)) : '',
      d1: m[2],
      c1: m[3],
      d2: m[4],
      c2: m[5],
    });
  }
  return out;
}

/** 数式の中の "文字列" を、同じ長さの記号に置き換える(文字列の中は参照として扱わない) */
function maskStrings_(formula) {
  return formula.replace(/"(?:[^"]|"")*"/g, (s) => '"' + '\u0001'.repeat(s.length - 2) + '"');
}

function unquoteSheetName_(name) {
  return name.charAt(0) === "'" ? name.slice(1, -1).replace(/''/g, "'") : name;
}

function boundedRef_(ref, lastRow) {
  return ref.prefix + ref.d1 + ref.c1 + '$1:' + ref.d2 + ref.c2 + '$' + lastRow;
}

function countOccurrences_(text, part) {
  let n = 0;
  for (let i = text.indexOf(part); i >= 0; i = text.indexOf(part, i + part.length)) n++;
  return n;
}

// ============================================================
// ③ 外部連携用シート(値のみ)
// ============================================================

function syncExternalSheet() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) return; // 別の更新が動いている
  try {
    const ss = SpreadsheetApp.getActive();
    const cfg = CONFIG.EXPORT;
    const source = ss.getSheetByName(cfg.SOURCE_SHEET);
    if (!source) throw new Error('シート「' + cfg.SOURCE_SHEET + '」が見つかりません');

    const range = source.getRange(cfg.SOURCE_RANGE);
    const values = range.getValues();
    const formats = range.getNumberFormats();
    const rows = values.length;
    const cols = values[0].length;
    const stamp = '最終更新: ' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy/MM/dd HH:mm:ss');

    let target = ss.getSheetByName(cfg.TARGET_SHEET);
    if (!target) target = ss.insertSheet(cfg.TARGET_SHEET, ss.getNumSheets());
    writeValues_(target, values, formats, stamp);

    if (cfg.PUSH_SPREADSHEET_ID) {
      const other = SpreadsheetApp.openById(cfg.PUSH_SPREADSHEET_ID);
      const pushed = other.getSheetByName(cfg.PUSH_SHEET_NAME) || other.insertSheet(cfg.PUSH_SHEET_NAME);
      writeValues_(pushed, values, formats, stamp);
    }

    PropertiesService.getScriptProperties().setProperty('lastSync', String(Date.now()));
    toast_('「' + cfg.TARGET_SHEET + '」を更新しました(' + rows + '行×' + cols + '列)');
  } finally {
    lock.releaseLock();
  }
}

/** 値と表示形式だけを書き込む。右に1列あけて最終更新日時を入れる。シートは必要な大きさだけにする */
function writeValues_(sheet, values, formats, stamp) {
  const rows = values.length;
  const cols = values[0].length;
  fitSheet_(sheet, rows, cols + 2);
  const out = sheet.getRange(1, 1, rows, cols);
  out.setValues(values);
  out.setNumberFormats(formats);
  sheet.getRange(1, cols + 2).setValue(stamp);
}

function fitSheet_(sheet, rows, cols) {
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  if (maxRows < rows) sheet.insertRowsAfter(maxRows, rows - maxRows);
  if (maxRows > rows) sheet.deleteRows(rows + 1, maxRows - rows);
  if (maxCols < cols) sheet.insertColumnsAfter(maxCols, cols - maxCols);
  if (maxCols > cols) sheet.deleteColumns(cols + 1, maxCols - cols);
}

/** 手で編集したとき用(インストール型トリガーから呼ばれる)。1分以内の連続編集では更新しない */
function syncOnEdit(e) {
  const last = Number(PropertiesService.getScriptProperties().getProperty('lastSync') || 0);
  if (Date.now() - last < 60 * 1000) return;
  if (e && e.range && e.range.getSheet().getName() === CONFIG.EXPORT.TARGET_SHEET) return;
  syncExternalSheet();
}

function enableAutoSync() {
  removeSyncTriggers_();
  const ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('syncExternalSheet').timeBased().everyMinutes(CONFIG.EXPORT.SYNC_EVERY_MINUTES).create();
  ScriptApp.newTrigger('syncOnEdit').forSpreadsheet(ss).onEdit().create();
  syncExternalSheet();
  notify_('③ 自動更新をオンにしました', [
    '・' + CONFIG.EXPORT.SYNC_EVERY_MINUTES + ' 分ごとに「' + CONFIG.EXPORT.TARGET_SHEET + '」を更新します',
    '・セルを手で編集したときも更新します(1分以内の連続編集はまとめます)',
    '・止めるときは「③ 自動更新をオフにする」',
  ]);
}

function disableAutoSync() {
  const n = removeSyncTriggers_();
  notify_('③ 自動更新をオフにしました', ['削除したトリガー: ' + n + ' 個']);
}

function removeSyncTriggers_() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (['syncExternalSheet', 'syncOnEdit'].indexOf(t.getHandlerFunction()) >= 0) {
      ScriptApp.deleteTrigger(t);
      n++;
    }
  });
  return n;
}

// ============================================================
// 共通
// ============================================================

/** 削除・置換の前の確認。Apps Script の画面から直接実行したときは、安全のため実行しない */
function confirm_(message) {
  try {
    const ui = SpreadsheetApp.getUi();
    return ui.alert('確認', message, ui.ButtonSet.YES_NO) === ui.Button.YES;
  } catch (e) {
    console.log('削除・置換はスプレッドシートのメニュー「🛠 軽量化ツール」から実行してください。');
    return false;
  }
}

/** 結果を表示する(メニューから実行したときは画面に、それ以外は実行ログに) */
function notify_(title, lines) {
  console.log(title + '\n' + lines.join('\n'));
  const shown = lines.length > 30 ? lines.slice(0, 30).concat(['…ほか ' + (lines.length - 30) + ' 行(全部は Apps Script の「実行数」で見られます)']) : lines;
  try {
    SpreadsheetApp.getUi().alert(title, shown.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    // トリガーや Apps Script の画面から実行したときは画面を出せない
  }
}

function toast_(message) {
  try {
    SpreadsheetApp.getActive().toast(message, '軽量化ツール', 5);
  } catch (e) {
    // 画面がないときは何もしない
  }
}

function columnLetter_(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
