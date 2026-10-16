/**
 * 歩合表 軽量化ツール v2 —— 診断版(対象のファイルは変更しない)
 *
 * このコードの置き場所
 *   新しく作る軽いスプレッドシート「軽量化ツール_診断」の 拡張機能 → Apps Script。
 *   本番の歩合表には貼りません(本番にもとからある GAS とぶつからないようにするため)。
 *
 * この版でできること
 *   1. 対象ファイルの登録 …… 本番・保管用バックアップ・試験用コピーの URL をメニューから登録する
 *   2. バックアップの確認 …… 3つのファイルが開けるか、シートの構成が同じかを調べる
 *   3. 軽い診断 …………… シート名・sheetId・表示状態・行数・列数などだけを一覧にする
 *   小さな試験 ………… 代表シート1枚だけを読み、読み取りにかかった時間を測る(予約を作らない)
 *   4. 詳しい診断 ………… シートを少しずつ読み、数式・参照・入力規則などを数える(止めても続きから)
 *   8〜10. 比較基準 ……… 大事な集計の値を保管用(変更前の状態)から記録し、試験用コピーと比べる
 *   11. 所要時間 ………… 試験用コピーの「変更後」と「変更前」のシートで、再計算にかかる時間を比べる
 *
 *   自動の続き(時間で区切ったあと、予約で続きを実行する)は、最初はオフです。
 *   小さな試験と、手動での「4.」1回分を確かめてから、メニューでオンにします。
 *
 * 安全のための約束
 *   ・1〜10 は、本番・保管用・試験用コピーのどれにも書き込みません(読むだけ)。
 *     書き込むのは、このコードを入れた「軽量化ツール_診断」ファイルだけです。
 *   ・11 だけは試験用コピーに書き込みます。指定のセルの数式を同じ値になる書き方に一時的に替えて
 *     再計算させ、毎回もとの数式に戻します。本番・保管用には書き込みません(ID が同じなら止まる)。
 *   ・ファイルの ID はコードに書かず、メニューから登録します(スクリプト プロパティに保存)。
 */

const LT = {
  VERSION: 'v2.3(2026-10-16)診断版+所要時間(読み取りと再計算待ちを分ける)',
  MENU: '🛠 軽量化ツール',
  REPRESENTATIVE_SHEET: '001', // 構造を比べるときの基準にするシート
  CHUNK_CELLS: 50000, // 1回に読むセルの数の上限(行数 = これ ÷ 列数)
  MIN_CHUNK_ROWS: 25, // 読み取りに失敗したら、ここまで行数を減らして読み直す
  MAX_RETRIES: 6, // 同じ場所で失敗してよい回数。超えたら「調べられなかった行」として先へ進む
  TIME_LIMIT_MS: 4.5 * 60 * 1000, // 1回の実行をここで区切る(Google の上限は1回6分)
  FINALIZE_RESERVE_MS: 60 * 1000, // 最後のまとめのために残しておく時間
  FLUSH_EVERY_SHEETS: 10, // 診断結果をこの枚数ごとに書き出す
  FLUSH_EVERY_MS: 60 * 1000, // または、この時間ごとに書き出す
  // 自動の続きのオン・オフは、メニューで切り替える(スクリプト プロパティ LT_AUTO。最初はオフ)
  CONTINUE_AFTER_MS: 60 * 1000, // 続きを実行するまでの待ち時間
  READ_VALUES: true, // 計算結果の値も読む(false にすると数式と入力規則だけを調べる)
  SCAN_VALIDATIONS: true, // 入力規則を数える
  ROW_FINGERPRINT_MAX_ROWS: 3000, // この行数以下のシートは、行ごとの数式の指紋を残す(構造の比較用)
  MAX_EXAMPLES: 10, // 1シート・1種類あたりに残す例の数
  SNIPPET: 200, // 例として残す数式の文字数
  CELL_LIMIT: 10000000, // Google スプレッドシートのセル数の上限
  BASELINE_NAME_HINT: /(支給|給与|売上|各店|集計|合計)/, // 比較対象の候補にするシート名
  BASELINE_CHUNK_CELLS: 20000,
  COMPARE_MAX_DETAILS: 500, // 比較結果に書く「違うセル」の最大数
  // 11. 所要時間: 試験用コピーの中の「変更後」と「変更前」のシート(同じ形)を交互に再計算させる
  TIMING: { SHEETS: ['001', '001_変更前'], CELL: 'K6', READ_RANGE: 'A1:AR77', ROUNDS: 3 },
};

const LT_ROLES = {
  PROD: { key: 'LT_PROD_ID', label: '本番' },
  BACKUP: { key: 'LT_BACKUP_ID', label: '保管用バックアップ' },
  TEST: { key: 'LT_TEST_ID', label: '試験用コピー' },
};

// このツールのファイルに作るシート
const S = {
  GUIDE: 'はじめに',
  SETUP: '設定とバックアップ確認',
  FILE: 'ファイル全体',
  LIST: 'シート一覧',
  DETAIL: 'シート詳細',
  FN: '関数の数',
  COLREF: '列全体参照',
  REFS: '参照関係',
  EXT: '外部・動的参照',
  STRUCT: '構造の比較',
  SUMMARY: '診断サマリー',
  SMALL: '小さな試験',
  LOG: '進捗ログ',
  TARGETS: '比較対象',
  BASE: '比較基準一覧',
  RESULT: '比較結果',
  TIMING: '所要時間',
  WORK: '_作業_状態',
  ROWFP: '_作業_行指紋',
};
const BASE_DATA_PREFIX = '基準データ_';
const CONTINUE_HANDLER = 'lt_diagContinueTrigger';

// 数式の中で探す関数
const FN_COLS = ['SUMIFS', 'COUNTIFS', 'SUMIF', 'COUNTIF', 'FILTER', 'QUERY', 'ARRAYFORMULA', 'SUMPRODUCT', 'VLOOKUP',
  'XLOOKUP', 'INDEX', 'MATCH', 'LET', 'LAMBDA', 'IF', 'IFS', 'IFERROR', 'SUBSTITUTE', 'SEARCH', 'TEXT', 'SUM', 'TODAY',
  'NOW', 'RAND', 'INDIRECT', 'OFFSET', 'IMPORTRANGE'];
const COND_FUNCS = ['SUMIFS', 'COUNTIFS', 'AVERAGEIFS', 'MAXIFS', 'MINIFS', 'SUMIF', 'COUNTIF', 'AVERAGEIF'];
const ARRAY_FUNCS = ['FILTER', 'SUMPRODUCT', 'ARRAYFORMULA', 'QUERY', 'MMULT', 'SORT', 'SORTN', 'UNIQUE', 'MAP', 'BYROW',
  'BYCOL', 'REDUCE', 'SCAN', 'XLOOKUP', 'VLOOKUP', 'MATCH'];
const TEXT_FUNCS = ['SUBSTITUTE', 'TEXT', 'VALUE', 'LEFT', 'RIGHT', 'MID', 'UPPER', 'LOWER', 'TRIM', 'CLEAN', 'SEARCH',
  'FIND', 'LEN', 'EXACT', 'REGEXMATCH', 'REGEXEXTRACT', 'REGEXREPLACE', 'DATEVALUE', 'TO_TEXT', 'YEAR', 'MONTH', 'DAY',
  'WEEKDAY', 'ISNUMBER', 'ISBLANK', 'ISTEXT', 'N'];
const VOLATILE_FUNCS = ['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'RANDARRAY'];
const DYNAMIC_FUNCS = ['INDIRECT', 'OFFSET'];
const EXTERNAL_FUNCS = ['IMPORTRANGE', 'IMPORTDATA', 'IMPORTHTML', 'IMPORTXML', 'IMPORTFEED', 'GOOGLEFINANCE', 'IMAGE',
  'GOOGLETRANSLATE', 'DETECTLANGUAGE'];

// ============================================================
// メニュー
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(LT.MENU)
    .addItem('はじめに(手順を表示)', 'lt_showGuide')
    .addSeparator()
    .addItem('1. 対象ファイルを登録する', 'lt_setup')
    .addItem('2. バックアップと試験用コピーを確認する(変更しない)', 'lt_checkBackup')
    .addSeparator()
    .addItem('3. 軽い診断:シート一覧(変更しない)', 'lt_diagLight')
    .addItem('3b. 小さな試験:代表シート1枚だけ読む(変更しない)', 'lt_smallTest')
    .addItem('4. 詳しい診断を始める/続ける(変更しない)', 'lt_diagStart')
    .addItem('5. 診断の進み具合を見る', 'lt_diagProgress')
    .addItem('6. 診断を止める', 'lt_diagStop')
    .addItem('7. 詳しい診断を最初からやり直す', 'lt_diagRestart')
    .addSeparator()
    .addItem('自動の続きをオンにする(小さな試験と手動1回の確認後)', 'lt_autoOn')
    .addItem('自動の続きをオフにする', 'lt_autoOff')
    .addSeparator()
    .addItem('8. 比較する場所の候補を作る', 'lt_baselineCandidates')
    .addItem('9. 比較基準を保管用(変更前)から記録する(変更しない)', 'lt_baselineRecord')
    .addItem('10. 試験用コピーを比較基準と比べる(変更しない)', 'lt_baselineCompare')
    .addSeparator()
    .addItem('11. 所要時間を測る(試験用コピーだけ・入れ直して戻す)', 'lt_timing')
    .addToUi();
}

function lt_showGuide() {
  const sh = ensureGuideSheet_();
  activate_(sh);
  notify_('はじめに', ['シート「' + S.GUIDE + '」に手順を書きました。上から順に進めてください。', '', 'ツールの版: ' + LT.VERSION]);
}

function ensureGuideSheet_() {
  const lines = [
    ['歩合表 軽量化ツール ' + LT.VERSION],
    [''],
    ['1〜10 は「読むだけ」です。本番・保管用・試験用コピーには何も書き込みません。結果はこのファイルに書きます。'],
    ['11 だけは試験用コピーのセルを一時的に入れ直し、毎回もとに戻します(本番・保管用には書き込みません)。'],
    [''],
    ['手順(メニュー「' + LT.MENU + '」から)'],
    ['1. 対象ファイルを登録する …… 本番・保管用バックアップ・試験用コピーの URL を貼る'],
    ['2. バックアップと試験用コピーを確認する …… 3つが開けて、シートの構成が同じかを見る → シート「' + S.SETUP + '」'],
    ['3. 軽い診断 …… シートの一覧だけを作る → シート「' + S.LIST + '」「' + S.FILE + '」'],
    ['3b. 小さな試験 …… 代表シート「' + LT.REPRESENTATIVE_SHEET + '」1枚だけを読み、かかった時間を測る → シート「' + S.SMALL + '」'],
    ['4. 詳しい診断 …… 少しずつ読む。約4.5分で区切ったら、同じメニューをもう一度押して続ける → 最後に「' + S.SUMMARY + '」'],
    ['   自動の続きは最初はオフ。小さな試験と手動の1回分を確かめてから「自動の続きをオンにする」で有効にする'],
    ['5. 進み具合を見る / 6. 止める / 7. 最初からやり直す'],
    ['8〜10. 比較基準 …… 大事な集計の場所を「' + S.TARGETS + '」に書き、保管用(変更前)の値を記録する。試験用コピーと比べる'],
    ['11. 所要時間 …… 試験用コピーの「' + LT.TIMING.SHEETS.join('」と「') + '」を交互に再計算させ、かかった時間を比べる → シート「' + S.TIMING + '」'],
    [''],
    ['注意'],
    ['・詳しい診断の間は、試験用コピーを編集しないでください(結果がずれます)。'],
    ['・「6. 止める」は、今読んでいる区切りが終わったところで止まります。1回の読み取りの途中では止まりません。'],
    ['・診断で見つけた「外側のセル」「参照されていないシート」は、削除してよいという意味ではありません。'],
  ];
  const sh = ensureSheet_(S.GUIDE, null);
  sh.clearContents();
  writeRows_(sh, 1, lines, 1);
  return sh;
}

// ============================================================
// 1. 対象ファイルの登録
// ============================================================

function lt_setup() {
  const ui = uiOrNull_();
  if (!ui) return needMenu_('1. 対象ファイルを登録する');
  const props = PropertiesService.getScriptProperties();
  const toolId = SpreadsheetApp.getActive().getId();
  const steps = [
    ['PROD', '本番の「歩合表」の URL を貼ってください。\n(本番をブラウザで開いたときの、アドレスバーの文字を全部コピーして貼れば OK です)'],
    ['BACKUP', '保管用バックアップ(触らずに残しておくコピー)の URL を貼ってください。'],
    ['TEST', '試験用コピー(診断や試験に使うコピー)の URL を貼ってください。'],
  ];
  const ids = {};
  for (const [role, message] of steps) {
    const current = props.getProperty(LT_ROLES[role].key);
    const hint = current ? '\n\n空欄のまま OK を押すと、今の登録(ID の末尾 …' + current.slice(-6) + ')のままにします。' : '';
    const res = ui.prompt('1. 対象ファイルを登録する(' + LT_ROLES[role].label + ')', message + hint, ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK) return toast_('登録をやめました(何も変更していません)');
    const text = res.getResponseText().trim();
    const id = text ? parseId_(text) : current;
    if (!id) {
      ui.alert('URL から ID を読み取れませんでした。もう一度メニューの「1.」からやり直してください。');
      return;
    }
    ids[role] = id;
  }

  const problems = [];
  if (ids.PROD === ids.TEST || ids.PROD === ids.BACKUP || ids.TEST === ids.BACKUP) {
    problems.push('同じファイルが2回登録されています。本番・保管用・試験用は、それぞれ別のファイルにしてください。');
  }
  if ([ids.PROD, ids.BACKUP, ids.TEST].indexOf(toolId) >= 0) problems.push('このツールのファイル自身が登録されています。');
  if (problems.length) {
    ui.alert('登録できません', problems.join('\n'), ui.ButtonSet.OK);
    return;
  }

  const names = {};
  for (const role of ['PROD', 'BACKUP', 'TEST']) {
    try {
      names[role] = SpreadsheetApp.openById(ids[role]).getName();
    } catch (e) {
      ui.alert('開けませんでした', LT_ROLES[role].label + ' のファイルを開けませんでした。URL と権限を確認してください。\n\n' + errText_(e), ui.ButtonSet.OK);
      return;
    }
  }
  const warn = [];
  if (names.TEST.indexOf('試験') < 0) warn.push('・試験用コピーの名前に「試験」が入っていません。取り違えを防ぐため、名前に「試験用」を入れることをおすすめします。');
  if (names.BACKUP.indexOf('バックアップ') < 0 && names.BACKUP.indexOf('保管') < 0) warn.push('・保管用の名前に「バックアップ」「保管」が入っていません。');
  if (names.PROD.indexOf('試験') >= 0 || names.PROD.indexOf('バックアップ') >= 0) warn.push('・本番として登録したファイルの名前に「試験」「バックアップ」が入っています。本当に本番ですか?');

  const lines = ['PROD', 'BACKUP', 'TEST'].map((r) => '【' + LT_ROLES[r].label + '】' + names[r] + '\n  ID の末尾: …' + ids[r].slice(-6));
  const msg = '次の内容で登録します。よろしいですか?\n\n' + lines.join('\n') + (warn.length ? '\n\n' + warn.join('\n') : '') +
    '\n\n登録しても、どのファイルの中身も変わりません。';
  if (ui.alert('登録の確認', msg, ui.ButtonSet.YES_NO) !== ui.Button.YES) return toast_('登録をやめました');

  props.setProperties({ LT_PROD_ID: ids.PROD, LT_BACKUP_ID: ids.BACKUP, LT_TEST_ID: ids.TEST });
  ensureGuideSheet_();
  log_('1. 登録', 'メニュー', '完了', lines.join(' / ').replace(/\n/g, ''));
  const st = loadState_();
  const extra = st && st.targetId !== ids.TEST ? ['', '※ 試験用コピーが変わったので、詳しい診断は「7. 最初からやり直す」で始め直してください。'] : [];
  notify_('登録しました', ['次は「2. バックアップと試験用コピーを確認する」を実行してください。'].concat(extra));
}

function parseId_(text) {
  const m = String(text).match(/\/d\/([-\w]{20,})/);
  if (m) return m[1];
  const t = String(text).trim();
  return /^[-\w]{20,}$/.test(t) ? t : '';
}

function getId_(role) {
  return PropertiesService.getScriptProperties().getProperty(LT_ROLES[role].key) || '';
}

/** 登録したファイルを開く。この版では、開いたファイルに書き込む処理はない */
function openRole_(role) {
  const id = getId_(role);
  if (!id) throw new Error('「' + LT_ROLES[role].label + '」が登録されていません。先に「1. 対象ファイルを登録する」を実行してください。');
  if (id === SpreadsheetApp.getActive().getId()) throw new Error('このツールのファイル自身が登録されています。');
  const ss = SpreadsheetApp.openById(id);
  return { ss: ss, id: id, name: ss.getName(), role: role, label: LT_ROLES[role].label };
}

// ============================================================
// 2. バックアップと試験用コピーの確認
// ============================================================

function lt_checkBackup() {
  withLock_('2. バックアップの確認', () => {
    const files = {};
    for (const role of ['PROD', 'BACKUP', 'TEST']) {
      const t0 = Date.now();
      try {
        const f = openRole_(role);
        files[role] = { ok: true, name: f.name, id: f.id, meta: sheetMeta_(f.ss) };
      } catch (e) {
        files[role] = { ok: false, name: '', id: getId_(role), error: errText_(e) };
      }
      files[role].sec = Math.round((Date.now() - t0) / 100) / 10;
    }

    const diffs = [];
    const diffCount = {};
    for (const role of ['BACKUP', 'TEST']) {
      if (!files[role].ok || !files.PROD.ok) {
        diffCount[role] = '比べられない';
        continue;
      }
      const d = compareMeta_(files.PROD.meta, files[role].meta);
      diffCount[role] = d.length;
      d.forEach((x) => diffs.push([LT_ROLES[role].label, x.sheet, x.text]));
    }

    const now = nowText_();
    const header = ['役割', 'ファイル名', 'ID', '開けたか', 'シート数', '非表示のシート', '確保セル数', '本番との違い(件)', '開くのにかかった秒数', '確認日時'];
    const rows = ['PROD', 'BACKUP', 'TEST'].map((role) => {
      const f = files[role];
      const m = f.meta || [];
      return [LT_ROLES[role].label, f.name, f.id, f.ok ? '開けた' : '開けない: ' + f.error, f.ok ? m.length : '', f.ok ? m.filter((x) => x.hidden).length : '',
        f.ok ? m.reduce((s, x) => s + x.rows * x.cols, 0) : '', role === 'PROD' ? '(基準)' : diffCount[role], f.sec, now];
    });
    const sh = ensureSheet_(S.SETUP, null);
    sh.clearContents();
    writeRows_(sh, 1, [header].concat(rows), header.length);
    const detailHeader = ['違いの詳細: 役割', 'シート名', '内容'];
    const detail = diffs.length ? diffs.slice(0, 500) : [['', '', '違いはありません']];
    writeRows_(sh, rows.length + 3, [detailHeader].concat(detail), 3);
    styleHeader_(sh, 1, header.length);
    styleHeader_(sh, rows.length + 3, 3);

    const ok = files.PROD.ok && files.BACKUP.ok && files.TEST.ok;
    log_('2. バックアップの確認', 'メニュー', ok ? '3つとも開けた' : '開けないファイルあり', '保管用の違い ' + diffCount.BACKUP + ' / 試験用の違い ' + diffCount.TEST);
    activate_(sh);
    const lines = ['PROD', 'BACKUP', 'TEST'].map((r) => '・' + LT_ROLES[r].label + ': ' + (files[r].ok ? '開けた(シート ' + files[r].meta.length + ' 枚)' : '開けない'));
    lines.push('', '本番とのシート構成の違い: 保管用 ' + diffCount.BACKUP + ' 件 / 試験用 ' + diffCount.TEST + ' 件');
    if (diffs.length) lines.push('違いの中身は、シート「' + S.SETUP + '」の下のほうにあります。コピーを作ったあとに本番を編集した場合は、違いが出ることがあります。');
    lines.push('', 'ここでは「開けるか」「シートの名前と大きさが同じか」だけを調べました。中の値までは比べていません(値は 8〜9 で記録します)。');
    lines.push('ブラウザで保管用バックアップを開いて、表示されることも確かめてください。');
    notify_('2. バックアップの確認', lines);
  });
}

function sheetMeta_(ss) {
  return ss.getSheets().map((sh) => ({
    name: sh.getName(),
    id: sh.getSheetId(),
    hidden: sh.isSheetHidden(),
    rows: sh.getMaxRows(),
    cols: sh.getMaxColumns(),
  }));
}

function compareMeta_(base, other) {
  const out = [];
  const byName = {};
  other.forEach((x) => (byName[x.name] = x));
  const seen = {};
  base.forEach((b) => {
    const o = byName[b.name];
    seen[b.name] = true;
    if (!o) return out.push({ sheet: b.name, text: 'コピー側にこのシートがない' });
    if (o.rows !== b.rows || o.cols !== b.cols) out.push({ sheet: b.name, text: '大きさが違う: 本番 ' + b.rows + '×' + b.cols + ' / コピー ' + o.rows + '×' + o.cols });
    if (o.hidden !== b.hidden) out.push({ sheet: b.name, text: '表示状態が違う: 本番 ' + (b.hidden ? '非表示' : '表示') + ' / コピー ' + (o.hidden ? '非表示' : '表示') });
    if (o.id !== b.id) out.push({ sheet: b.name, text: 'sheetId が違う: 本番 ' + b.id + ' / コピー ' + o.id + '(情報として記録)' });
  });
  other.forEach((o) => {
    if (!seen[o.name]) out.push({ sheet: o.name, text: '本番にないシートがコピー側にある' });
  });
  return out;
}

// ============================================================
// 3. 軽い診断(シート一覧)
// ============================================================

function lt_diagLight() {
  withLock_('3. 軽い診断', () => {
    const f = openRole_('TEST');
    const r = lightDiag_(f);
    log_('3. 軽い診断', 'メニュー', '完了', r.sheets + ' シート');
    activate_(SpreadsheetApp.getActive().getSheetByName(S.LIST));
    notify_('3. 軽い診断(変更なし)', [
      '対象: ' + f.name + '(試験用コピー)',
      'シート ' + r.sheets + ' 枚(非表示 ' + r.hidden + ' 枚)',
      '確保セル数の合計: ' + fmtNum_(r.cells) + '(上限 1,000万の ' + pct_(r.cells, LT.CELL_LIMIT) + ')',
      '',
      '結果はシート「' + S.LIST + '」と「' + S.FILE + '」です。',
      '分類の列は名前と大きさからの推定で、判断には使いません。',
    ]);
  });
}

function lightDiag_(f) {
  const ss = f.ss;
  const sheets = ss.getSheets();
  const rep = ss.getSheetByName(LT.REPRESENTATIVE_SHEET);
  const repSize = rep ? rep.getMaxRows() + '×' + rep.getMaxColumns() : '';
  const header = ['順番', 'シート名', 'sheetId', '種類', '表示', '確保 行', '確保 列', '確保セル数', '固定 行', '固定 列', '分類(推定・参考)'];
  let cells = 0;
  let hidden = 0;
  const rows = sheets.map((sh, i) => {
    const type = String(sh.getType());
    const isGrid = sh.getType() === SpreadsheetApp.SheetType.GRID;
    const r = isGrid ? sh.getMaxRows() : 0;
    const c = isGrid ? sh.getMaxColumns() : 0;
    const h = sh.isSheetHidden();
    cells += r * c;
    if (h) hidden++;
    const name = sh.getName();
    let kind = 'その他';
    if (!isGrid) kind = '表ではない(' + type + ')';
    else if (/^\d{3}$/.test(name)) kind = '3桁の数字';
    else if (repSize && r + '×' + c === repSize) kind = '代表「' + LT.REPRESENTATIVE_SHEET + '」と同じ大きさ';
    return [i + 1, name, sh.getSheetId(), type, h ? '非表示' : '表示', r, c, r * c, isGrid ? sh.getFrozenRows() : '', isGrid ? sh.getFrozenColumns() : '', kind];
  });
  writeTable_(S.LIST, header, rows);

  const fileRows = [
    ['対象ファイル', f.name, f.label],
    ['調査日時', nowText_(), ''],
    ['ロケール / タイムゾーン', safe_(() => ss.getSpreadsheetLocale()) + ' / ' + safe_(() => ss.getSpreadsheetTimeZone()), ''],
    ['再計算の設定', recalcText_(safe_(() => String(ss.getRecalculationInterval()))), 'ファイル → 設定 → 計算'],
    ['反復計算', safe_(() => (ss.isIterativeCalculationEnabled() ? 'オン' : 'オフ')), ''],
    ['シート数', sheets.length, '非表示 ' + hidden + ' 枚'],
    ['確保セル数の合計', cells, '上限 1,000万の ' + pct_(cells, LT.CELL_LIMIT)],
  ];
  const named = [];
  try {
    ss.getNamedRanges().forEach((nr) => named.push([nr.getName(), safe_(() => nr.getRange().getSheet().getName() + '!' + nr.getRange().getA1Notation())]));
    fileRows.push(['名前付き範囲', named.length, '下の一覧']);
  } catch (e) {
    fileRows.push(['名前付き範囲', '未取得', errText_(e)]);
  }
  fileRows.push(['保護(シート全体)', safe_(() => ss.getProtections(SpreadsheetApp.ProtectionType.SHEET).length), '']);
  fileRows.push(['保護(範囲)', safe_(() => ss.getProtections(SpreadsheetApp.ProtectionType.RANGE).length), '']);
  const sh = ensureSheet_(S.FILE, null);
  sh.clearContents();
  writeRows_(sh, 1, [['項目', '値', 'メモ']].concat(fileRows), 3);
  styleHeader_(sh, 1, 3);
  const start = fileRows.length + 3;
  writeRows_(sh, start, [['名前付き範囲の名前', '範囲']].concat(named.length ? named : [['(なし)', '']]), 2);
  styleHeader_(sh, start, 2);
  return { sheets: sheets.length, hidden: hidden, cells: cells, list: rows };
}

function recalcText_(v) {
  return { ON_CHANGE: '変更時(ON_CHANGE)', MINUTE: '変更時と毎分(MINUTE)', HOUR: '変更時と毎時(HOUR)' }[v] || v;
}

// ============================================================
// 4〜7. 詳しい診断(少しずつ読む・止めても続きから)
// ============================================================

const H_DETAIL = [
  ['sheetId', 'sheetId'], ['name', 'シート名'], ['hidden', '表示'], ['maxRows', '確保 行'], ['maxCols', '確保 列'], ['cells', '確保セル数'],
  ['formulas', '数式セル'], ['constants', '定数セル'], ['lastFR', '数式の最終行'], ['lastFC', '数式の最終列'],
  ['lastCR', '定数の最終行'], ['lastCC', '定数の最終列'], ['lastVR', '値がある最終行'], ['lastVC', '値がある最終列'],
  ['dv', '入力規則セル'], ['lastDvR', '入力規則の最終行'], ['lastDvC', '入力規則の最終列'],
  ['cf', '条件付き書式ルール'], ['cfCustom', 'うちカスタム数式'], ['lastCfR', '条件付き書式の最終行'], ['lastCfC', '条件付き書式の最終列'],
  ['merge', '結合'], ['lastMR', '結合の最終行'], ['lastMC', '結合の最終列'],
  ['protSheet', '保護(シート全体)'], ['protRange', '保護(範囲)'], ['lastPR', '範囲保護の最終行'], ['lastPC', '範囲保護の最終列'],
  ['named', '名前付き範囲'], ['lastNR', '名前付き範囲の最終行'], ['lastNC', '名前付き範囲の最終列'],
  ['charts', 'グラフ'], ['pivots', 'ピボット'], ['filter', 'フィルタ'], ['bandings', '交互の背景色'], ['slicers', 'スライサー'],
  ['images', 'セル上の画像'], ['drawings', '図形'], ['scripts', '図形に割り当てたスクリプト'],
  ['keepR', '内容と構造物の最終行(参考)'], ['keepC', '内容と構造物の最終列(参考)'], ['outside', 'その外側のセル数(参考・削除してよい意味ではない)'],
  ['colRefs', '列全体参照(個)'], ['openRefs', 'A2:A 形式の参照(個)'], ['catCond', '条件集計×列全体(セル)'],
  ['catArray', '配列計算×列全体(セル)'], ['catText', '列全体を1行ずつ加工(セル)'], ['volatile', '揮発性関数(セル)'],
  ['dynamic', 'INDIRECT・OFFSET(セル)'], ['external', '外部取得関数(セル)'],
  ['errors', 'エラー値(セル)'], ['loading', '読み込み中(セル)'], ['numSum', '数値の合計'], ['fpV', '値の指紋'], ['fpF', '数式の指紋'],
  ['refsTo', '参照しているシート'], ['inCount', '参照されているシート数'], ['inFrom', '参照しているシート(例)'],
  ['unscanned', '調べられなかった行'], ['noValues', '値を読めなかった行'], ['notes', 'メモ'], ['scannedAt', '調査日時'], ['sec', '処理秒数'],
];
const DI = indexOf_(H_DETAIL);
const H_FN = ['sheetId', 'シート名', '数式セル'].concat(FN_COLS).concat(['その他の関数(多い順)']);
const H_COLREF = ['参照元 sheetId', '参照元シート', '参照先シート', '参照', '種類', '個数', '列数', '参照先の確保行数', '走査の目安(推定セル数)'];
const H_REFS = ['参照元 sheetId', '参照元シート', '参照先シート', '方法', '数式セル数など', '参照先があるか'];
const H_EXT = ['sheetId', 'シート名', '種類', '関数', 'セル', '数式(先頭)', '取得元', '取得範囲'];
const H_ROWFP = ['sheetId', 'シート名', '行数', '行ごとの数式の指紋'];
const H_LOG = ['日時', '処理', '起動', '結果', '内容'];

function lt_diagStart() {
  const title = '4. 詳しい診断';
  const testId = getId_('TEST');
  if (!testId) return notify_(title, ['先に「1. 対象ファイルを登録する」を実行してください。']);
  const st = loadState_();
  if (st && st.targetId !== testId) {
    return notify_(title, ['登録している試験用コピーが、診断を始めたときと違います。', '「7. 詳しい診断を最初からやり直す」を実行してください。']);
  }
  if (st && st.phase === 'done') {
    return notify_(title, ['詳しい診断は完了しています(' + fmtTime_(st.updatedAt) + ')。', '結果はシート「' + S.SUMMARY + '」です。', 'やり直すときは「7.」を実行してください。']);
  }
  if (!st) {
    let name = '';
    try {
      name = openRole_('TEST').name;
    } catch (e) {
      return notify_(title, ['試験用コピーを開けませんでした。', errText_(e)]);
    }
    const ok = confirm_(title, [
      '対象ファイル: ' + name + '(試験用コピー)',
      '対象シート: すべて(1枚ずつ、少しずつ読みます)',
      '変更内容: なし(読むだけ)。結果はこのツールのファイルに書きます。',
      '復旧方法: 不要(対象ファイルは変更しません)',
      '',
      '約4.5分ごとに区切り、' + (autoContinueOn_() ? '1分後に自動で続きを実行します(自動の続き: オン)。' : '続きは同じメニューをもう一度押します(自動の続き: オフ)。'),
      '診断の間は、試験用コピーを編集しないでください。',
      '',
      '始めますか?',
    ]);
    if (!ok) return;
  }
  PropertiesService.getScriptProperties().deleteProperty('LT_STOP');
  runDiagSlice_('メニュー');
}

function lt_diagContinueTrigger() {
  deleteTriggers_(CONTINUE_HANDLER);
  if (PropertiesService.getScriptProperties().getProperty('LT_STOP')) {
    log_('4. 詳しい診断', '自動の続き', '止める指示があるため実行しない', '');
    return;
  }
  if (!autoContinueOn_()) {
    log_('4. 詳しい診断', '自動の続き', '自動の続きがオフのため実行しない', '');
    return;
  }
  const st = loadState_();
  if (!st || st.phase === 'done') return;
  runDiagSlice_('自動の続き');
}

function lt_diagStop() {
  PropertiesService.getScriptProperties().setProperty('LT_STOP', '1');
  const n = deleteTriggers_(CONTINUE_HANDLER);
  log_('6. 診断を止める', 'メニュー', '止める指示', '自動の続きの予約 ' + n + ' 件を取り消し');
  notify_('6. 診断を止める', [
    '止める指示を出しました。',
    '・自動の続きの予約 ' + n + ' 件を取り消しました。',
    '・いま実行中の処理は、読んでいる区切りが終わったところで止まります(1回の読み取りの途中では止まりません)。',
    '・それまでの結果は残ります。再開は「4. 詳しい診断を始める/続ける」です。',
  ]);
}

function lt_diagRestart() {
  const title = '7. 最初からやり直す';
  if (!confirm_(title, ['詳しい診断の結果(このツールのファイルの「' + S.DETAIL + '」など)を消して、最初からやり直します。', '対象のファイルは変更しません。', '', 'やり直しますか?'])) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return notify_(title, ['ほかの処理が動いています。「6. 診断を止める」を押して、少し待ってからもう一度実行してください。']);
  try {
    deleteTriggers_(CONTINUE_HANDLER);
    PropertiesService.getScriptProperties().deleteProperty('LT_STATE');
    PropertiesService.getScriptProperties().deleteProperty('LT_STOP');
    clearDiagSheets_();
    log_(title, 'メニュー', '結果を消した', '');
  } finally {
    lock.releaseLock();
  }
  runDiagSlice_('メニュー');
}

function lt_diagProgress() {
  const st = loadState_();
  if (!st) return notify_('5. 進み具合', ['詳しい診断はまだ始めていません。']);
  const reserved = ScriptApp.getProjectTriggers().filter((t) => t.getHandlerFunction() === CONTINUE_HANDLER).length;
  const stop = PropertiesService.getScriptProperties().getProperty('LT_STOP');
  const acc = st.phase === 'sheets' ? loadAcc_() : null;
  const lines = [
    '状態: ' + { sheets: 'シートを調べている', finalize: 'まとめを作っている', done: '完了' }[st.phase],
    '対象: ' + st.targetName,
    '進み具合: ' + Math.min(st.qi, st.queue.length) + ' / ' + st.queue.length + ' シート',
  ];
  if (acc && acc.sheetId === st.queue[st.qi]) lines.push('今のシート: 「' + acc.name + '」 ' + (acc.nextRow - 1) + ' / ' + acc.maxRows + ' 行');
  lines.push(
    '実行した回数: ' + st.runs + ' 回 / 読み取り ' + st.chunks + ' 回',
    '開始: ' + fmtTime_(st.startedAt) + ' / 最終更新: ' + fmtTime_(st.updatedAt),
    '自動の続き: ' + (autoContinueOn_() ? 'オン' : 'オフ') + ' / 予約: ' + (reserved ? 'あり' : 'なし') + (stop ? '(止める指示あり)' : '')
  );
  if (st.errors && st.errors.length) lines.push('', '最近のエラー:', '・' + st.errors.slice(-3).join('\n・'));
  notify_('5. 進み具合', lines);
}

// ============================================================
// 3b. 小さな試験(代表シート1枚だけを読み、かかった時間を測る)
//     予約(トリガー)を作らず、詳しい診断の途中経過(LT_STATE)も変えない
// ============================================================

const H_SMALL = ['項目', '結果', '所要ミリ秒', 'メモ'];

function lt_smallTest() {
  const title = '3b. 小さな試験(代表シート1枚)';
  const testId = getId_('TEST');
  if (!testId) return notify_(title, ['先に「1. 対象ファイルを登録する」を実行してください。']);
  const ok = confirm_(title, [
    '対象ファイル: 試験用コピー(ID の末尾 …' + testId.slice(-6) + ')',
    '対象シート: 「' + LT.REPRESENTATIVE_SHEET + '」の1枚だけ',
    '変更内容: なし(読むだけ)。結果はこのツールのファイルのシート「' + S.SMALL + '」に書きます。',
    '自動の続きの予約は作りません。詳しい診断の途中経過も変えません。',
    '復旧方法: 不要(対象ファイルは変更しません)',
    '',
    '始めますか?',
  ]);
  if (!ok) return;
  withLock_(title, () => runSmallTest_(title));
}

function runSmallTest_(title) {
  const t0 = Date.now();
  const rows = [];
  const add = (item, result, ms, memo) => rows.push([item, result, ms === '' ? '' : Math.round(ms), memo || '']);
  let failed = '';
  try {
    let s = Date.now();
    const target = openRole_('TEST');
    add('試験用コピーを開く', target.name, Date.now() - s, 'ID の末尾 …' + target.id.slice(-6));
    s = Date.now();
    const sheets = target.ss.getSheets();
    add('シートの一覧を取得', sheets.length + ' 枚', Date.now() - s, '');
    const sheet = sheets.filter((sh) => sh.getName() === LT.REPRESENTATIVE_SHEET)[0];
    if (!sheet) throw new Error('代表シート「' + LT.REPRESENTATIVE_SHEET + '」が試験用コピーにありません。');
    s = Date.now();
    const acc = newAcc_(sheet);
    add('シートの大きさ', acc.maxRows + ' 行 × ' + acc.maxCols + ' 列' + (acc.hidden ? '(非表示)' : ''), Date.now() - s, 'sheetId ' + acc.sheetId);
    if (!acc.isGrid) throw new Error('代表シートが表ではありません。');
    s = Date.now();
    const fc = fileCache_(target.ss);
    add('名前付き範囲と保護(ファイル全体)', fc.err.length ? '一部取得できず: ' + fc.err.join('・') : '取得できた', Date.now() - s, '');
    s = Date.now();
    scanStructure_(sheet, acc, fc);
    const st = acc.struct;
    add('結合・条件付き書式・グラフ・図形', '結合 ' + orBlank_(st.merge) + ' / 条件付き書式 ' + orBlank_(st.cf) + ' / グラフ ' + orBlank_(st.charts) + ' / 図形 ' + orBlank_(st.drawings), Date.now() - s, '');

    // 詳しい診断と同じ区切り(CHUNK_CELLS)で読む。API ごとの時間を分けて測る
    const ms = { f: 0, v: 0, d: 0, p: 0 };
    let reads = 0;
    let cells = 0;
    const cache = new Map();
    for (let r0 = 1; r0 <= acc.maxRows; r0 += acc.rowsPerChunk) {
      if (Date.now() - t0 > LT.TIME_LIMIT_MS) {
        acc.unscanned.push([r0, acc.maxRows]);
        acc.notes.push(r0 + '行目から先は時間の都合で読んでいない');
        break;
      }
      const n = Math.min(acc.rowsPerChunk, acc.maxRows - r0 + 1);
      const range = sheet.getRange(r0, 1, n, acc.maxCols);
      s = Date.now();
      const formulas = range.getFormulas();
      ms.f += Date.now() - s;
      s = Date.now();
      const values = acc.readValues ? range.getValues() : null;
      ms.v += Date.now() - s;
      s = Date.now();
      const dvs = LT.SCAN_VALIDATIONS ? range.getDataValidations() : null;
      ms.d += Date.now() - s;
      s = Date.now();
      processChunk_(acc, r0, formulas, values, dvs, cache);
      ms.p += Date.now() - s;
      acc.nextRow = r0 + n;
      reads++;
      cells += n * acc.maxCols;
    }
    const per = reads + ' 回 / ' + fmtNum_(cells) + ' セル';
    add('数式を読む(getFormulas)', per, ms.f, '1回あたり最大 ' + acc.rowsPerChunk + ' 行');
    add('値を読む(getValues)', acc.readValues ? per : '読んでいない', ms.v, '');
    add('入力規則を読む(getDataValidations)', LT.SCAN_VALIDATIONS ? per : '読んでいない', ms.d, '');
    add('数える(このツールの中の計算)', '', ms.p, '');
    const colKinds = Object.keys(acc.colRefs).length;
    let colTotal = 0;
    for (const k in acc.colRefs) colTotal += acc.colRefs[k];
    add('数式のセル', acc.formulas, '', '');
    add('列全体参照', colKinds + ' 種類 / ' + colTotal + ' 個', '', '「A:A」のように列全体を指す参照');
    ['SUMIFS', 'FILTER', 'TODAY', 'INDIRECT', 'IMPORTRANGE'].forEach((k) => add(k, acc.fn[k] || 0, '', 'この関数を含む数式セルの数'));
    add('エラー値のセル', acc.readValues ? acc.errors : '未取得', '', '#REF! など');
    add('読み込み中のセル', acc.readValues ? acc.loading : '未取得', '', '');
    add('入力規則のセル', LT.SCAN_VALIDATIONS ? acc.dv : '未取得', '', '');
    add('調べられなかった行', acc.unscanned.length ? acc.unscanned.map((u) => u[0] + '-' + u[1]).join(', ') : 'なし', '', '');
    add('メモ', acc.notes.length ? acc.notes.join(' / ') : 'なし', '', '');
  } catch (e) {
    failed = errText_(e);
    add('エラー', failed, '', 'ここで止まりました。対象ファイルは変更していません。');
  }
  const reserved = ScriptApp.getProjectTriggers().filter((t) => t.getHandlerFunction() === CONTINUE_HANDLER).length;
  add('自動の続きの予約', reserved + ' 件', '', 'この試験では予約を作りません');
  add('合計', '', Date.now() - t0, 'Google の上限は1回6分');
  const top = [
    ['実行日時', nowText_(), '', LT.VERSION],
    ['結果', failed ? 'エラーで止まった' : '最後まで読めた', '', ''],
  ];
  const sh = writeTable_(S.SMALL, H_SMALL, top.concat(rows));
  const props = PropertiesService.getScriptProperties();
  if (failed) props.deleteProperty('LT_SMALL_OK');
  else props.setProperty('LT_SMALL_OK', String(Date.now()));
  log_(title, 'メニュー', failed ? 'エラー: ' + failed : '完了', '合計 ' + Math.round((Date.now() - t0) / 1000) + ' 秒');
  activate_(sh);
  notify_(title + '(変更なし)', [
    failed ? 'エラーで止まりました: ' + failed : '代表シート「' + LT.REPRESENTATIVE_SHEET + '」を最後まで読めました。',
    '合計 ' + Math.round((Date.now() - t0) / 1000) + ' 秒。結果はシート「' + S.SMALL + '」です。',
    '',
    failed ? '対象ファイルは変更していません。' : '次は「4. 詳しい診断を始める/続ける」を手動で1回押してください。',
  ]);
}

// ============================================================
// 自動の続き(最初はオフ。小さな試験と手動1回分を確かめてからオンにする)
// ============================================================

function autoContinueOn_() {
  return PropertiesService.getScriptProperties().getProperty('LT_AUTO') === '1';
}

function lt_autoOn() {
  const title = '自動の続きをオンにする';
  const props = PropertiesService.getScriptProperties();
  const st = loadState_();
  const warn = [];
  if (!props.getProperty('LT_SMALL_OK')) warn.push('・「3b. 小さな試験」がまだ最後まで成功していません。');
  if (!st || !st.runs) warn.push('・「4. 詳しい診断」をまだ手動で1回も実行していません。');
  const lines = [
    'オンにすると、詳しい診断が約4.5分で区切られたとき、1分後に続きを自動で実行する予約(トリガー)を作ります。',
    '予約は同時に1件までです。診断の完了・止める指示・エラー・3回続けて進まない、のどれかで作らなくなります。',
    '消し方: 「6. 診断を止める」か「自動の続きをオフにする」。',
    '対象のファイルは変更しません(読むだけ)。',
  ];
  if (warn.length) lines.push('', '注意:', ...warn, '先に確かめることをおすすめします。');
  lines.push('', 'オンにしますか?');
  if (!confirm_(title, lines)) return;
  props.setProperty('LT_AUTO', '1');
  log_(title, 'メニュー', 'オン', warn.join(' '));
  notify_(title, ['オンにしました。', '次に「4. 詳しい診断を始める/続ける」を押すと、区切りのあと自動で続きます。']);
}

function lt_autoOff() {
  const title = '自動の続きをオフにする';
  PropertiesService.getScriptProperties().deleteProperty('LT_AUTO');
  const n = deleteTriggers_(CONTINUE_HANDLER);
  log_(title, 'メニュー', 'オフ', '予約 ' + n + ' 件を取り消し');
  notify_(title, ['オフにしました。', '・自動の続きの予約 ' + n + ' 件を取り消しました。', '・続きは「4. 詳しい診断を始める/続ける」を手動で押してください。']);
}

/** 詳しい診断を1回分(約4.5分)進める */
function runDiagSlice_(origin) {
  const title = '4. 詳しい診断';
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) {
    if (origin === '自動の続き') scheduleContinue_();
    return notify_(title, ['ほかの処理が動いています。少し待ってからもう一度実行してください。']);
  }
  const t0 = Date.now();
  let st = null;
  let target = null;
  const buf = newBuffer_();
  let doneThisRun = 0;
  let result = '';
  try {
    st = loadState_();
    target = openRole_('TEST');
    if (!st) st = initDiag_(target);
    if (st.targetId !== target.id) throw new Error('試験用コピーの登録が、診断を始めたときと違います。「7.」でやり直してください。');
    st.runs++;
    const byId = {};
    target.ss.getSheets().forEach((sh) => (byId[sh.getSheetId()] = sh));
    const cache = new Map();
    let fc = null;
    let lastFlush = Date.now();
    let estimate = 10000;
    let acc = loadAcc_(); // 途中まで調べたシートがあれば、その続きから
    st.pendingAcc = acc;
    const startMark = progressMark_(st, acc);
    let worked = false; // 1回の実行で、少なくとも1回は読み進める(見込み時間だけで止まり続けないように)

    while (st.phase === 'sheets' && st.qi < st.queue.length) {
      if (isStopRequested_()) {
        result = '止める指示で止めた';
        break;
      }
      if (worked && Date.now() - t0 + estimate * 1.5 > LT.TIME_LIMIT_MS) {
        result = '時間で区切った';
        break;
      }
      const c0 = Date.now();
      const sid = st.queue[st.qi];
      const sheet = byId[sid];
      if (!sheet) {
        pushError_(st, 'sheetId ' + sid + ' のシートが見つからない(削除・作り直しの可能性)');
        st.qi++;
        continue;
      }
      if (!acc || acc.sheetId !== sid) {
        if (!fc) fc = fileCache_(target.ss);
        acc = newAcc_(sheet);
        scanStructure_(sheet, acc, fc);
      }
      if (acc.nextRow <= acc.maxRows) {
        scanNextChunk_(sheet, acc, cache);
        st.chunks++;
      }
      acc.ms += Date.now() - c0;
      worked = true;
      if (acc.nextRow > acc.maxRows) {
        bufferSheet_(buf, acc);
        acc = null;
        st.qi++;
        doneThisRun++;
      }
      st.pendingAcc = acc;
      estimate = Math.max(3000, Date.now() - c0);
      if (buf.count >= LT.FLUSH_EVERY_SHEETS || Date.now() - lastFlush > LT.FLUSH_EVERY_MS) {
        checkpoint_(buf, st);
        lastFlush = Date.now();
      }
    }
    checkpoint_(buf, st);

    if (st.phase === 'sheets' && st.qi >= st.queue.length) st.phase = 'finalize';
    st.noProgress = st.phase === 'sheets' && progressMark_(st, acc) === startMark ? (st.noProgress || 0) + 1 : 0;
    if (st.phase === 'finalize' && !isStopRequested_()) {
      if (Date.now() - t0 + LT.FINALIZE_RESERVE_MS < LT.TIME_LIMIT_MS) {
        finalizeDiag_(st, target);
        st.phase = 'done';
        result = '完了';
      } else {
        result = '時間で区切った(次はまとめ)';
      }
    }
    saveState_(st);
  } catch (e) {
    result = 'エラー: ' + errText_(e);
    if (st) {
      pushError_(st, errText_(e));
      saveStateQuiet_(st);
    }
  } finally {
    lock.releaseLock();
  }

  const finished = st && st.phase === 'done';
  const stopped = isStopRequested_();
  let next = '';
  const stuck = st && st.noProgress >= 3;
  if (stuck) result += '(3回続けて進まなかったので、自動の続きを止めた)';
  if (!finished && !stopped && !stuck && st && autoContinueOn_() && result.indexOf('エラー') < 0) {
    scheduleContinue_();
    next = '1分後に自動で続きを実行します。';
  }
  const progress = st ? Math.min(st.qi, st.queue.length) + ' / ' + st.queue.length + ' シート' : '';
  log_(title, origin, result, '今回 ' + doneThisRun + ' シート / 累計 ' + progress + ' / ' + Math.round((Date.now() - t0) / 1000) + ' 秒');

  const lines = ['対象: ' + (target ? target.name : '(開けなかった)'), '今回調べたシート: ' + doneThisRun + ' 枚 / 累計 ' + progress, '結果: ' + result];
  if (finished) {
    lines.push('', '結果はシート「' + S.SUMMARY + '」です。ほかのシートに細かい数字があります。');
    activate_(SpreadsheetApp.getActive().getSheetByName(S.SUMMARY));
  } else if (stopped) {
    lines.push('', '止める指示があるので、続きは実行しません。再開は「4.」です。');
  } else if (next) {
    lines.push('', next, 'この画面は閉じてかまいません。進み具合は「5.」で見られます。');
  } else {
    lines.push('', '続きは「4. 詳しい診断を始める/続ける」をもう一度押してください。');
  }
  notify_(title + '(変更なし)', lines);
}

/** 進み具合の目印(どのシートの何行目まで進んだか) */
function progressMark_(st, acc) {
  return st.qi + ':' + (acc ? acc.nextRow : 0);
}

function initDiag_(target) {
  clearDiagSheets_();
  const light = lightDiag_(target);
  const queue = light.list.map((r) => r[2]);
  const st = {
    v: 2,
    targetId: target.id,
    targetName: target.name,
    phase: 'sheets',
    queue: queue,
    qi: 0,
    runs: 0,
    chunks: 0,
    errors: [],
    startedAt: Date.now(),
    updatedAt: Date.now(),
    pendingAcc: null,
  };
  saveAcc_(null);
  saveState_(st);
  return st;
}

function clearDiagSheets_() {
  const ss = SpreadsheetApp.getActive();
  [S.DETAIL, S.FN, S.COLREF, S.REFS, S.EXT, S.STRUCT, S.SUMMARY, S.ROWFP, S.WORK].forEach((n) => {
    const sh = ss.getSheetByName(n);
    if (sh) sh.clearContents();
  });
}

/** 結果を書いてから状態を保存する(途中で止まっても、最後に保存した状態から重複なしでやり直せる) */
function checkpoint_(buf, st) {
  flushBuffer_(buf);
  saveAcc_(st.pendingAcc || null);
  saveState_(st);
}

function newAcc_(sheet) {
  const isGrid = sheet.getType() === SpreadsheetApp.SheetType.GRID;
  const maxRows = isGrid ? sheet.getMaxRows() : 0;
  const maxCols = isGrid ? sheet.getMaxColumns() : 0;
  return {
    sheetId: sheet.getSheetId(),
    name: sheet.getName(),
    hidden: sheet.isSheetHidden(),
    isGrid: isGrid,
    maxRows: maxRows,
    maxCols: maxCols,
    frozenRows: isGrid ? sheet.getFrozenRows() : 0,
    frozenCols: isGrid ? sheet.getFrozenColumns() : 0,
    rowsPerChunk: Math.max(1, Math.floor(LT.CHUNK_CELLS / Math.max(1, maxCols))),
    nextRow: 1,
    retries: 0,
    ms: 0,
    readValues: LT.READ_VALUES,
    formulas: 0,
    constants: 0,
    lastF: [0, 0],
    lastC: [0, 0],
    lastV: [0, 0],
    dv: 0,
    lastDv: [0, 0],
    errors: 0,
    loading: 0,
    numSum: 0,
    fpF: '',
    fpV: '',
    rowFp: maxRows > 0 && maxRows <= LT.ROW_FINGERPRINT_MAX_ROWS ? [] : null,
    fn: {},
    colRefs: {},
    openRefs: {},
    sheetRefs: {},
    cat: { cond: 0, array: 0, text: 0, volatile: 0, dynamic: 0, external: 0 },
    exCount: {},
    ex: [],
    struct: {},
    unscanned: [],
    noValues: [],
    notes: isGrid ? [] : ['表ではないシートのため、セルは調べていない'],
  };
}

/** シートの構造物(結合・条件付き書式・グラフなど)を1回だけ調べる */
function scanStructure_(sheet, acc, fc) {
  const s = acc.struct;
  const tryGet = (label, fn) => {
    try {
      fn();
    } catch (e) {
      acc.notes.push(label + ' を取得できなかった');
    }
  };
  const last = (key, r, c) => {
    if (!s[key]) s[key] = [0, 0];
    s[key][0] = Math.max(s[key][0], r);
    s[key][1] = Math.max(s[key][1], c);
  };
  const named = fc.named[acc.name] || [];
  s.named = named.length;
  named.forEach((n) => last('namedLast', n[1], n[2]));
  s.protSheet = fc.protSheet[acc.name] || 0;
  const prot = fc.protRange[acc.name] || [];
  s.protRange = prot.length;
  prot.forEach((p) => last('protLast', p[0], p[1]));
  if (fc.err.length) acc.notes.push('ファイル全体の ' + fc.err.join('・') + ' を取得できなかった');

  tryGet('グラフ', () => {
    const charts = sheet.getCharts();
    s.charts = charts.length;
    s.chartSources = [];
    charts.forEach((ch) => {
      tryGet('グラフの位置', () => {
        const ci = ch.getContainerInfo();
        last('anchorLast', ci.getAnchorRow(), ci.getAnchorColumn());
      });
      ch.getRanges().forEach((r) => s.chartSources.push(r.getSheet().getName()));
    });
  });
  if (!acc.isGrid) return;

  tryGet('結合', () => {
    const ms = sheet.getRange(1, 1, acc.maxRows, acc.maxCols).getMergedRanges();
    s.merge = ms.length;
    ms.forEach((m) => last('mergeLast', m.getLastRow(), m.getLastColumn()));
  });
  tryGet('条件付き書式', () => {
    const rules = sheet.getConditionalFormatRules();
    s.cf = rules.length;
    s.cfCustom = 0;
    rules.forEach((rule) => {
      rule.getRanges().forEach((r) => last('cfLast', r.getLastRow(), r.getLastColumn()));
      const cond = rule.getBooleanCondition();
      if (cond && cond.getCriteriaType() === SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) {
        s.cfCustom++;
        const f = String(cond.getCriteriaValues()[0] || '');
        if (/INDIRECT\s*\(/i.test(f)) addEx_(acc, '条件付き書式のINDIRECT', 'INDIRECT', rule.getRanges().map((r) => r.getA1Notation()).join(','), f, '', '');
      }
    });
  });
  tryGet('ピボット', () => {
    const pts = sheet.getPivotTables();
    s.pivots = pts.length;
    s.pivotSources = pts.map((p) => p.getSourceDataRange().getSheet().getName());
  });
  tryGet('フィルタ', () => {
    const fl = sheet.getFilter();
    s.filter = fl ? fl.getRange().getA1Notation() : '';
    if (fl) last('filterLast', fl.getRange().getLastRow(), fl.getRange().getLastColumn());
  });
  tryGet('交互の背景色', () => (s.bandings = sheet.getBandings().length));
  tryGet('スライサー', () => (s.slicers = sheet.getSlicers().length));
  tryGet('画像', () => {
    const imgs = sheet.getImages();
    s.images = imgs.length;
    imgs.forEach((im) => tryGet('画像の位置', () => last('anchorLast', im.getAnchorCell().getRow(), im.getAnchorCell().getColumn())));
  });
  tryGet('図形', () => {
    const ds = sheet.getDrawings();
    s.drawings = ds.length;
    s.scripts = [];
    ds.forEach((d) => {
      const action = safe_(() => d.getOnAction());
      if (action && action !== '未取得') {
        s.scripts.push(action);
        addEx_(acc, '図形に割り当てたスクリプト', action, '', '既存の GAS の関数「' + action + '」を呼ぶボタン', '', '');
      }
      tryGet('図形の位置', () => last('anchorLast', d.getContainerInfo().getAnchorRow(), d.getContainerInfo().getAnchorColumn()));
    });
  });
}

/** 名前付き範囲と保護は、ファイル全体で1回だけ取得する */
function fileCache_(ss) {
  const c = { named: {}, protSheet: {}, protRange: {}, err: [] };
  try {
    ss.getNamedRanges().forEach((nr) => {
      try {
        const r = nr.getRange();
        const n = r.getSheet().getName();
        (c.named[n] = c.named[n] || []).push([nr.getName(), r.getLastRow(), r.getLastColumn()]);
      } catch (e) {
        // 参照先が壊れている名前付き範囲は数えない
      }
    });
  } catch (e) {
    c.err.push('名前付き範囲');
  }
  try {
    ss.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach((p) => {
      const n = p.getRange().getSheet().getName();
      c.protSheet[n] = (c.protSheet[n] || 0) + 1;
    });
  } catch (e) {
    c.err.push('シートの保護');
  }
  try {
    ss.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach((p) => {
      const r = p.getRange();
      const n = r.getSheet().getName();
      (c.protRange[n] = c.protRange[n] || []).push([r.getLastRow(), r.getLastColumn()]);
    });
  } catch (e) {
    c.err.push('範囲の保護');
  }
  return c;
}

/** 次の区切りを読む。失敗したら行数を減らして読み直し、それでもだめなら記録して先へ進む */
function scanNextChunk_(sheet, acc, cache) {
  const maxRows = sheet.getMaxRows();
  if (maxRows !== acc.maxRows) {
    acc.notes.push('診断中に行数が ' + acc.maxRows + ' → ' + maxRows + ' に変わった');
    acc.maxRows = Math.min(acc.maxRows, maxRows);
    if (acc.nextRow > acc.maxRows) return;
  }
  const r0 = acc.nextRow;
  const rows = Math.min(acc.rowsPerChunk, acc.maxRows - r0 + 1);
  const range = sheet.getRange(r0, 1, rows, acc.maxCols);
  let formulas = null;
  let values = null;
  let dvs = null;
  try {
    formulas = range.getFormulas();
    if (acc.readValues) values = range.getValues();
    if (LT.SCAN_VALIDATIONS) dvs = range.getDataValidations();
  } catch (e) {
    acc.retries++;
    if (rows > LT.MIN_CHUNK_ROWS && acc.retries <= LT.MAX_RETRIES) {
      acc.rowsPerChunk = Math.max(LT.MIN_CHUNK_ROWS, Math.floor(rows / 2));
      acc.notes.push(r0 + '行目から読み取りに失敗 → ' + acc.rowsPerChunk + ' 行ずつにして読み直す');
      return;
    }
    if (!formulas) {
      acc.unscanned.push([r0, r0 + rows - 1]);
      acc.notes.push(r0 + '〜' + (r0 + rows - 1) + '行目を読めなかった: ' + errText_(e).slice(0, 80));
      acc.nextRow = r0 + rows;
      acc.retries = 0;
      return;
    }
    if (acc.readValues && !values) acc.noValues.push([r0, r0 + rows - 1]);
    else if (LT.SCAN_VALIDATIONS && !dvs) acc.notes.push(r0 + '〜' + (r0 + rows - 1) + '行目の入力規則を読めなかった');
  }
  processChunk_(acc, r0, formulas, values, dvs, cache);
  acc.nextRow = r0 + rows;
  acc.retries = 0;
}

/** 読み取った区切りを数える(値の配列だけを使うので、ローカルでも試験できる) */
function processChunk_(acc, r0, formulas, values, dvs, cache) {
  for (let i = 0; i < formulas.length; i++) {
    const row = r0 + i;
    const fr = formulas[i];
    const vr = values ? values[i] : null;
    const dr = dvs ? dvs[i] : null;
    let rowF = '';
    let rowV = '';
    for (let j = 0; j < fr.length; j++) {
      const col = j + 1;
      const f = fr[j];
      if (f) {
        acc.formulas++;
        bump_(acc.lastF, row, col);
        rowF += j + '\u0002' + f + '\u0003';
        countFormula_(acc, analyzeCached_(f, cache), f, row, col);
      }
      if (vr) {
        const v = vr[j];
        const key = cellKey_(v);
        if (key !== '') {
          bump_(acc.lastV, row, col);
          if (!f) {
            acc.constants++;
            bump_(acc.lastC, row, col);
          }
          if (typeof v === 'number') acc.numSum += v;
          else if (typeof v === 'string') {
            if (isErrorValue_(v)) acc.errors++;
            else if (isLoading_(v)) acc.loading++;
          }
        }
        rowV += key + '\u0001';
      }
      if (dr && dr[j]) {
        acc.dv++;
        bump_(acc.lastDv, row, col);
      }
    }
    const rh = rowF ? hash53_(rowF) : '0';
    if (acc.rowFp) acc.rowFp.push(rh);
    acc.fpF = hash53_(acc.fpF + '|' + rh);
    if (vr) acc.fpV = hash53_(acc.fpV + '|' + hash53_(rowV));
  }
}

function countFormula_(acc, a, f, row, col) {
  for (const k in a.fn) acc.fn[k] = (acc.fn[k] || 0) + 1;
  for (const k in a.colRefs) acc.colRefs[k] = (acc.colRefs[k] || 0) + a.colRefs[k];
  for (const k in a.openRefs) acc.openRefs[k] = (acc.openRefs[k] || 0) + a.openRefs[k];
  for (const k in a.sheets) if (k !== acc.name) acc.sheetRefs[k] = (acc.sheetRefs[k] || 0) + 1;
  if (a.cond) acc.cat.cond++;
  if (a.array) acc.cat.array++;
  if (a.text) acc.cat.text++;
  if (a.volatile.length) {
    acc.cat.volatile++;
    addEx_(acc, '揮発性関数', a.volatile.join(','), a1_(row, col), f, '', '');
  }
  if (a.dynamic.length) {
    acc.cat.dynamic++;
    addEx_(acc, 'INDIRECT・OFFSET(参照先は未確認)', a.dynamic.join(','), a1_(row, col), f, '', '');
  }
  if (a.external.length) {
    acc.cat.external++;
    if (a.imports.length) a.imports.forEach((im) => addEx_(acc, '外部取得', 'IMPORTRANGE', a1_(row, col), f, im.id, im.range));
    else addEx_(acc, '外部取得', a.external.join(','), a1_(row, col), f, '', '');
  }
}

function addEx_(acc, type, fn, cell, formula, src, range) {
  acc.exCount[type] = (acc.exCount[type] || 0) + 1;
  if (acc.exCount[type] > LT.MAX_EXAMPLES) return;
  acc.ex.push([type, fn, cell, String(formula).slice(0, LT.SNIPPET), src, range]);
}

function analyzeCached_(f, cache) {
  let a = cache.get(f);
  if (a) return a;
  a = analyzeFormula_(f);
  if (cache.size > 20000) cache.clear();
  cache.set(f, a);
  return a;
}

// ------------------------------------------------------------
// 数式の解析(スプレッドシートを使わない関数。ローカルで試験している)
// ------------------------------------------------------------

// 列全体参照(例: '編明細'!$R:$R、編明細!R:R、$A:$C)。1: シート名と「!」 2,4: 「$」 3,5: 列の英字
const COLUMN_REF_RE =
  /('(?:[^']|'')+'!|[^\s'!(),;:&=+\-*\/<>^"{}\[\]\u0001]+!)?(?<![A-Za-z0-9_$.])(\$?)([A-Z]{1,3}):(\$?)([A-Z]{1,3})(?![A-Za-z0-9_(.])/g;
// 終わりの行がない範囲(例: '編明細'!$A$2:$A)。列全体参照と同じくらい広い
const OPEN_REF_RE =
  /('(?:[^']|'')+'!|[^\s'!(),;:&=+\-*\/<>^"{}\[\]\u0001]+!)?(?<![A-Za-z0-9_$.])(\$?)([A-Z]{1,3})(\$?)(\d+):(\$?)([A-Z]{1,3})(?![A-Za-z0-9_(.$])/g;
const SHEET_REF_RE = /('(?:[^']|'')+'|[^\s'!(),;:&=+\-*\/<>^"{}\[\]\u0001]+)!/g;
const FUNC_RE = /(?<![A-Za-z0-9_.$\u0004])([A-Z][A-Z0-9_.]*)\s*\(/g;

function analyzeFormula_(f) {
  const masked = maskStrings_(f);
  const noSheet = masked.replace(/'(?:[^']|'')*'/g, (s) => "'" + '\u0004'.repeat(Math.max(0, s.length - 2)) + "'");
  const fn = {};
  const calls = [];
  const fre = new RegExp(FUNC_RE.source, 'g');
  let m;
  while ((m = fre.exec(noSheet)) !== null) {
    fn[m[1]] = 1;
    calls.push([m[1], m.index + m[0].length - 1]);
  }
  const colRefs = {};
  findColumnRefs_(masked).forEach((r) => (colRefs[r.text] = (colRefs[r.text] || 0) + 1));
  const openRefs = {};
  findOpenRefs_(masked).forEach((r) => (openRefs[r.text] = (openRefs[r.text] || 0) + 1));
  const sheets = {};
  findSheetRefs_(masked).forEach((n) => (sheets[n] = 1));
  const wide = Object.keys(colRefs).length + Object.keys(openRefs).length > 0;
  const has = (name) => fn[name] === 1;
  let text = false;
  if (wide) {
    for (const call of calls) {
      if (TEXT_FUNCS.indexOf(call[0]) < 0) continue;
      const top = topLevelArgs_(noSheet, call[1]);
      if (findColumnRefs_(top).length || findOpenRefs_(top).length) {
        text = true;
        break;
      }
    }
  }
  return {
    fn: fn,
    colRefs: colRefs,
    openRefs: openRefs,
    sheets: sheets,
    cond: wide && COND_FUNCS.some(has),
    array: wide && ARRAY_FUNCS.some(has),
    text: text,
    volatile: VOLATILE_FUNCS.filter(has),
    dynamic: DYNAMIC_FUNCS.filter(has),
    external: EXTERNAL_FUNCS.filter(has),
    imports: has('IMPORTRANGE') ? parseImportRange_(f) : [],
  };
}

/** 数式の中の "文字列" を、同じ長さの記号に置き換える(文字列の中は参照として扱わない) */
function maskStrings_(formula) {
  return formula.replace(/"(?:[^"]|"")*"/g, (s) => '"' + '\u0001'.repeat(s.length - 2) + '"');
}

function findColumnRefs_(masked) {
  const re = new RegExp(COLUMN_REF_RE.source, 'g');
  const out = [];
  let m;
  while ((m = re.exec(masked)) !== null) {
    const prefix = m[1] || '';
    out.push({ text: m[0], prefix: prefix, sheetName: prefix ? unquoteSheetName_(prefix.slice(0, -1)) : '', c1: m[3], c2: m[5] });
  }
  return out;
}

function findOpenRefs_(masked) {
  const re = new RegExp(OPEN_REF_RE.source, 'g');
  const out = [];
  let m;
  while ((m = re.exec(masked)) !== null) {
    const prefix = m[1] || '';
    out.push({ text: m[0], prefix: prefix, sheetName: prefix ? unquoteSheetName_(prefix.slice(0, -1)) : '', c1: m[3], row1: Number(m[5]), c2: m[7] });
  }
  return out;
}

function findSheetRefs_(masked) {
  const re = new RegExp(SHEET_REF_RE.source, 'g');
  const out = [];
  let m;
  while ((m = re.exec(masked)) !== null) out.push(unquoteSheetName_(m[1]));
  return out;
}

function unquoteSheetName_(name) {
  return name.charAt(0) === "'" ? name.slice(1, -1).replace(/''/g, "'") : name;
}

/** 関数の「(」から対応する「)」までのうち、入れ子になっていない部分だけを返す(入れ子の中は空白にする) */
function topLevelArgs_(s, open) {
  let depth = 0;
  let out = '';
  for (let i = open; i < s.length; i++) {
    const ch = s.charAt(i);
    if (ch === "'") {
      let j = i + 1;
      while (j < s.length) {
        if (s.charAt(j) === "'") {
          if (s.charAt(j + 1) === "'") {
            j += 2;
            continue;
          }
          break;
        }
        j++;
      }
      out += depth === 1 ? s.slice(i, j + 1) : ' '.repeat(j + 1 - i);
      i = j;
      continue;
    }
    if (ch === '(') {
      depth++;
      out += ' ';
      continue;
    }
    if (ch === ')') {
      depth--;
      if (depth === 0) break;
      out += ' ';
      continue;
    }
    out += depth === 1 ? ch : ' ';
  }
  return out;
}

function parseImportRange_(f) {
  const out = [];
  const re = /IMPORTRANGE\s*\(\s*("(?:[^"]|"")*"|[^,;()]+)\s*[,;]\s*("(?:[^"]|"")*"|[^,;()]+)/gi;
  const lit = (s) => (s.charAt(0) === '"' ? s.slice(1, -1).replace(/""/g, '"') : null);
  let m;
  while ((m = re.exec(f)) !== null) {
    const key = lit(m[1].trim());
    const rng = lit(m[2].trim());
    let id;
    if (key === null) id = '(セル参照など・未確認: ' + m[1].trim().slice(0, 40) + ')';
    else {
      const mm = key.match(/\/d\/([-\w]{20,})/);
      id = mm ? mm[1] : key.slice(0, 80);
    }
    out.push({ id: id, range: rng === null ? '(式・未確認: ' + m[2].trim().slice(0, 40) + ')' : rng });
  }
  return out;
}

// ------------------------------------------------------------
// 1シート分の結果をまとめる
// ------------------------------------------------------------

function newBuffer_() {
  return { count: 0, keys: [], detail: [], fn: [], colref: [], refs: [], ext: [], rowfp: [] };
}

function bufferSheet_(buf, acc) {
  buf.count++;
  buf.keys.push(String(acc.sheetId));
  buf.detail.push(detailRow_(acc));
  const others = Object.keys(acc.fn)
    .filter((k) => FN_COLS.indexOf(k) < 0)
    .sort((a, b) => acc.fn[b] - acc.fn[a])
    .slice(0, 15)
    .map((k) => k + ':' + acc.fn[k])
    .join(' ');
  buf.fn.push([acc.sheetId, acc.name, acc.formulas].concat(FN_COLS.map((k) => acc.fn[k] || 0)).concat([others]));
  Object.keys(acc.colRefs).forEach((k) => {
    const r = findColumnRefs_(k)[0];
    if (!r) return;
    buf.colref.push([acc.sheetId, acc.name, r.sheetName || acc.name, k, '列全体', acc.colRefs[k], colIndex_(r.c2) - colIndex_(r.c1) + 1, '', '']);
  });
  Object.keys(acc.openRefs).forEach((k) => {
    const r = findOpenRefs_(k)[0];
    if (!r) return;
    buf.colref.push([acc.sheetId, acc.name, r.sheetName || acc.name, k, 'A2:A 形式(' + r.row1 + '行目から)', acc.openRefs[k], colIndex_(r.c2) - colIndex_(r.c1) + 1, '', '']);
  });
  Object.keys(acc.sheetRefs).forEach((k) => buf.refs.push([acc.sheetId, acc.name, k, '数式', acc.sheetRefs[k], '']));
  countList_(acc.struct.chartSources || []).forEach((x) => x[0] !== acc.name && buf.refs.push([acc.sheetId, acc.name, x[0], 'グラフ', x[1], '']));
  countList_(acc.struct.pivotSources || []).forEach((x) => x[0] !== acc.name && buf.refs.push([acc.sheetId, acc.name, x[0], 'ピボット', x[1], '']));
  acc.ex.forEach((e) => buf.ext.push([acc.sheetId, acc.name].concat(e)));
  Object.keys(acc.exCount).forEach((t) => {
    if (acc.exCount[t] > LT.MAX_EXAMPLES) buf.ext.push([acc.sheetId, acc.name, t, '', '', '…ほか ' + (acc.exCount[t] - LT.MAX_EXAMPLES) + ' 件(例は最大 ' + LT.MAX_EXAMPLES + ' 件)', '', '']);
  });
  buf.rowfp.push([acc.sheetId, acc.name, acc.maxRows, acc.rowFp ? acc.rowFp.join(',') : '']);
}

function detailRow_(acc) {
  const s = acc.struct;
  const g = (k, i) => (s[k] ? s[k][i] : 0);
  const keepR = Math.max(acc.lastF[0], acc.lastC[0], acc.lastV[0], acc.lastDv[0], g('cfLast', 0), g('mergeLast', 0), g('protLast', 0),
    g('namedLast', 0), g('filterLast', 0), g('anchorLast', 0), acc.frozenRows);
  const keepC = Math.max(acc.lastF[1], acc.lastC[1], acc.lastV[1], acc.lastDv[1], g('cfLast', 1), g('mergeLast', 1), g('protLast', 1),
    g('namedLast', 1), g('filterLast', 1), g('anchorLast', 1), acc.frozenCols);
  const cells = acc.maxRows * acc.maxCols;
  const sumOf = (o) => Object.keys(o).reduce((t, k) => t + o[k], 0);
  const noValue = !acc.readValues;
  const v = {
    sheetId: acc.sheetId,
    name: acc.name,
    hidden: acc.hidden ? '非表示' : '表示',
    maxRows: acc.maxRows,
    maxCols: acc.maxCols,
    cells: cells,
    formulas: acc.formulas,
    constants: noValue ? '未取得' : acc.constants,
    lastFR: acc.lastF[0],
    lastFC: acc.lastF[1],
    lastCR: noValue ? '未取得' : acc.lastC[0],
    lastCC: noValue ? '未取得' : acc.lastC[1],
    lastVR: noValue ? '未取得' : acc.lastV[0],
    lastVC: noValue ? '未取得' : acc.lastV[1],
    dv: LT.SCAN_VALIDATIONS ? acc.dv : '未取得',
    lastDvR: acc.lastDv[0],
    lastDvC: acc.lastDv[1],
    cf: orBlank_(s.cf),
    cfCustom: orBlank_(s.cfCustom),
    lastCfR: g('cfLast', 0),
    lastCfC: g('cfLast', 1),
    merge: orBlank_(s.merge),
    lastMR: g('mergeLast', 0),
    lastMC: g('mergeLast', 1),
    protSheet: s.protSheet,
    protRange: s.protRange,
    lastPR: g('protLast', 0),
    lastPC: g('protLast', 1),
    named: s.named,
    lastNR: g('namedLast', 0),
    lastNC: g('namedLast', 1),
    charts: orBlank_(s.charts),
    pivots: orBlank_(s.pivots),
    filter: s.filter || '',
    bandings: orBlank_(s.bandings),
    slicers: orBlank_(s.slicers),
    images: orBlank_(s.images),
    drawings: orBlank_(s.drawings),
    scripts: (s.scripts || []).join(', '),
    keepR: keepR,
    keepC: keepC,
    outside: acc.isGrid ? cells - Math.min(keepR, acc.maxRows) * Math.min(keepC, acc.maxCols) : 0,
    colRefs: sumOf(acc.colRefs),
    openRefs: sumOf(acc.openRefs),
    catCond: acc.cat.cond,
    catArray: acc.cat.array,
    catText: acc.cat.text,
    volatile: acc.cat.volatile,
    dynamic: acc.cat.dynamic,
    external: acc.cat.external,
    errors: noValue ? '未取得' : acc.errors,
    loading: noValue ? '未取得' : acc.loading,
    numSum: noValue ? '未取得' : Math.round(acc.numSum * 1e6) / 1e6,
    fpV: noValue ? '未取得' : acc.fpV,
    fpF: acc.fpF,
    refsTo: Object.keys(acc.sheetRefs).sort((a, b) => acc.sheetRefs[b] - acc.sheetRefs[a]).map((k) => k + '(' + acc.sheetRefs[k] + ')').join(', '),
    inCount: '',
    inFrom: '',
    unscanned: rangesText_(acc.unscanned),
    noValues: rangesText_(acc.noValues),
    notes: acc.notes.slice(0, 20).join(' / '),
    scannedAt: nowText_(),
    sec: Math.round(acc.ms / 100) / 10,
  };
  return H_DETAIL.map((h) => (v[h[0]] === undefined ? '' : v[h[0]]));
}

function flushBuffer_(buf) {
  if (!buf.count) return;
  const keys = new Set(buf.keys);
  upsertRows_(S.DETAIL, H_DETAIL.map((h) => h[1]), keys, buf.detail);
  upsertRows_(S.FN, H_FN, keys, buf.fn);
  upsertRows_(S.COLREF, H_COLREF, keys, buf.colref);
  upsertRows_(S.REFS, H_REFS, keys, buf.refs);
  upsertRows_(S.EXT, H_EXT, keys, buf.ext);
  upsertRows_(S.ROWFP, H_ROWFP, keys, buf.rowfp, true);
  const fresh = newBuffer_();
  Object.keys(fresh).forEach((k) => (buf[k] = fresh[k]));
}

// ------------------------------------------------------------
// まとめ(全シートを調べ終わったあと)
// ------------------------------------------------------------

function finalizeDiag_(st, target) {
  const detail = readRows_(S.DETAIL);
  const names = {};
  detail.forEach((r) => (names[r[DI.name]] = r));

  // 参照されている側の集計と、参照先があるかの確認
  const refs = readRows_(S.REFS);
  const incoming = {};
  refs.forEach((r) => {
    const to = String(r[2]);
    r[5] = names[to] ? 'ある' : 'ない(名前の違い・削除・INDIRECT の文字列など)';
    (incoming[to] = incoming[to] || {})[String(r[1])] = true;
  });
  rewriteRows_(S.REFS, H_REFS, refs);
  detail.forEach((r) => {
    const from = Object.keys(incoming[String(r[DI.name])] || {});
    r[DI.inCount] = from.length;
    r[DI.inFrom] = from.slice(0, 10).join(', ') + (from.length > 10 ? ' …ほか' + (from.length - 10) : '');
  });
  rewriteRows_(S.DETAIL, H_DETAIL.map((h) => h[1]), detail);

  // 列全体参照: 参照先の確保行数と、走査の目安(推定)
  const colref = readRows_(S.COLREF);
  const scanByTarget = {};
  colref.forEach((r) => {
    const t = names[String(r[2])];
    if (!t) {
      r[7] = '不明';
      r[8] = '不明';
      return;
    }
    const m = String(r[4]).match(/(\d+)行目から/);
    const rows = Number(t[DI.maxRows]) - (m ? Number(m[1]) - 1 : 0);
    r[7] = Number(t[DI.maxRows]);
    r[8] = Number(r[5]) * Math.max(0, rows) * (Number(r[6]) || 1);
    scanByTarget[r[2]] = (scanByTarget[r[2]] || 0) + r[8];
  });
  rewriteRows_(S.COLREF, H_COLREF, colref);

  const struct = buildStructure_(detail);
  writeTable_(S.STRUCT, struct.header, struct.rows);
  writeTable_(S.SUMMARY, ['項目', '値', '根拠・注記'], buildSummary_(st, target, detail, scanByTarget, struct));
}

function buildStructure_(detail) {
  const fps = {};
  readRows_(S.ROWFP).forEach((r) => (fps[String(r[0])] = String(r[3] || '')));
  const repRow = detail.filter((r) => String(r[DI.name]) === LT.REPRESENTATIVE_SHEET)[0];
  const repFp = repRow ? fps[String(repRow[DI.sheetId])] || '' : '';
  const rep = repFp ? repFp.split(',') : null;
  const groups = {};
  let gNo = 0;
  const counts = { same: 0, part: 0, other: 0, none: 0, na: 0 };
  const rows = detail.map((r) => {
    const formulas = Number(r[DI.formulas]) || 0;
    let group = '';
    if (formulas > 0) {
      const key = String(r[DI.fpF]);
      if (!groups[key]) groups[key] = 'G' + ++gNo;
      group = groups[key];
    }
    const fp = fps[String(r[DI.sheetId])];
    let same = '';
    let diffN = '';
    let diffText = '';
    let judge;
    if (formulas === 0) {
      judge = '数式なし';
      counts.none++;
    } else if (!rep) {
      judge = '比べられない(代表「' + LT.REPRESENTATIVE_SHEET + '」がないか、行の指紋がない)';
      counts.na++;
    } else if (!fp) {
      judge = '比べられない(行が多いので行の指紋を残していない)';
      counts.na++;
    } else {
      const a = fp.split(',');
      const n = Math.max(a.length, rep.length);
      const diff = [];
      let repRows = 0;
      let match = 0;
      for (let i = 0; i < n; i++) {
        if (rep[i] && rep[i] !== '0') repRows++;
        if (a[i] === rep[i]) {
          if (rep[i] && rep[i] !== '0') match++;
        } else diff.push(i + 1);
      }
      same = match + ' / ' + repRows;
      diffN = diff.length;
      diffText = rowsText_(diff);
      if (diff.length === 0) {
        judge = '代表と同じ';
        counts.same++;
      } else if (repRows && match / repRows >= 0.9) {
        judge = '一部だけ違う';
        counts.part++;
      } else {
        judge = '別の作り';
        counts.other++;
      }
    }
    return [r[DI.name], r[DI.hidden], r[DI.maxRows] + '×' + r[DI.maxCols], formulas, group, same, diffN, diffText, judge];
  });
  return {
    header: ['シート名', '表示', '確保 行×列', '数式セル', '同じ数式のグループ', '代表と同じ数式の行', '違う行の数', '違う行(例)', '判定'],
    rows: rows,
    counts: counts,
    groups: gNo,
  };
}

function buildSummary_(st, target, detail, scanByTarget, struct) {
  const n = (r, k) => (typeof r[DI[k]] === 'number' ? r[DI[k]] : 0);
  const sum = (k, filter) => detail.filter(filter || (() => true)).reduce((t, r) => t + n(r, k), 0);
  const hidden = (r) => r[DI.hidden] === '非表示';
  const three = (r) => /^\d{3}$/.test(String(r[DI.name]));
  const cells = sum('cells');
  const formulas = sum('formulas');
  const unscanned = detail.filter((r) => r[DI.unscanned]).length;
  const noValues = detail.filter((r) => r[DI.noValues]).length;
  const scripts = detail.filter((r) => r[DI.scripts]).map((r) => r[DI.name] + ': ' + r[DI.scripts]);
  const recalc = safe_(() => recalcText_(String(target.ss.getRecalculationInterval())));
  const targets = Object.keys(scanByTarget).sort((a, b) => scanByTarget[b] - scanByTarget[a]);
  const scanTotal = targets.reduce((t, k) => t + scanByTarget[k], 0);

  const rows = [
    ['■ 対象', '', ''],
    ['ファイル', st.targetName, '試験用コピー(ID の末尾 …' + st.targetId.slice(-6) + ')'],
    ['調査の期間', fmtTime_(st.startedAt) + ' 〜 ' + nowText_(), '実行 ' + st.runs + ' 回 / 読み取り ' + st.chunks + ' 回'],
    ['調べたシート', detail.length + ' / ' + st.queue.length + ' 枚', '見つからなかったシートは「' + S.LOG + '」と進み具合のエラーを参照'],
    ['調べられなかった行があるシート', unscanned + ' 枚', unscanned ? '「' + S.DETAIL + '」の「調べられなかった行」列' : ''],
    ['値を読めなかった行があるシート', noValues + ' 枚', noValues ? '「' + S.DETAIL + '」の「値を読めなかった行」列' : ''],
    ['', '', ''],
    ['■ 大きさ(実測)', '', ''],
    ['確保セル数の合計', cells, '上限 1,000万の ' + pct_(cells, LT.CELL_LIMIT)],
    ['うち非表示のシート', sum('cells', hidden), detail.filter(hidden).length + ' 枚'],
    ['うち名前が3桁の数字のシート', sum('cells', three), detail.filter(three).length + ' 枚'],
    ['内容と構造物の外側のセル(参考)', sum('outside'), '入力規則・条件付き書式・結合・保護・名前付き範囲・図形も含めた外側。削除してよいという意味ではない'],
    ['', '', ''],
    ['■ 数式(実測)', '', ''],
    ['数式セルの合計', formulas, ''],
    ['うち非表示のシート', sum('formulas', hidden), '非表示のシートの数式も計算は止まらない'],
    ['うち名前が3桁の数字のシート', sum('formulas', three), ''],
    ['列全体参照(個)', sum('colRefs'), '例: \'編明細\'!$R:$R'],
    ['A2:A 形式の参照(個)', sum('openRefs'), '終わりの行がない範囲'],
    ['条件集計(SUMIFS など)×列全体(セル)', sum('catCond'), ''],
    ['配列計算(FILTER など)×列全体(セル)', sum('catArray'), ''],
    ['列全体を1行ずつ加工(SUBSTITUTE・SEARCH など)(セル)', sum('catText'), '参照先の全行を毎回加工する式'],
    ['揮発性関数(TODAY・NOW など)(セル)', sum('volatile'), '再計算の設定: ' + recalc],
    ['INDIRECT・OFFSET(セル)', sum('dynamic'), '参照先は文字列で決まるため、参照関係は「未確認」'],
    ['外部取得関数(IMPORTRANGE など)(セル)', sum('external'), '「' + S.EXT + '」に例'],
    ['エラー値(セル)', sum('errors'), '#REF!・#N/A など。コピーでは IMPORTRANGE の許可待ちでも出る'],
    ['', '', ''],
    ['■ 構造(行ごとの数式を代表「' + LT.REPRESENTATIVE_SHEET + '」と比べた)', '', ''],
    ['代表と同じ', struct.counts.same + ' 枚', ''],
    ['一部だけ違う', struct.counts.part + ' 枚', '代表の数式の行のうち 9割以上が同じ'],
    ['別の作り', struct.counts.other + ' 枚', ''],
    ['数式なし / 比べられない', struct.counts.none + ' / ' + struct.counts.na + ' 枚', ''],
    ['数式が完全に同じグループの数', struct.groups, '「' + S.STRUCT + '」の「同じ数式のグループ」'],
    ['', '', ''],
    ['■ 列全体参照の走査量の目安(推定)', scanTotal, '個数 × 参照先の確保行数 × 列数。計算時間の実測ではない'],
  ];
  targets.slice(0, 10).forEach((t) => rows.push(['参照先「' + t + '」', scanByTarget[t], pct_(scanByTarget[t], scanTotal) + '(目安の中での割合)']));

  rows.push(['', '', ''], ['■ ボトルネックの候補(数は実測、効き目は未測定)', '', '']);
  const cands = [];
  if (scanTotal > 0) cands.push(['列全体参照が多い', '目安 ' + fmtNum_(scanTotal) + ' セル分', '最も大きい参照先: 「' + targets[0] + '」']);
  if (sum('catText') > 0) cands.push(['列全体を1行ずつ加工する式', sum('catText') + ' セル', '年・月を正規化した列を用意すると減らせる見込み(優先順位4)']);
  if (sum('catArray') > 0) cands.push(['配列計算で列全体を使う式', sum('catArray') + ' セル', 'FILTER などは条件の数だけ全行を調べる']);
  if (sum('formulas', hidden) > 0) cands.push(['非表示のシートの数式', sum('formulas', hidden) + ' セル', '非表示でも計算は止まらない。使っているかは人の判断が必要']);
  if (cells / LT.CELL_LIMIT > 0.7) cands.push(['セル数が上限に近い', pct_(cells, LT.CELL_LIMIT), '']);
  if (sum('volatile') > 0) cands.push(['揮発性関数', sum('volatile') + ' セル', '再計算の設定: ' + recalc]);
  if (sum('external') > 0) cands.push(['外部取得関数', sum('external') + ' セル', '取得先が重いと待ちが出る']);
  if (!cands.length) cands.push(['(目立つ候補なし)', '', '']);
  cands.forEach((c) => rows.push(c));
  rows.push(['', '', '効き目は、優先順位3・4の試験で変更前後を測って確かめる']);

  rows.push(['', '', ''], ['■ 調べていないこと・未確認', '', '']);
  [
    ['ほかのファイルからこのファイルへの参照(IMPORTRANGE など)', '未確認', 'このファイルの中からは見えない。取り込み側のファイルで確認が必要'],
    ['既存の GAS(Apps Script)とトリガー', scripts.length ? '手がかりあり: ' + scripts.slice(0, 5).join(' / ') : '未確認', '本番の 拡張機能 → Apps Script で確認する'],
    ['INDIRECT・OFFSET の参照先', '未確認', '文字列で決まるため'],
    ['非表示の行・列、グループ化', '未調査', ''],
    ['書式(背景色・罫線)が付いている範囲', '未調査', '「外側のセル」に含めていない'],
    ['フィルタ表示、入力規則の参照先', '未調査', ''],
    ['配列数式が広がる範囲', '値で近似', '広がった結果は「値がある最終行」に含まれる'],
  ].forEach((x) => rows.push(x));
  return rows;
}

// ============================================================
// 8〜10. 比較基準
// ============================================================

const H_TARGETS = ['使う', 'シート名', '範囲(空欄 = 使われている範囲全体)', '何を確かめるか', '候補の理由'];
const H_BASE = ['基準の名前', '番号', '記録日時', 'シート名', '範囲', '行', '列', '数値の合計', '空白でないセル', 'エラー値', '読み込み中', '値の指紋', '表示の指紋', '本番との違い(セル)', '本番との違い(例)', 'メモ'];
const H_BASEDATA = ['番号', '行', '値', '表示'];
const H_RESULT = ['番号', 'シート名', 'セル', '基準の表示', '今の表示', '基準の値', '今の値'];

function lt_baselineCandidates() {
  withLock_('8. 比較する場所の候補', () => {
    const sh = ensureSheet_(S.TARGETS, H_TARGETS);
    if (sh.getLastRow() > 1) {
      activate_(sh);
      return notify_('8. 比較する場所の候補', ['シート「' + S.TARGETS + '」にはもう行があります。直接書き直してください。', '(2行目より下を消してからもう一度実行すると、候補を作り直します)']);
    }
    const f = openRole_('BACKUP'); // 比較基準は保管用(変更前の状態)から取る
    const names = f.ss.getSheets().filter((s) => s.getType() === SpreadsheetApp.SheetType.GRID).map((s) => s.getName());
    const rows = [];
    if (names.indexOf(LT.REPRESENTATIVE_SHEET) >= 0) rows.push([true, LT.REPRESENTATIVE_SHEET, '', '代表の人別シート(金額・時間・表示)', '代表シート']);
    names.forEach((n) => {
      const m = n.match(LT.BASELINE_NAME_HINT);
      if (m && n !== LT.REPRESENTATIVE_SHEET) rows.push([true, n, '', '', '名前に「' + m[1] + '」を含む']);
    });
    if (rows.length) {
      writeRows_(sh, 2, rows, H_TARGETS.length);
      try {
        sh.getRange(2, 1, rows.length, 1).insertCheckboxes();
      } catch (e) {
        // チェックボックスにできなくても、TRUE/FALSE で動く
      }
    }
    log_('8. 比較する場所の候補', 'メニュー', rows.length + ' 件', '');
    activate_(sh);
    notify_('8. 比較する場所の候補', [
      '候補を ' + rows.length + ' 件書きました(シート「' + S.TARGETS + '」)。',
      '',
      '・特に間違ってはいけない歩合・給与・店舗別合計・支給額の場所が入っているか確かめてください。',
      '・いらない行はチェックを外し、足りない場所は行を足してください(範囲は A1:H50 のように書きます)。',
      '・できたら「9. 比較基準を記録する」を実行します。',
    ]);
  });
}

function lt_baselineRecord() {
  const title = '9. 比較基準を記録';
  withLock_(title, () => {
    const props = PropertiesService.getScriptProperties();
    // 基準は保管用(変更前の状態)から取る。試験用コピーは軽量化の試験で変えるので、基準には使わない
    const src = openRole_('BACKUP');
    const prod = openRole_('PROD');
    const test = openRole_('TEST');
    let job = JSON.parse(props.getProperty('LT_BASELINE_JOB') || 'null');
    if (job && job.srcId !== src.id) {
      props.deleteProperty('LT_BASELINE_JOB');
      job = null;
    }
    if (!job) {
      const items = readTargets_();
      if (!items.length) return notify_(title, ['シート「' + S.TARGETS + '」に、チェックの入った行がありません。先に「8.」を実行してください。']);
      const resolved = [];
      const problems = [];
      items.forEach((it) => {
        const sh = src.ss.getSheetByName(it.sheet);
        if (!sh) return problems.push('保管用に「' + it.sheet + '」がない');
        if (!prod.ss.getSheetByName(it.sheet)) problems.push('本番に「' + it.sheet + '」がない(本番との比較は「未比較」になる)');
        if (!test.ss.getSheetByName(it.sheet)) problems.push('試験用コピーに「' + it.sheet + '」がない(「10.」で比べられない)');
        try {
          const rg = it.a1 ? sh.getRange(it.a1) : sh.getRange(1, 1, Math.max(1, sh.getLastRow()), Math.max(1, sh.getLastColumn()));
          resolved.push({ sheet: it.sheet, a1: rg.getA1Notation(), row: rg.getRow(), col: rg.getColumn(), rows: rg.getNumRows(), cols: rg.getNumColumns(), memo: it.memo });
        } catch (e) {
          problems.push('「' + it.sheet + '」の範囲「' + it.a1 + '」を読めない');
        }
      });
      if (!resolved.length) return notify_(title, ['記録できる範囲がありません。'].concat(problems));
      const cells = resolved.reduce((t, r) => t + r.rows * r.cols, 0);
      const ok = confirm_(title, [
        '基準を取るファイル: ' + src.name + '(保管用バックアップ=変更前の状態)',
        '比べるファイル: ' + prod.name + '(本番。コピーのあとに本番が変わったかを見る)',
        '対象: ' + resolved.length + ' か所 / ' + fmtNum_(cells) + ' セル',
      ].concat(resolved.slice(0, 15).map((r) => '・' + r.sheet + '!' + r.a1)).concat(resolved.length > 15 ? ['…ほか ' + (resolved.length - 15) + ' か所'] : []).concat([
        '',
        '変更内容: なし(読むだけ)。保管用の値をこのファイルに記録し、本番と同じかも比べます。',
        '復旧方法: 不要(対象ファイルは変更しません)',
      ]).concat(problems.length ? ['', '注意:'].concat(problems) : []).concat(['', '記録しますか?']));
      if (!ok) return;
      job = { setId: Utilities.formatDate(new Date(), tz_(), 'yyyyMMdd_HHmmss'), items: resolved, next: 0, srcId: src.id };
      props.setProperty('LT_BASELINE_JOB', JSON.stringify(job));
    }
    const t0 = Date.now();
    while (job.next < job.items.length && Date.now() - t0 < LT.TIME_LIMIT_MS - 30000) {
      recordBaselineItem_(job, job.next, src, prod);
      job.next++;
      props.setProperty('LT_BASELINE_JOB', JSON.stringify(job));
    }
    const done = job.next >= job.items.length;
    if (done) {
      props.deleteProperty('LT_BASELINE_JOB');
      props.setProperty('LT_BASELINE_LAST', job.setId);
    }
    log_(title, 'メニュー', done ? '完了' : '時間で区切った', job.setId + ' ' + job.next + ' / ' + job.items.length);
    const base = SpreadsheetApp.getActive().getSheetByName(S.BASE);
    activate_(base);
    const rows = readRows_(S.BASE).filter((r) => String(r[0]) === job.setId);
    const diffs = rows.filter((r) => typeof r[13] === 'number' && r[13] > 0);
    notify_(title, [
      done ? '記録しました。基準の名前: ' + job.setId : '時間で区切りました。もう一度「9.」を押すと続きを記録します(' + job.next + ' / ' + job.items.length + ')。',
      '',
      '結果はシート「' + S.BASE + '」です。',
      diffs.length ? '本番と保管用で値が違う場所が ' + diffs.length + ' か所あります。コピーのあとに本番を編集したか、IMPORTRANGE などの許可待ちが考えられます。' : '記録した範囲では、本番と保管用の値は同じでした。',
      done ? '次は「10. 試験用コピーを比較基準と比べる」です。' : '',
    ]);
  });
}

function recordBaselineItem_(job, idx, src, prod) {
  const it = job.items[idx];
  const tSh = src.ss.getSheetByName(it.sheet);
  const pSh = prod.ss.getSheetByName(it.sheet);
  const step = Math.max(1, Math.floor(LT.BASELINE_CHUNK_CELLS / it.cols));
  const data = [];
  let numSum = 0;
  let nonEmpty = 0;
  let errors = 0;
  let loading = 0;
  let fpV = '';
  let fpD = '';
  let prodDiff = 0;
  const prodEx = [];
  let note = pSh ? '' : '本番に同じ名前のシートがない';
  for (let r = 0; r < it.rows; r += step) {
    const n = Math.min(step, it.rows - r);
    const rg = tSh.getRange(it.row + r, it.col, n, it.cols);
    const vals = rg.getValues();
    const disp = rg.getDisplayValues();
    let pv = null;
    if (pSh && !note) {
      try {
        pv = pSh.getRange(it.row + r, it.col, n, it.cols).getValues();
      } catch (e) {
        note = '本番を読めなかった: ' + errText_(e).slice(0, 80);
      }
    }
    for (let i = 0; i < n; i++) {
      const keys = vals[i].map(cellKey_);
      keys.forEach((k, j) => {
        const v = vals[i][j];
        if (k !== '') {
          nonEmpty++;
          if (typeof v === 'number') numSum += v;
          else if (typeof v === 'string' && isErrorValue_(v)) errors++;
          else if (isLoading_(v)) loading++;
        }
        if (pv && !sameKey_(cellKey_(pv[i][j]), k)) {
          prodDiff++;
          if (prodEx.length < 5) prodEx.push(a1_(it.row + r + i, it.col + j) + ' 本番 ' + showKey_(cellKey_(pv[i][j])) + ' / コピー ' + showKey_(k));
        }
      });
      fpV = hash53_(fpV + '|' + keys.join('\u0001'));
      fpD = hash53_(fpD + '|' + disp[i].join('\u0001'));
      data.push([idx + 1, it.row + r + i, JSON.stringify(keys), JSON.stringify(disp[i])]);
    }
  }
  const no = idx + 1;
  replaceRows_(BASE_DATA_PREFIX + job.setId, H_BASEDATA, (row) => Number(row[0]) === no, data);
  replaceRows_(S.BASE, H_BASE, (row) => String(row[0]) === job.setId && Number(row[1]) === no, [[
    job.setId, no, nowText_(), it.sheet, it.a1, it.rows, it.cols, Math.round(numSum * 1e6) / 1e6, nonEmpty, errors, loading, fpV, fpD,
    note ? '未比較' : prodDiff, prodEx.join(' / '), note || it.memo || '',
  ]]);
}

function lt_baselineCompare() {
  const title = '10. 比較基準と比べる';
  withLock_(title, () => {
    const setId = PropertiesService.getScriptProperties().getProperty('LT_BASELINE_LAST');
    if (!setId) return notify_(title, ['まだ比較基準を記録していません。先に「9.」を実行してください。']);
    const test = openRole_('TEST');
    const items = readRows_(S.BASE).filter((r) => String(r[0]) === setId);
    const stored = {};
    readRows_(BASE_DATA_PREFIX + setId).forEach((r) => ((stored[Number(r[0])] = stored[Number(r[0])] || {})[Number(r[1])] = r));
    const summary = [];
    const details = [];
    let totalDiff = 0;
    const t0 = Date.now();
    for (const it of items) {
      if (Date.now() - t0 > LT.TIME_LIMIT_MS - 30000) {
        summary.push([it[1], it[3], it[4], '', '時間切れで未比較']);
        continue;
      }
      const no = Number(it[1]);
      const sh = test.ss.getSheetByName(String(it[3]));
      if (!sh) {
        summary.push([no, it[3], it[4], '', '試験用コピーにシートがない']);
        continue;
      }
      const rg = sh.getRange(String(it[4]));
      const step = Math.max(1, Math.floor(LT.BASELINE_CHUNK_CELLS / rg.getNumColumns()));
      let diff = 0;
      let missing = 0;
      for (let r = 0; r < rg.getNumRows(); r += step) {
        const n = Math.min(step, rg.getNumRows() - r);
        const part = sh.getRange(rg.getRow() + r, rg.getColumn(), n, rg.getNumColumns());
        const vals = part.getValues();
        const disp = part.getDisplayValues();
        for (let i = 0; i < n; i++) {
          const rowNo = rg.getRow() + r + i;
          const s = (stored[no] || {})[rowNo];
          if (!s) {
            missing++;
            continue;
          }
          const before = JSON.parse(String(s[2]));
          const beforeD = JSON.parse(String(s[3]));
          vals[i].forEach((v, j) => {
            const k = cellKey_(v);
            const b = before[j] === undefined ? '' : before[j];
            if (sameKey_(b, k)) return;
            diff++;
            if (details.length < LT.COMPARE_MAX_DETAILS) details.push([no, it[3], a1_(rowNo, rg.getColumn() + j), beforeD[j] === undefined ? '' : beforeD[j], disp[i][j], showKey_(b), showKey_(k)]);
          });
        }
      }
      totalDiff += diff;
      summary.push([no, it[3], it[4], diff, diff === 0 && missing === 0 ? '同じ' : diff ? '違いあり' : '基準にない行あり(' + missing + ')']);
    }
    const sh = ensureSheet_(S.RESULT, null);
    sh.clearContents();
    const head = [['比べた基準', setId, '比べた日時', nowText_(), '対象', test.name], ['', '', '', '', '', '']];
    writeRows_(sh, 1, head, 6);
    writeRows_(sh, 3, [['番号', 'シート名', '範囲', '違うセル', '判定']].concat(summary), 5);
    styleHeader_(sh, 3, 5);
    const start = summary.length + 5;
    writeRows_(sh, start, [H_RESULT].concat(details.length ? details : [['', '', '', '', '違いはありません', '', '']]), H_RESULT.length);
    styleHeader_(sh, start, H_RESULT.length);
    log_(title, 'メニュー', '違うセル ' + totalDiff, setId);
    activate_(sh);
    notify_(title + '(変更なし)', [
      '基準: ' + setId + '(保管用=変更前の状態から記録した値)',
      '比べたファイル: ' + test.name + '(試験用コピー)',
      '違うセルの合計: ' + totalDiff + (details.length >= LT.COMPARE_MAX_DETAILS ? '(詳細は最初の ' + LT.COMPARE_MAX_DETAILS + ' 件だけ)' : ''),
      '',
      '結果はシート「' + S.RESULT + '」です。',
      'TODAY() などを使うセルは、日付が変わると違いとして出ます。',
    ]);
  });
}

// ============================================================
// 11. 所要時間(試験用コピーだけ。数式を同じ値になる書き方に一時的に替え、毎回もとに戻す)
// ============================================================

const H_TIMING = ['日時', '回', 'シート', 'セル', '合計=書き込み〜読み終わり(ミリ秒)', '変更なしで読むだけ(ミリ秒)', '書き込みと反映の指示(ミリ秒)', '書き込み後の最初の読み取り(再計算待ちを含む)', '再計算・反映待ちの目安(合計−読むだけ)', '値は変更前と同じか', 'メモ'];

function lt_timing() {
  const title = '11. 所要時間を測る';
  const ids = { PROD: getId_('PROD'), BACKUP: getId_('BACKUP'), TEST: getId_('TEST') };
  if (!ids.TEST) return notify_(title, ['先に「1. 対象ファイルを登録する」を実行してください。']);
  if (ids.TEST === ids.PROD || ids.TEST === ids.BACKUP) return notify_(title, ['試験用コピーが本番か保管用と同じファイルです。この処理は実行しません。']);
  withLock_(title, () => {
    const cfg = LT.TIMING;
    const t = openRole_('TEST');
    const sheets = cfg.SHEETS.map((n) => t.ss.getSheetByName(n));
    const missing = cfg.SHEETS.filter((n, i) => !sheets[i]);
    if (missing.length) return notify_(title, ['試験用コピーに次のシートがありません: ' + missing.join('、'), '「変更後」と「変更前」の2枚を用意してから実行してください。']);
    const cells = sheets.map((sh) => sh.getRange(cfg.CELL));
    const originals = cells.map((c) => c.getFormula());
    if (originals.some((f) => !f)) return notify_(title, ['「' + cfg.CELL + '」が数式でないシートがあります。この処理は数式のセルだけを入れ直します。']);
    const ok = confirm_(title, [
      '対象ファイル: ' + t.name + '(試験用コピー。ID の末尾 …' + t.id.slice(-6) + ')',
      '対象: ' + cfg.SHEETS.map((n, i) => '「' + n + '」!' + cfg.CELL + '(今の数式 ' + originals[i] + ')').join(' / '),
      '件数: ' + cfg.SHEETS.length + ' セル × ' + cfg.ROUNDS + ' 回(1回につき 書き方を替える → 元に戻す の2回の再計算)',
      '変更内容: 数式を丸かっこで囲んだ同じ値の書き方に替えて再計算させ、' + cfg.READ_RANGE + ' の値を読み終わるまでの時間を測ります。毎回もとの数式に戻します。',
      '本番・保管用には書き込みません。',
      '復旧方法: 終わったら元の数式に戻ったかを確かめます。戻っていなければ、上の「今の数式」をそのセルに入れ直してください。',
      '',
      '始めますか?(数分かかることがあります。終わるまで試験用コピーを編集しないでください)',
    ]);
    if (!ok) return;
    // 6分の上限で強制終了されると finally が動かないので、先に元の数式を記録しておく(復旧用)
    log_(title, 'メニュー', '開始', '元の数式: ' + cfg.SHEETS.map((n, i) => n + '!' + cfg.CELL + ' ' + originals[i]).join(' / '));
    const t0 = Date.now();
    const rows = [];
    const before = sheets.map((sh) => sh.getRange(cfg.READ_RANGE).getValues());
    // 1回ごとに 読むだけ / 書き込みと反映の指示 / 書き込み後の最初の読み取り(再計算待ちを含む) を分けて測る
    const ms = cfg.SHEETS.map(() => ({ total: [], read: [], write: [], first: [], wait: [] }));
    let stopped = '';
    try {
      for (let r = 1; r <= cfg.ROUNDS && !stopped; r++) {
        const order = r % 2 ? cfg.SHEETS.map((_, i) => i) : cfg.SHEETS.map((_, i) => i).reverse(); // 順番の影響を減らす
        for (const i of order) {
          if (Date.now() - t0 > LT.TIME_LIMIT_MS - 60000) {
            stopped = '時間の都合で ' + r + ' 回目の途中で止めた';
            break;
          }
          const sh = sheets[i];
          [['=(' + originals[i].slice(1) + ')', '書き方を替えたとき'], [originals[i], '元の数式に戻したとき']].forEach(([f, memo]) => {
            let s = Date.now();
            sh.getRange(cfg.READ_RANGE).getValues();
            const read = Date.now() - s;
            s = Date.now();
            cells[i].setFormula(f);
            SpreadsheetApp.flush();
            const write = Date.now() - s;
            s = Date.now();
            const v = sh.getRange(cfg.READ_RANGE).getValues();
            const first = Date.now() - s;
            const total = write + first;
            const wait = Math.max(0, total - read);
            const m = ms[i];
            m.total.push(total);
            m.read.push(read);
            m.write.push(write);
            m.first.push(first);
            m.wait.push(wait);
            rows.push([nowText_(), r, cfg.SHEETS[i], cfg.CELL, total, read, write, first, wait, sameGrid_(before[i], v) ? '同じ' : '違う', memo]);
          });
        }
      }
    } finally {
      cells.forEach((c, i) => {
        if (c.getFormula() !== originals[i]) c.setFormula(originals[i]);
      });
      SpreadsheetApp.flush();
    }
    const restored = cells.every((c, i) => c.getFormula() === originals[i]);
    const med = ms.map((m) => ({ total: median_(m.total), read: median_(m.read), write: median_(m.write), first: median_(m.first), wait: median_(m.wait) }));
    const pct = (a, b) => (a !== '' && b ? Math.round((a / b) * 1000) / 10 + '%' : '');
    const summary = cfg.SHEETS.map((n, i) => ['まとめ', ms[i].total.length + ' 回', n, cfg.CELL, med[i].total, med[i].read, med[i].write, med[i].first, med[i].wait, '', '中央値(ミリ秒)']);
    const ratio = pct(med[0].total, med[1].total);
    summary.push(['まとめ', '', '比', '', ratio, pct(med[0].read, med[1].read), '', '', pct(med[0].wait, med[1].wait), '', '「' + cfg.SHEETS[0] + '」÷「' + cfg.SHEETS[1] + '」(小さいほど速い)']);
    summary.push(['まとめ', '', '元に戻ったか', '', restored ? '戻った' : '戻っていない', '', '', '', '', '', stopped || '最後まで測った']);
    const sh = writeTable_(S.TIMING, H_TIMING, summary.concat(rows));
    log_(title, 'メニュー', restored ? '完了' : '要確認: 元の数式に戻っていない', cfg.SHEETS.map((n, i) => n + ' 合計 ' + med[i].total + 'ms(読むだけ ' + med[i].read + ' / 再計算・反映待ち ' + med[i].wait + ')').join(' / ') + (ratio === '' ? '' : ' 比 ' + ratio));
    activate_(sh);
    const changed = rows.filter((x) => x[9] !== '同じ').length;
    notify_(title, [
      cfg.SHEETS.map((n, i) => '「' + n + '」: 合計 ' + med[i].total + ' ミリ秒 = 再計算・反映待ち 約 ' + med[i].wait + ' + 読むだけ 約 ' + med[i].read + '(中央値、' + ms[i].total.length + ' 回)').join('\n'),
      ratio === '' ? '' : '比: ' + ratio + '(小さいほど速い。このシート1枚を計算し直す時間で、ファイル全体の速さではありません)',
      '',
      restored ? '入れ直したセルは、すべて元の数式に戻りました。' : '⚠ 元の数式に戻っていないセルがあります。シート「' + S.TIMING + '」を確認してください。',
      changed ? '⚠ 値が変更前と違った回が ' + changed + ' 回あります(TODAY() などの影響か確認が必要)。' : '値は毎回、変更前と同じでした。',
      '',
      '結果はシート「' + S.TIMING + '」です。時間は Google 側の混み具合でも変わるので、1回だけで判断しないでください。',
    ]);
  });
}

function median_(list) {
  if (!list.length) return '';
  const a = list.slice().sort((x, y) => x - y);
  const m = Math.floor(a.length / 2);
  return a.length % 2 ? a[m] : Math.round((a[m - 1] + a[m]) / 2);
}

function sameGrid_(a, b) {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    if (a[i].length !== b[i].length) return false;
    for (let j = 0; j < a[i].length; j++) if (cellKey_(a[i][j]) !== cellKey_(b[i][j])) return false;
  }
  return true;
}

function readTargets_() {
  return readRows_(S.TARGETS)
    .filter((r) => r[0] === true && String(r[1]).trim())
    .map((r) => ({ sheet: String(r[1]).trim(), a1: String(r[2] || '').trim(), memo: String(r[3] || '') }));
}

/** 値の比較用の文字列。種類(数値・文字列・日付・真偽)も含める */
function cellKey_(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (typeof v === 'number') return 'n' + v;
  if (typeof v === 'boolean') return v ? 'bTRUE' : 'bFALSE';
  if (Object.prototype.toString.call(v) === '[object Date]') return 'd' + v.getTime();
  return 's' + String(v);
}

function sameKey_(a, b) {
  if (a === b) return true;
  if (a && b && a.charAt(0) === 'n' && b.charAt(0) === 'n') {
    const x = Number(a.slice(1));
    const y = Number(b.slice(1));
    return Math.abs(x - y) <= 1e-9 * Math.max(1, Math.abs(x), Math.abs(y));
  }
  return false;
}

function showKey_(k) {
  if (!k) return '(空白)';
  const t = k.charAt(0);
  const body = k.slice(1);
  if (t === 'd') return Utilities.formatDate(new Date(Number(body)), tz_(), 'yyyy/MM/dd HH:mm:ss');
  if (t === 'n') return body;
  if (t === 'b') return body;
  return '"' + body.slice(0, 100) + '"';
}

// ============================================================
// 共通
// ============================================================

function loadState_() {
  const s = PropertiesService.getScriptProperties().getProperty('LT_STATE');
  return s ? JSON.parse(s) : null;
}

function saveState_(st) {
  st.updatedAt = Date.now();
  const copy = Object.assign({}, st);
  delete copy.pendingAcc; // 途中のシートの集計は「_作業_状態」シートに保存する
  PropertiesService.getScriptProperties().setProperty('LT_STATE', JSON.stringify(copy));
}

function saveStateQuiet_(st) {
  try {
    const saved = loadState_();
    if (saved) {
      saved.errors = st.errors;
      saveState_(saved);
    }
  } catch (e) {
    // 状態の保存に失敗しても、最後に保存した状態から続けられる
  }
}

function pushError_(st, text) {
  st.errors = (st.errors || []).concat([nowText_() + ' ' + String(text).slice(0, 150)]).slice(-5);
}

function loadAcc_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(S.WORK);
  if (!sh || sh.getLastRow() < 1) return null;
  const v = sh.getRange(1, 1, 2, 1).getValues();
  if (!v[0][0]) return null;
  const acc = JSON.parse(String(v[0][0]));
  if (acc.rowFp) acc.rowFp = v[1][0] ? String(v[1][0]).split(',') : [];
  return acc;
}

function saveAcc_(acc) {
  const sh = ensureSheet_(S.WORK, null, true);
  if (!acc) {
    writeRows_(sh, 1, [[''], ['']], 1);
    return;
  }
  const copy = Object.assign({}, acc);
  const fp = acc.rowFp ? acc.rowFp.join(',') : '';
  copy.rowFp = acc.rowFp ? [] : null;
  let json = JSON.stringify(copy);
  if (json.length > 44000) {
    copy.notes = copy.notes.slice(0, 10).concat(['(メモが多いため省略)']);
    copy.ex = copy.ex.slice(0, 20);
    json = JSON.stringify(copy);
  }
  if (json.length > 44000 || fp.length > 44000) throw new Error('途中の集計が大きすぎて保存できません(シート「' + acc.name + '」)。LT.ROW_FINGERPRINT_MAX_ROWS を小さくしてください。');
  writeRows_(sh, 1, [[json], [fp]], 1);
}

function isStopRequested_() {
  return !!PropertiesService.getScriptProperties().getProperty('LT_STOP');
}

/** 続きの予約は1つだけにする(重複して作らない) */
function scheduleContinue_() {
  deleteTriggers_(CONTINUE_HANDLER);
  if (isStopRequested_()) return;
  ScriptApp.newTrigger(CONTINUE_HANDLER).timeBased().after(LT.CONTINUE_AFTER_MS).create();
}

function deleteTriggers_(handler) {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === handler) {
      ScriptApp.deleteTrigger(t);
      n++;
    }
  });
  return n;
}

function withLock_(title, fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) {
    notify_(title, ['ほかの処理(自動の続きなど)が動いています。少し待ってからもう一度実行してください。']);
    return;
  }
  try {
    return fn();
  } catch (e) {
    log_(title, '', 'エラー', errText_(e));
    notify_(title, ['エラーで止まりました。', errText_(e), '', '1〜10 は対象ファイルに書き込まないので、対象ファイルは変わっていません。', '11 の途中で止まったときは、入れ直したセルを元の数式に戻してから止めます(シート「' + S.TIMING + '」とログで確認できます)。']);
  } finally {
    lock.releaseLock();
  }
}

function ensureSheet_(name, header, hidden) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name, ss.getSheets().length);
    if (hidden) sh.hideSheet();
  }
  if (header) {
    const width = Math.min(header.length, sh.getMaxColumns());
    const cur = sh.getLastRow() >= 1 ? sh.getRange(1, 1, 1, width).getValues()[0] : [];
    if (cur.join('\u0001') !== header.join('\u0001')) {
      writeRows_(sh, 1, [header], header.length);
      styleHeader_(sh, 1, header.length);
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

function writeTable_(name, header, rows) {
  const sh = ensureSheet_(name, null);
  sh.clearContents();
  writeRows_(sh, 1, [header].concat(rows), header.length);
  styleHeader_(sh, 1, header.length);
  sh.setFrozenRows(1);
  return sh;
}

/** 2行目以降を読む(見出しは除く) */
function readRows_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(1, sh.getLastColumn())).getValues();
}

/** 先に上書きしてから、余った行だけを消す(書き込みの途中で失敗しても、前の結果が消えないように) */
function rewriteRows_(name, header, rows) {
  const sh = ensureSheet_(name, header);
  const last = sh.getLastRow();
  if (rows.length) writeRows_(sh, 2, rows.map((r) => r.slice(0, header.length)), header.length);
  const extra = last - (rows.length + 1);
  if (extra > 0) sh.getRange(rows.length + 2, 1, extra, Math.min(sh.getMaxColumns(), Math.max(header.length, sh.getLastColumn()))).clearContent();
}

/** 1列目が keys に入っている行を、新しい行で置き換える(同じシートを2回調べても重複しない) */
function upsertRows_(name, header, keys, rows, hidden) {
  const sh = ensureSheet_(name, header, hidden);
  const old = readRows_(name).filter((r) => !keys.has(String(r[0])));
  rewriteRows_(name, header, old.concat(rows));
  return sh;
}

function replaceRows_(name, header, isOld, rows) {
  ensureSheet_(name, header);
  const keep = readRows_(name).filter((r) => !isOld(r));
  rewriteRows_(name, header, keep.concat(rows));
}

function writeRows_(sh, startRow, rows, width) {
  if (!rows.length) return;
  const need = startRow + rows.length - 1;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  if (sh.getMaxColumns() < width) sh.insertColumnsAfter(sh.getMaxColumns(), width - sh.getMaxColumns());
  const out = rows.map((r) => {
    const row = [];
    for (let j = 0; j < width; j++) row.push(cellOut_(r[j]));
    return row;
  });
  sh.getRange(startRow, 1, rows.length, width).setValues(out);
}

/** 書き込んだ文字列が数式や数値に変わらないように、先頭に ' を付ける */
function cellOut_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') return isFinite(v) ? v : String(v);
  if (typeof v !== 'string') return v;
  let s = v.length > 45000 ? v.slice(0, 45000) + '…' : v;
  if (/^[0-9=+\-@'.$¥￥(#\s]/.test(s) || /^(true|false)$/i.test(s)) s = "'" + s;
  return s;
}

function styleHeader_(sh, row, width) {
  try {
    sh.getRange(row, 1, 1, width).setFontWeight('bold').setBackground('#e8eaed');
  } catch (e) {
    // 見た目だけなので失敗しても続ける
  }
}

function log_(what, origin, result, detail) {
  try {
    const sh = ensureSheet_(S.LOG, H_LOG);
    writeRows_(sh, sh.getLastRow() + 1, [[nowText_(), what, origin, result, String(detail || '').slice(0, 500)]], H_LOG.length);
  } catch (e) {
    console.log('ログを書けなかった: ' + errText_(e));
  }
}

function uiOrNull_() {
  try {
    return SpreadsheetApp.getUi();
  } catch (e) {
    return null;
  }
}

function needMenu_(title) {
  console.log(title + ': スプレッドシートのメニュー「' + LT.MENU + '」から実行してください。');
}

function confirm_(title, lines) {
  const ui = uiOrNull_();
  if (!ui) {
    needMenu_(title);
    return false;
  }
  return ui.alert(title, lines.join('\n'), ui.ButtonSet.YES_NO) === ui.Button.YES;
}

function notify_(title, lines) {
  console.log(title + '\n' + lines.join('\n'));
  const ui = uiOrNull_();
  if (!ui) return;
  try {
    ui.alert(title, lines.slice(0, 40).join('\n'), ui.ButtonSet.OK);
  } catch (e) {
    // 自動の続き(トリガー)から呼ばれたときは画面を出せない
  }
}

function toast_(message) {
  try {
    SpreadsheetApp.getActive().toast(message, '軽量化ツール', 5);
  } catch (e) {
    // 画面がないときは何もしない
  }
}

function activate_(sh) {
  if (!sh || !uiOrNull_()) return;
  try {
    sh.activate();
  } catch (e) {
    // 画面がないときは何もしない
  }
}

function safe_(fn) {
  try {
    return fn();
  } catch (e) {
    return '未取得';
  }
}

function errText_(e) {
  return String(e && e.message ? e.message : e).slice(0, 300);
}

function tz_() {
  try {
    return SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  } catch (e) {
    return Session.getScriptTimeZone();
  }
}

function nowText_() {
  return Utilities.formatDate(new Date(), tz_(), 'yyyy/MM/dd HH:mm:ss');
}

function fmtTime_(ms) {
  return ms ? Utilities.formatDate(new Date(ms), tz_(), 'yyyy/MM/dd HH:mm:ss') : '';
}

function fmtNum_(n) {
  return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function pct_(a, b) {
  return b ? Math.round((a / b) * 1000) / 10 + '%' : '-';
}

function orBlank_(v) {
  return v === undefined ? '未取得' : v;
}

function bump_(arr, r, c) {
  if (r > arr[0]) arr[0] = r;
  if (c > arr[1]) arr[1] = c;
}

function indexOf_(header) {
  const o = {};
  header.forEach((h, i) => (o[h[0]] = i));
  return o;
}

function countList_(list) {
  const m = {};
  list.forEach((x) => (m[x] = (m[x] || 0) + 1));
  return Object.keys(m).map((k) => [k, m[k]]);
}

function isErrorValue_(v) {
  return /^#(N\/A|REF!|DIV\/0!|VALUE!|NAME\?|NUM!|NULL!|ERROR!|SPILL!|CALC!)/.test(v);
}

function isLoading_(v) {
  return typeof v === 'string' && (v === 'Loading...' || v.indexOf('読み込み中') === 0 || v.indexOf('読み込んでいます') === 0);
}

function rangesText_(list) {
  if (!list.length) return '';
  const parts = list.slice(0, 20).map((x) => (x[0] === x[1] ? String(x[0]) : x[0] + '-' + x[1]));
  return parts.join(', ') + (list.length > 20 ? ' …ほか' : '');
}

/** [3,4,5,9] → "3-5, 9" */
function rowsText_(rows) {
  const out = [];
  for (let i = 0; i < rows.length; i++) {
    let j = i;
    while (j + 1 < rows.length && rows[j + 1] === rows[j] + 1) j++;
    out.push(i === j ? String(rows[i]) : rows[i] + '-' + rows[j]);
    i = j;
  }
  return out.slice(0, 20).join(', ') + (out.length > 20 ? ' …ほか' : '');
}

function a1_(row, col) {
  return columnLetter_(col) + row;
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

function colIndex_(letters) {
  let n = 0;
  for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
  return n;
}

/** 53ビットの軽いハッシュ(値や数式が変わったかを見分けるための指紋。暗号用ではない) */
function hash53_(str) {
  let h1 = 0xdeadbeef;
  let h2 = 0x41c6ce57;
  for (let i = 0; i < str.length; i++) {
    const ch = str.charCodeAt(i);
    h1 = Math.imul(h1 ^ ch, 2654435761);
    h2 = Math.imul(h2 ^ ch, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507);
  h1 ^= Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507);
  h2 ^= Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(36);
}
