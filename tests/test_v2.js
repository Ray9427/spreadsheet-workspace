// 軽量化ツール v2 のローカル試験(匿名の作り物のデータだけを使う)
//   node tests/test_v2.js
'use strict';
const path = require('path');
const assert = require('assert');
const { createEnv, colLetter } = require('./mock_gas');

const GS = path.join(__dirname, '..', 'gas', '軽量化ツール_v2.gs');
const ID = { PROD: 'PRODFILE_anon_000000000001', BACKUP: 'BACKUPFILE_anon_0000000002', TEST: 'TESTFILE_anon_000000000003' };
const url = (id) => 'https://docs.google.com/spreadsheets/d/' + id + '/edit#gid=0';
const STORE = '店舗A';
let passed = 0;
const plain = (o) => JSON.parse(JSON.stringify(o)); // vm の中で作ったオブジェクトを比べられる形にする
function test(name, fn) {
  fn();
  passed++;
  console.log('OK  ' + name);
}

// ---------------- 匿名の試験用ブック ----------------
const SUMIFS_F = (L) => `=IFERROR(LET(total, SUMIFS('編明細'!$R:$R, '編明細'!$C:$C, "${STORE}", '編明細'!$M:$M, $D$6, '編明細'!$N:$N, $K$6, '編明細'!$O:$O, ${L}$2, '編明細'!$P:$P, $C$9, '編明細'!$Q:$Q, "A"), IF(total=0, "", total)), "")`;
const SUMIFS_VARIANT = (L) => SUMIFS_F(L).replace("'編明細'!$R:$R", "'編明細'!$S:$S");
const FILTER_F = (L) => `=IFERROR(INDEX(FILTER('編打刻'!$AR:$AR, ('編打刻'!$AK:$AK="${STORE}") * (SUBSTITUTE('編打刻'!$AM:$AM,"年","")=SUBSTITUTE($D$6,"年","")) * (SUBSTITUTE('編打刻'!$AN:$AN,"月","")=SUBSTITUTE($K$6,"月","")) * ('編打刻'!$AO:$AO=${L}$2) * (ISNUMBER(SEARCH($C$9, '編打刻'!$AP:$AP)))), 1), "")`;

function put(s, r, c, f, v) {
  if (f) (s.f[r] = s.f[r] || {})[c] = f;
  if (v !== undefined) (s.v[r] = s.v[r] || {})[c] = v;
}

function personSheet(name, id, o) {
  const s = { name, id, maxRows: o.rows || 60, maxCols: 20, f: {}, v: {}, hidden: !!o.hidden };
  put(s, 6, 4, null, '2026年');
  put(s, 6, 11, null, '9月');
  put(s, 9, 3, null, o.person);
  for (let c = 12; c <= 20; c++) put(s, 2, c, `='000'!${colLetter(c)}2`, String(c - 11) + '日');
  for (let r = 12; r <= 40; r++) {
    for (let c = 5; c <= 15; c++) {
      const L = colLetter(c);
      const f = o.variant && r >= 20 && r <= 22 ? SUMIFS_VARIANT(L) : SUMIFS_F(L);
      put(s, r, c, f, (r * c) % 7 === 0 ? '' : r * 100 + c);
    }
  }
  for (let r = 41; r <= 50; r++) for (let c = 5; c <= 10; c++) put(s, r, c, FILTER_F(colLetter(c)), '');
  put(s, 55, 2, '=TODAY()', new Date(Date.UTC(2026, 9, 15)));
  return s;
}

function workbook(id, name, tweak) {
  const sheets = [];
  const h = { name: '000', id: 100000001, maxRows: 10, maxCols: 20, f: {}, v: {} };
  for (let c = 1; c <= 20; c++) put(h, 2, c, null, '見出し' + c);
  sheets.push(h);
  sheets.push(personSheet('001', 1487000001, { person: 'A01' }));
  sheets.push(personSheet('002', 1487000002, { person: 'A02' }));
  sheets.push(personSheet('003', 1487000003, { person: 'A03', variant: true }));
  sheets.push(personSheet('004', 1487000004, { person: 'A04', hidden: true }));
  sheets.push(personSheet('005', 1487000005, { person: 'A05', rows: 59 }));
  const m = { name: '編明細', id: 200000001, maxRows: 300, maxCols: 27, f: {}, v: {} };
  for (let r = 2; r <= 250; r++) {
    put(m, r, 3, null, r % 3 ? STORE : '店舗B');
    put(m, r, 13, null, '2026年');
    put(m, r, 14, null, '9月');
    put(m, r, 15, null, (r % 30) + 1 + '日');
    put(m, r, 16, null, 'A0' + (r % 5));
    put(m, r, 17, null, 'A');
    put(m, r, 18, null, r * 10);
  }
  sheets.push(m);
  const d = { name: '編打刻', id: 200000002, maxRows: 500, maxCols: 49, f: {}, v: {} };
  for (let r = 2; r <= 400; r++) [37, 39, 40, 41, 42, 44, 46].forEach((c) => put(d, r, c, null, c === 44 ? r : STORE + r));
  sheets.push(d);
  const t = {
    name: '集計', id: 300000001, maxRows: 30, maxCols: 10, f: {}, v: {},
    dv: {}, merges: ['A1:C1'],
    cf: [{ ranges: ['D2:D25'], type: 'CUSTOM_FORMULA', formula: '=INDIRECT("x!A1")>0' }],
    drawings: [{ action: 'oldMacro', row: 28, col: 9 }],
  };
  put(t, 1, 1, null, '集計表');
  put(t, 2, 2, "=SUM('編明細'!$R$2:$R)", 123450);
  put(t, 3, 2, '=INDIRECT("\'"&A3&"\'!B2")', 5);
  put(t, 4, 2, '=IMPORTRANGE("https://docs.google.com/spreadsheets/d/ANONYMOUSFILEID00000000000/edit","シートA!A1:C10")', '#REF!');
  put(t, 5, 2, '=TODAY()', new Date(Date.UTC(2026, 9, 15)));
  for (let r = 2; r <= 10; r++) put(t, r, 3, null, r * 1000);
  for (let r = 2; r <= 20; r++) t.dv[r + ',1'] = true;
  sheets.push(t);
  sheets.push({ name: 'グラフ1', id: 300000002, type: 'OBJECT', f: {}, v: {}, charts: [{ sources: [{ sheet: '集計' }] }] });
  const x = { name: 'シートX', id: 300000003, maxRows: 20, maxCols: 5, f: {}, v: {}, hidden: true };
  for (let r = 1; r <= 10; r++) put(x, r, 1, null, 'メモ' + r);
  sheets.push(x);
  const spec = {
    id, name, sheets,
    named: [{ name: '集計範囲', sheet: '集計', a1: 'A1:B5' }],
    protections: [{ type: 'RANGE', sheet: '集計', a1: 'A1:J3' }],
  };
  if (tweak) tweak(spec);
  return spec;
}

function newEnv(opts) {
  opts = opts || {};
  const env = createEnv({});
  env.addTarget(workbook(ID.PROD, '260900 歩合表(匿名)', (s) => (s.sheets.find((x) => x.name === '集計').v[2][3] = 99999)));
  env.addTarget(workbook(ID.BACKUP, '260900 歩合表_バックアップ(匿名)', (s) => (s.sheets.find((x) => x.name === '編打刻').maxRows = 480)));
  env.testSpec = workbook(ID.TEST, '【試験用】260900 歩合表(匿名)', opts.tweakTest);
  env.addTarget(env.testSpec);
  env.load(GS);
  const LT = env.ctx.LT;
  LT.CHUNK_CELLS = 2000;
  LT.MIN_CHUNK_ROWS = 10;
  LT.FLUSH_EVERY_SHEETS = 3;
  LT.CONTINUE_AFTER_MS = 60000;
  if (opts.timeLimit) {
    LT.TIME_LIMIT_MS = opts.timeLimit;
    LT.FINALIZE_RESERVE_MS = 1500;
    LT.FLUSH_EVERY_MS = 2500;
  } else {
    LT.TIME_LIMIT_MS = 1e12;
  }
  env.prompts.push(url(ID.PROD), url(ID.BACKUP), url(ID.TEST));
  env.ctx.lt_setup();
  if (opts.auto) env.props.LT_AUTO = '1'; // 自動の続きをオンにした状態(既定はオフ)
  return env;
}

const state = (env) => JSON.parse(env.props.LT_STATE || 'null');
const header = (env, n) => env.rows(n)[0];
function rowOf(env, sheetName, keyCol, key) {
  const rows = env.rows(sheetName);
  const h = rows[0];
  const r = rows.find((x) => String(x[keyCol]) === key);
  if (!r) throw new Error('行がない: ' + sheetName + ' ' + key);
  const o = {};
  h.forEach((k, i) => (o[k] = r[i]));
  return o;
}
function snapshot(env) {
  const S = env.ctx.S;
  const out = {};
  [S.DETAIL, S.FN, S.COLREF, S.REFS, S.EXT, S.STRUCT, S.SUMMARY, S.ROWFP].forEach((n) => {
    let rows = env.rows(n).map((r) => r.slice());
    if (n === S.DETAIL) {
      const drop = ['調査日時', '処理秒数', 'メモ'].map((x) => rows[0].indexOf(x));
      rows = rows.map((r) => r.map((v, i) => (drop.indexOf(i) >= 0 ? '' : v)));
    }
    if (n === S.SUMMARY) rows = rows.filter((r) => r[0] !== '調査の期間');
    out[n] = rows.map((r) => r.map((v) => (v instanceof Date ? v.getTime() : v)));
  });
  return JSON.stringify(out);
}
/** 2つの snapshot の最初の違いを 5 件まで文章にする(試験が落ちたときの手がかり) */
function snapDiff(a, b) {
  const x = JSON.parse(a);
  const y = JSON.parse(b);
  const out = [];
  Object.keys(x).forEach((n) => {
    const rx = x[n] || [];
    const ry = y[n] || [];
    if (rx.length !== ry.length) out.push(n + ': 行数 ' + rx.length + ' / ' + ry.length);
    for (let i = 0; i < Math.min(rx.length, ry.length); i++) {
      for (let j = 0; j < Math.max(rx[i].length, ry[i].length); j++) {
        if (JSON.stringify(rx[i][j]) !== JSON.stringify(ry[i][j])) out.push(n + ' 行' + (i + 1) + ' 列' + (j + 1) + '(' + rx[0][j] + '): ' + JSON.stringify(rx[i][j]) + ' / ' + JSON.stringify(ry[i][j]));
      }
    }
  });
  return out.slice(0, 5).join('\n');
}
const sameAsBaseline = (env) => {
  const s = snapshot(env);
  assert.ok(s === baseline, '1回で通したときと結果が違う:\n' + snapDiff(s, baseline));
};
function noProblems(env) {
  assert.deepStrictEqual(env.violations, [], '用意していないメソッドや対象への書き込みがあった');
  assert.deepStrictEqual(env.formulaWrites, [], 'ツールのファイルに数式として書き込まれた文字列がある');
  const errs = env.alerts.filter((a) => /エラーで止まりました/.test(a.msg));
  assert.deepStrictEqual(errs.map((a) => a.msg), [], 'エラーの画面が出た');
}
function runToEnd(env, maxSteps) {
  let steps = 0;
  let maxTriggers = 0;
  let midSheet = false; // 1枚のシートの途中で区切った回があったか
  while (state(env) && state(env).phase !== 'done' && steps++ < (maxSteps || 500)) {
    maxTriggers = Math.max(maxTriggers, env.triggers.length);
    if (!env.fireDueTriggers()) break;
    maxTriggers = Math.max(maxTriggers, env.triggers.length);
    const w = env.sheet(env.ctx.S.WORK);
    if (w && w.grid[0] && w.grid[0][0]) midSheet = true;
  }
  return { steps, maxTriggers, midSheet };
}

// ---------------- 1. 数式の解析 ----------------
test('数式の解析: SUMIFS 型', () => {
  const env = createEnv({});
  env.load(GS);
  const a = env.ctx.analyzeFormula_(SUMIFS_F('L'));
  assert.strictEqual(Object.values(a.colRefs).reduce((x, y) => x + y, 0), 7);
  assert.deepStrictEqual(plain(Object.keys(a.sheets)), ['編明細']);
  assert.ok(a.cond && !a.array && !a.text);
  assert.ok(a.fn.LET && a.fn.SUMIFS && a.fn.IFERROR && a.fn.IF && !a.fn.total);
});

test('数式の解析: FILTER 型(列全体に SUBSTITUTE・SEARCH)', () => {
  const env = createEnv({});
  env.load(GS);
  const a = env.ctx.analyzeFormula_(FILTER_F('AA'));
  assert.strictEqual(Object.values(a.colRefs).reduce((x, y) => x + y, 0), 6);
  assert.ok(a.array && a.text && !a.cond);
  assert.deepStrictEqual(plain(Object.keys(a.sheets)), ['編打刻']);
});

test('数式の解析: 文字列・シート名の中・INDIRECT・IMPORTRANGE・A2:A', () => {
  const env = createEnv({});
  env.load(GS);
  const c = env.ctx;
  let a = c.analyzeFormula_(`=COUNTIF(A1:A9,"'編打刻'!$R:$R")+SUM('編打刻'!$R:$R)`);
  assert.deepStrictEqual(plain(a).colRefs, { "'編打刻'!$R:$R": 1 });
  a = c.analyzeFormula_(`=SUM('集計(月)'!A:A)+LEN('集計(月)'!B:B)`);
  assert.deepStrictEqual(plain(Object.keys(a.sheets)), ['集計(月)']);
  assert.ok(a.text, 'LEN に列全体を渡している');
  assert.ok(!a.fn['月'], 'シート名の中のかっこを関数と間違えない');
  a = c.analyzeFormula_(`=INDIRECT("'"&A1&"'!B2")`);
  assert.deepStrictEqual(plain(a).dynamic, ['INDIRECT']);
  assert.deepStrictEqual(plain(a).sheets, {});
  a = c.analyzeFormula_('=IMPORTRANGE("https://docs.google.com/spreadsheets/d/ANONYMOUSFILEID00000000000/edit","シートA!A1:C10")');
  assert.deepStrictEqual(plain(a).imports.map((x) => [x.id, x.range]), [['ANONYMOUSFILEID00000000000', 'シートA!A1:C10']]);
  a = c.analyzeFormula_('=IMPORTRANGE(B1, "x!A1")');
  assert.ok(/未確認/.test(a.imports[0].id));
  a = c.analyzeFormula_("=SUM('編明細'!$R$2:$R)+SUM(A2:A10)+SUM(B:B)");
  assert.deepStrictEqual(plain(a).openRefs, { "'編明細'!$R$2:$R": 1 });
  assert.deepStrictEqual(plain(a).colRefs, { 'B:B': 1 });
  a = c.analyzeFormula_('=TODAY()-DATEDIF(A1,B1,"M")');
  assert.deepStrictEqual(plain(a).volatile, ['TODAY']);
  a = c.analyzeFormula_("=SUM('編明細'!$C:$CA)");
  assert.deepStrictEqual(plain(Object.keys(a.colRefs)), ["'編明細'!$C:$CA"]);
});

test('書き込む文字列が数式・数値に変わらない', () => {
  const env = createEnv({});
  env.load(GS);
  const c = env.ctx;
  ['001', '=SUM(A:A)', '+1', '-x', '#REF!', 'TRUE', ' 1', '1e5', '20261015_101500', '¥100'].forEach((s) => assert.strictEqual(c.cellOut_(s).charAt(0), "'", s));
  ['編明細', 'abc', '{"a":1}', '["n1"]', ''].forEach((s) => assert.strictEqual(c.cellOut_(s), s));
});

// ---------------- 2. 登録とバックアップの確認 ----------------
test('登録: 3つの URL を登録できる / 同じファイルの二重登録は断る', () => {
  const env = newEnv();
  assert.strictEqual(env.props.LT_PROD_ID, ID.PROD);
  assert.strictEqual(env.props.LT_TEST_ID, ID.TEST);
  env.prompts.push(url(ID.PROD), url(ID.PROD), url(ID.TEST));
  env.ctx.lt_setup();
  assert.ok(env.alerts.some((a) => a.title === '登録できません'));
  assert.strictEqual(env.props.LT_BACKUP_ID, ID.BACKUP, '断ったときは登録を変えない');
  env.prompts.push(url(env.tool.id), '', '');
  env.ctx.lt_setup();
  assert.ok(env.alerts.filter((a) => a.title === '登録できません').length === 2, 'ツール自身は登録できない');
  noProblems(env);
});

test('バックアップの確認: 開けたか・構成の違いを記録する', () => {
  const env = newEnv();
  env.ctx.lt_checkBackup();
  const rows = env.rows(env.ctx.S.SETUP);
  assert.deepStrictEqual(rows.slice(1, 4).map((r) => r[3]), ['開けた', '開けた', '開けた']);
  assert.strictEqual(rows[2][7], 1, '保管用は 編打刻 の行数が違う');
  assert.strictEqual(rows[3][7], 0, '試験用は同じ');
  assert.ok(rows.some((r) => r[1] === '編打刻' && /大きさが違う/.test(r[2])));
  noProblems(env);
});

test('軽い診断: シート一覧(名前が 001 でも文字のまま)', () => {
  const env = newEnv();
  env.ctx.lt_diagLight();
  const rows = env.rows(env.ctx.S.LIST);
  assert.strictEqual(rows.length, 1 + 11);
  assert.strictEqual(rows[2][1], '001');
  assert.strictEqual(rows[2][10], '3桁の数字');
  assert.ok(rows.some((r) => r[1] === 'グラフ1' && /表ではない/.test(r[10])));
  noProblems(env);
});

// ---------------- 3. 詳しい診断 ----------------
let baseline = null;
test('詳しい診断: 1回で最後まで', () => {
  const env = newEnv();
  env.ctx.lt_diagStart();
  assert.strictEqual(state(env).phase, 'done');
  const S = env.ctx.S;
  const d = rowOf(env, S.DETAIL, 1, '001');
  assert.strictEqual(d['数式セル'], 319 + 60 + 9 + 1);
  assert.strictEqual(d['列全体参照(個)'], 319 * 7 + 60 * 6);
  assert.strictEqual(d['条件集計×列全体(セル)'], 319);
  assert.strictEqual(d['配列計算×列全体(セル)'], 60);
  assert.strictEqual(d['列全体を1行ずつ加工(セル)'], 60);
  assert.strictEqual(d['揮発性関数(セル)'], 1);
  assert.strictEqual(d['数式の最終行'], 55);
  assert.strictEqual(d['参照しているシート'], '編明細(319), 編打刻(60), 000(9)');
  const sum = rowOf(env, S.DETAIL, 1, '集計');
  assert.strictEqual(sum['A2:A 形式の参照(個)'], 1);
  assert.strictEqual(sum['INDIRECT・OFFSET(セル)'], 1);
  assert.strictEqual(sum['外部取得関数(セル)'], 1);
  assert.strictEqual(sum['入力規則セル'], 19);
  assert.strictEqual(sum['入力規則の最終行'], 20);
  assert.strictEqual(sum['結合'], 1);
  assert.strictEqual(sum['うちカスタム数式'], 1);
  assert.strictEqual(sum['図形に割り当てたスクリプト'], 'oldMacro');
  assert.strictEqual(sum['内容と構造物の最終行(参考)'], 28, '図形の位置まで含める');
  assert.strictEqual(sum['エラー値(セル)'], 1);
  const m = rowOf(env, S.DETAIL, 1, '編明細');
  assert.strictEqual(m['定数の最終行'], 250);
  assert.strictEqual(m['参照されているシート数'], 6, '001〜005 と 集計');
  assert.strictEqual(rowOf(env, S.DETAIL, 1, 'シートX')['参照されているシート数'], 0);
  assert.strictEqual(rowOf(env, S.DETAIL, 1, '集計')['参照されているシート数'], 1, 'グラフ1 から');
  const st = rowOf(env, S.STRUCT, 0, '002');
  assert.strictEqual(st['判定'], '代表と同じ');
  const s3 = rowOf(env, S.STRUCT, 0, '003');
  assert.strictEqual(s3['判定'], '一部だけ違う');
  assert.strictEqual(s3['違う行(例)'], '20-22');
  assert.strictEqual(rowOf(env, S.STRUCT, 0, '005')['違う行(例)'], '60', '行数の違いも違う行として出す');
  assert.strictEqual(rowOf(env, S.STRUCT, 0, '004')['判定'], '代表と同じ');
  const ext = env.rows(S.EXT);
  assert.ok(ext.some((r) => r[2] === '外部取得' && r[6] === 'ANONYMOUSFILEID00000000000' && r[7] === 'シートA!A1:C10'));
  assert.ok(ext.some((r) => r[2] === '図形に割り当てたスクリプト' && r[3] === 'oldMacro'));
  assert.ok(ext.some((r) => r[2] === '条件付き書式のINDIRECT'));
  const cr = env.rows(S.COLREF).find((r) => r[1] === '001' && r[3] === "'編明細'!$R:$R");
  assert.deepStrictEqual([cr[2], cr[5], cr[6], cr[7], cr[8]], ['編明細', 29 * 11, 1, 300, 319 * 300]);
  const summary = env.rows(S.SUMMARY).map((r) => r[0]);
  ['■ ボトルネックの候補(数は実測、効き目は未測定)', '■ 調べていないこと・未確認', '列全体参照が多い', '非表示のシートの数式'].forEach((k) => assert.ok(summary.indexOf(k) >= 0, k));
  assert.ok(env.triggers.length === 0, '完了したら続きの予約はない');
  noProblems(env);
  baseline = snapshot(env);
});

test('詳しい診断: 時間で区切り、自動で続ける(結果は1回で通したときと同じ)', () => {
  const env = newEnv({ timeLimit: 4000, auto: true });
  env.ctx.lt_diagStart();
  const r = runToEnd(env);
  assert.strictEqual(state(env).phase, 'done');
  assert.ok(state(env).runs >= 4, '何回かに分かれた: ' + state(env).runs);
  assert.ok(r.maxTriggers <= 1, '続きの予約は同時に1つまで');
  assert.ok(r.midSheet, '1枚のシートの途中で区切って、続きから読んだ');
  sameAsBaseline(env);
  noProblems(env);
});

test('詳しい診断: 途中でエラーが起きても、再開すれば重複なく同じ結果', () => {
  const env = newEnv({ timeLimit: 4000, auto: true });
  env.ctx.lt_diagStart();
  env.fireDueTriggers();
  env.failToolWriteAt = env.toolWrites + 4; // 次の実行の途中で書き込みを1回失敗させる
  env.fireDueTriggers();
  assert.strictEqual(env.triggers.length, 0, 'エラーのときは自動で続けない');
  assert.notStrictEqual(state(env).phase, 'done');
  assert.ok(state(env).errors.length >= 1);
  env.ctx.lt_diagStart(); // 人がメニューから再開
  runToEnd(env);
  assert.strictEqual(state(env).phase, 'done');
  sameAsBaseline(env);
  assert.deepStrictEqual(env.violations, []);
});

test('診断を止める: 予約を消し、新しい処理を始めない / 4. で再開できる', () => {
  const env = newEnv({ timeLimit: 4000, auto: true });
  env.ctx.lt_diagStart();
  assert.strictEqual(env.triggers.length, 1);
  const before = state(env);
  env.ctx.lt_diagStop();
  assert.strictEqual(env.triggers.length, 0);
  assert.strictEqual(env.fireDueTriggers(), false);
  env.ctx.ScriptApp.newTrigger('lt_diagContinueTrigger').timeBased().after(1000).create(); // 消し損ねた予約があっても
  env.fireDueTriggers();
  assert.strictEqual(state(env).qi, before.qi, '止める指示のあとは進まない');
  env.ctx.lt_diagStart();
  runToEnd(env);
  assert.strictEqual(state(env).phase, 'done');
  sameAsBaseline(env);
  noProblems(env);
});

test('既定では自動の続きはオフ: 区切っても予約を作らず、手動で続ければ同じ結果', () => {
  const env = newEnv({ timeLimit: 4000 });
  env.ctx.lt_diagStart();
  assert.strictEqual(env.triggers.length, 0, '既定では予約を作らない');
  assert.notStrictEqual(state(env).phase, 'done', '1回では終わらない大きさで試す');
  assert.ok(env.alerts.some((a) => /もう一度押してください/.test(a.msg)), '手動で続ける案内を出す');
  const qi = state(env).qi;
  const runs = state(env).runs;
  env.ctx.ScriptApp.newTrigger('lt_diagContinueTrigger').timeBased().after(1000).create(); // 前の版の予約が残っていても
  env.fireDueTriggers();
  assert.strictEqual(state(env).runs, runs, 'オフのときは予約から実行しない');
  assert.strictEqual(state(env).qi, qi);
  assert.ok(env.rows(env.ctx.S.LOG).some((r) => r[3] === '自動の続きがオフのため実行しない'));
  let n = 0;
  while (state(env).phase !== 'done' && n++ < 200) env.ctx.lt_diagStart();
  assert.strictEqual(state(env).phase, 'done');
  assert.strictEqual(env.triggers.length, 0);
  sameAsBaseline(env);
  noProblems(env);
});

test('小さな試験: 代表シートだけを読み、予約を作らず、詳しい診断の状態を変えない', () => {
  const read = new Set();
  const env = newEnv({
    tweakTest: (spec) => spec.sheets.forEach((s) => (s.fail = () => (read.add(s.name), false))),
  });
  env.ctx.lt_smallTest();
  assert.deepStrictEqual([...read], ['001'], '読んだのは代表シートだけ');
  assert.strictEqual(env.triggers.length, 0);
  assert.strictEqual(env.props.LT_STATE, undefined, '詳しい診断の状態を作らない');
  const S = env.ctx.S;
  const val = (item) => {
    const r = env.rows(S.SMALL).find((x) => x[0] === item);
    if (!r) throw new Error('行がない: ' + item);
    return r[1];
  };
  assert.strictEqual(val('結果'), '最後まで読めた');
  assert.strictEqual(val('数式のセル'), 319 + 60 + 9 + 1, '詳しい診断の 001 と同じ数');
  assert.strictEqual(val('列全体参照'), '13 種類 / ' + (319 * 7 + 60 * 6) + ' 個');
  assert.strictEqual(val('SUMIFS'), 319);
  assert.strictEqual(val('FILTER'), 60);
  assert.strictEqual(val('自動の続きの予約'), '0 件');
  assert.ok(env.rows(S.SMALL).find((x) => x[0] === '数式を読む(getFormulas)')[2] >= 0, '所要ミリ秒を書く');
  assert.ok(env.props.LT_SMALL_OK, '成功を記録する');
  noProblems(env);

  // 途中で詳しい診断をしていても、その状態は変えない
  const env2 = newEnv({ timeLimit: 4000 });
  env2.ctx.lt_diagStart();
  const before = env2.props.LT_STATE;
  env2.ctx.lt_smallTest();
  assert.strictEqual(env2.props.LT_STATE, before);
  assert.strictEqual(env2.triggers.length, 0);
  noProblems(env2);

  // 代表シートがなければ、エラーとして記録する(対象は変えない)
  const env3 = newEnv({ tweakTest: (spec) => (spec.sheets.find((s) => s.name === '001').name = '001_旧') });
  env3.ctx.lt_smallTest();
  assert.strictEqual(env3.rows(S.SMALL).find((x) => x[0] === '結果')[1], 'エラーで止まった');
  assert.strictEqual(env3.props.LT_SMALL_OK, undefined);
  assert.deepStrictEqual(env3.violations, []);
});

test('自動の続きのオン・オフ: 確認前は注意を出す / オフで予約を消す', () => {
  const env = newEnv({ timeLimit: 4000 });
  env.ctx.lt_autoOn();
  const msg = env.alerts[env.alerts.length - 2].msg; // 確認の画面(最後は「オンにしました」)
  assert.ok(/小さな試験」がまだ/.test(msg) && /手動で1回も/.test(msg), msg);
  assert.strictEqual(env.props.LT_AUTO, '1');
  env.ctx.lt_diagStart();
  assert.strictEqual(env.triggers.length, 1, 'オンなら予約を作る');
  env.ctx.lt_autoOff();
  assert.strictEqual(env.triggers.length, 0, 'オフで予約を消す');
  assert.strictEqual(env.props.LT_AUTO, undefined);

  env.ctx.lt_smallTest();
  env.ctx.lt_autoOn();
  const msg2 = env.alerts[env.alerts.length - 2].msg;
  assert.ok(!/注意:/.test(msg2), '小さな試験と手動1回のあとは注意なし: ' + msg2);
  env.answer = 'NO';
  env.ctx.lt_autoOff();
  env.ctx.lt_autoOn();
  assert.strictEqual(env.props.LT_AUTO, undefined, '「いいえ」ならオンにしない');
  noProblems(env);
});

test('読み取りの失敗: 行数を減らして読み直す / 読めない行は記録して先へ進む', () => {
  const env = newEnv({
    tweakTest: (spec) => {
      spec.sheets.find((x) => x.name === '編打刻').fail = (kind, r, nr) => kind === 'v' && nr > 20;
      spec.sheets.find((x) => x.name === '002').fail = (kind, r, nr) => kind === 'f' && r <= 55 && r + nr - 1 >= 55;
    },
  });
  env.ctx.lt_diagStart();
  assert.strictEqual(state(env).phase, 'done');
  const S = env.ctx.S;
  const d = rowOf(env, S.DETAIL, 1, '編打刻');
  assert.strictEqual(d['値を読めなかった行'], '');
  assert.strictEqual(d['定数の最終行'], 400);
  assert.ok(/読み直す/.test(d['メモ']));
  const p = rowOf(env, S.DETAIL, 1, '002');
  assert.ok(/55/.test(p['調べられなかった行']), p['調べられなかった行']);
  assert.ok(rowOf(env, S.SUMMARY, 0, '調べられなかった行があるシート')['値'] === '1 枚');
  noProblems(env);
});

test('シートが多くても状態は 9KB 以内(スクリプト プロパティの上限)', () => {
  const env = newEnv({
    tweakTest: (spec) => {
      for (let i = 0; i < 160; i++) spec.sheets.push({ name: String(100 + i), id: 1900000000 + i * 7919, maxRows: 5, maxCols: 3, f: {}, v: {} });
    },
    timeLimit: 3000,
  });
  env.ctx.lt_diagStart();
  assert.ok(env.props.LT_STATE.length < 9000, String(env.props.LT_STATE.length));
  noProblems(env);
});

// ---------------- 4. 比較基準 ----------------
test('比較基準: 候補 → 記録(本番との違いも) → 比較(変化なし → 1セル変化)', () => {
  const env = newEnv();
  const S = env.ctx.S;
  env.ctx.lt_baselineCandidates();
  const t = env.rows(S.TARGETS);
  assert.deepStrictEqual(t.slice(1).map((r) => r[1]), ['001', '集計']);
  env.ctx.lt_baselineRecord();
  const b1 = rowOf(env, S.BASE, 3, '001');
  const b2 = rowOf(env, S.BASE, 3, '集計');
  assert.strictEqual(b1['本番との違い(セル)'], 0);
  assert.strictEqual(b2['本番との違い(セル)'], 1);
  assert.ok(/C2/.test(b2['本番との違い(例)']));
  assert.strictEqual(b1['範囲'], 'A1:T55');
  env.ctx.lt_baselineCompare();
  let res = env.rows(S.RESULT);
  assert.deepStrictEqual(res.filter((r) => r[4] === '同じ').map((r) => r[1]), ['001', '集計']);
  env.testSpec.sheets.find((x) => x.name === '001').v[12][5] = 777;
  env.ctx.lt_baselineCompare();
  res = env.rows(S.RESULT);
  const diff = res.find((r) => r[2] === 'E12');
  assert.ok(diff, '変えたセルが出る');
  assert.strictEqual(diff[6], '777');
  assert.strictEqual(res.filter((r) => r[4] === '違いあり').length, 1);
  noProblems(env);
});

console.log('\nすべての試験に合格: ' + passed + ' 件');
