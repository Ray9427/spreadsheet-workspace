// Google Apps Script の一部(このツールが使う範囲)を Node で真似る試験用の部品。
// 対象ファイル(本番・保管用・試験用コピー)の模擬は、書き込みや未知のメソッドを呼ぶと例外を出す。
'use strict';
const fs = require('fs');
const vm = require('vm');

const COST = { call: 20, perCell: 0.05 }; // 模擬の時計を進める量(ミリ秒)

function strict(obj, label, env) {
  return new Proxy(obj, {
    get(t, p) {
      if (typeof p === 'symbol' || p === 'then' || p === 'toJSON' || p === 'inspect') return t[p];
      if (p in t) return t[p];
      env.violations.push(label + '.' + String(p));
      throw new Error('MOCK: 用意していないメソッド ' + label + '.' + String(p));
    },
  });
}

function colLetter(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
function colIndex(l) {
  let n = 0;
  for (const ch of l) n = n * 26 + ch.charCodeAt(0) - 64;
  return n;
}
function parseA1(a1) {
  const m = String(a1).match(/^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/);
  if (!m) throw new Error('MOCK: A1 形式を読めない ' + a1);
  const r1 = Number(m[2]);
  const c1 = colIndex(m[1]);
  const r2 = m[4] ? Number(m[4]) : r1;
  const c2 = m[3] ? colIndex(m[3]) : c1;
  return [r1, c1, r2 - r1 + 1, c2 - c1 + 1];
}
function a1(r, c, nr, nc) {
  const s = colLetter(c) + r;
  return nr === 1 && nc === 1 ? s : s + ':' + colLetter(c + nc - 1) + (r + nr - 1);
}

function createEnv(opts) {
  const env = {
    now: Date.UTC(2026, 9, 15, 1, 0, 0),
    violations: [],
    formulaWrites: [],
    errorWrites: [],
    logs: [],
    alerts: [],
    prompts: (opts && opts.prompts) || [],
    answer: 'YES',
    props: {},
    triggers: [],
    triggerSeq: 0,
    lockHeld: false,
    failToolWriteAt: -1,
    toolWrites: 0,
    targets: {},
    readCalls: 0,
  };
  const tick = (cells) => (env.now += COST.call + (cells || 0) * COST.perCell);

  // ---------- このツールのファイル(書き込める) ----------
  function parseInput(v) {
    if (typeof v !== 'string') return v;
    if (v.length > 50000) throw new Error('MOCK: 1セル 50000 文字を超えた');
    if (v.charAt(0) === "'") return v.slice(1);
    if (v.charAt(0) === '=') {
      env.formulaWrites.push(v);
      return v;
    }
    if (v.charAt(0) === '#') env.errorWrites.push(v);
    if (v.trim() !== '' && !isNaN(Number(v))) return Number(v);
    if (/^(true|false)$/i.test(v)) return v.toUpperCase() === 'TRUE';
    return v;
  }
  function toolSheet(ss, name, id) {
    const sh = { name, id, hidden: false, grid: [], maxRows: 1000, maxCols: 26, frozen: 0 };
    const cell = (r, c) => (sh.grid[r - 1] && sh.grid[r - 1][c - 1] !== undefined ? sh.grid[r - 1][c - 1] : '');
    const lastRow = () => {
      for (let r = sh.grid.length; r >= 1; r--) if ((sh.grid[r - 1] || []).some((v) => v !== '' && v !== undefined)) return r;
      return 0;
    };
    const lastCol = () => sh.grid.reduce((m, row) => {
      for (let c = (row || []).length; c >= 1; c--) if (row[c - 1] !== '' && row[c - 1] !== undefined) return Math.max(m, c);
      return m;
    }, 0);
    function range(r, c, nr, nc) {
      if (r < 1 || c < 1 || nr < 1 || nc < 1 || r + nr - 1 > sh.maxRows || c + nc - 1 > sh.maxCols) {
        throw new Error('MOCK: 範囲がシートの外 ' + name + ' ' + [r, c, nr, nc].join(','));
      }
      const api = {
        getValues() {
          tick(nr * nc);
          const out = [];
          for (let i = 0; i < nr; i++) {
            const row = [];
            for (let j = 0; j < nc; j++) row.push(cell(r + i, c + j));
            out.push(row);
          }
          return out;
        },
        setValues(vals) {
          tick(nr * nc);
          env.toolWrites++;
          if (env.failToolWriteAt === env.toolWrites) throw new Error('MOCK: 書き込みの途中で失敗(試験用の障害)');
          if (vals.length !== nr || vals.some((x) => x.length !== nc)) throw new Error('MOCK: setValues の大きさが違う');
          for (let i = 0; i < nr; i++) {
            sh.grid[r - 1 + i] = sh.grid[r - 1 + i] || [];
            for (let j = 0; j < nc; j++) sh.grid[r - 1 + i][c - 1 + j] = parseInput(vals[i][j]);
          }
          return api;
        },
        setValue(v) {
          return api.setValues([[v]]);
        },
        clearContent() {
          for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) if (sh.grid[r - 1 + i]) sh.grid[r - 1 + i][c - 1 + j] = '';
          return api;
        },
        setFontWeight: () => api,
        setBackground: () => api,
        insertCheckboxes: () => api,
        getA1Notation: () => a1(r, c, nr, nc),
      };
      return strict(api, 'ToolRange', env);
    }
    const api = {
      getName: () => sh.name,
      getSheetId: () => sh.id,
      getRange: (r, c, nr, nc) => (typeof r === 'string' ? range(...parseA1(r)) : range(r, c, nr || 1, nc || 1)),
      getLastRow: lastRow,
      getLastColumn: lastCol,
      getMaxRows: () => sh.maxRows,
      getMaxColumns: () => sh.maxCols,
      insertRowsAfter: (after, n) => (sh.maxRows += n),
      insertColumnsAfter: (after, n) => (sh.maxCols += n),
      clearContents: () => (sh.grid = []),
      setFrozenRows: (n) => (sh.frozen = n),
      hideSheet: () => (sh.hidden = true),
      activate: () => api,
      _grid: sh,
    };
    return strict(api, 'ToolSheet(' + name + ')', env);
  }
  const tool = { id: 'TOOL_FILE_ID_000000000000000000', name: '軽量化ツール_診断', sheets: [] };
  const toolSs = strict(
    {
      getId: () => tool.id,
      getName: () => tool.name,
      getSheets: () => tool.sheets.slice(),
      getSheetByName: (n) => tool.sheets.find((s) => s.getName() === n) || null,
      insertSheet: (n, idx) => {
        if (tool.sheets.some((s) => s.getName() === n)) throw new Error('MOCK: 同じ名前のシートがある ' + n);
        const s = toolSheet(toolSs, n, 1000 + tool.sheets.length);
        tool.sheets.splice(idx === undefined ? tool.sheets.length : idx, 0, s);
        return s;
      },
      toast: () => {},
      getSpreadsheetTimeZone: () => 'Asia/Tokyo',
    },
    'ToolSpreadsheet',
    env
  );
  tool.sheets.push(toolSheet(toolSs, 'シート1', 0));

  // ---------- 対象ファイル(読むだけ) ----------
  function targetSs(spec) {
    const sheets = spec.sheets.map((s) => targetSheet(s, spec));
    const ss = {
      getId: () => spec.id,
      getName: () => spec.name,
      getSheets: () => (tick(), sheets.slice()),
      getSheetByName: (n) => sheets.find((s) => s.getName() === n) || null,
      getNamedRanges: () => (spec.named || []).map((n) => strict({ getName: () => n.name, getRange: () => sheets.find((s) => s.getName() === n.sheet).getRange(n.a1) }, 'NamedRange', env)),
      getProtections: (type) => (spec.protections || []).filter((p) => p.type === type).map((p) => {
        const sh = sheets.find((s) => s.getName() === p.sheet);
        return strict({ getRange: () => (p.a1 ? sh.getRange(p.a1) : sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns())) }, 'Protection', env);
      }),
      getRecalculationInterval: () => 'ON_CHANGE',
      isIterativeCalculationEnabled: () => false,
      getSpreadsheetLocale: () => 'ja_JP',
      getSpreadsheetTimeZone: () => 'Asia/Tokyo',
    };
    return strict(ss, 'Target(' + spec.name + ')', env);
  }
  function targetSheet(s, spec) {
    const isGrid = (s.type || 'GRID') === 'GRID';
    const rangeOf = (r, c, nr, nc) => {
      if (!isGrid) throw new Error('MOCK: 表ではないシートの getRange');
      if (r < 1 || c < 1 || nr < 1 || nc < 1 || r + nr - 1 > s.maxRows || c + nc - 1 > s.maxCols) throw new Error('MOCK: 範囲がシートの外 ' + s.name);
      const get = (kind) => {
        tick(nr * nc);
        env.readCalls++;
        if (s.fail && s.fail(kind, r, nr)) throw new Error('Service Spreadsheets timed out while accessing document(試験用の障害)');
        const out = [];
        for (let i = 0; i < nr; i++) {
          const row = [];
          for (let j = 0; j < nc; j++) {
            const rr = r + i;
            const cc = c + j;
            if (kind === 'f') row.push((s.f[rr] && s.f[rr][cc]) || '');
            else if (kind === 'v') row.push(s.v[rr] && s.v[rr][cc] !== undefined ? s.v[rr][cc] : '');
            else if (kind === 'd') {
              const v = s.v[rr] && s.v[rr][cc] !== undefined ? s.v[rr][cc] : '';
              row.push(Object.prototype.toString.call(v) === '[object Date]' ? v.toISOString().slice(0, 10) : String(v));
            } else row.push(s.dv && s.dv[rr + ',' + cc] ? strict({}, 'DataValidation', env) : null);
          }
          out.push(row);
        }
        return out;
      };
      return strict(
        {
          getFormulas: () => get('f'),
          getValues: () => get('v'),
          getDisplayValues: () => get('d'),
          getDataValidations: () => get('dv'),
          getMergedRanges: () => (s.merges || []).map((m) => sheetApi.getRange(m)),
          getA1Notation: () => a1(r, c, nr, nc),
          getRow: () => r,
          getColumn: () => c,
          getNumRows: () => nr,
          getNumColumns: () => nc,
          getLastRow: () => r + nr - 1,
          getLastColumn: () => c + nc - 1,
          getSheet: () => sheetApi,
        },
        'TargetRange(' + s.name + ')',
        env
      );
    };
    const lastRowCol = () => {
      let lr = 0;
      let lc = 0;
      [s.f, s.v].forEach((g) => Object.keys(g).forEach((r) => Object.keys(g[r]).forEach((c) => {
        const val = g[r][c];
        if (val !== '' && val !== undefined && val !== null) {
          lr = Math.max(lr, Number(r));
          lc = Math.max(lc, Number(c));
        }
      })));
      return [lr, lc];
    };
    const sheetApi = strict(
      {
        getName: () => s.name,
        getSheetId: () => s.id,
        isSheetHidden: () => !!s.hidden,
        getType: () => s.type || 'GRID',
        getMaxRows: () => (isGrid ? s.maxRows : 0),
        getMaxColumns: () => (isGrid ? s.maxCols : 0),
        getFrozenRows: () => s.frozenRows || 0,
        getFrozenColumns: () => s.frozenCols || 0,
        getLastRow: () => lastRowCol()[0],
        getLastColumn: () => lastRowCol()[1],
        getRange: (r, c, nr, nc) => (typeof r === 'string' ? rangeOf(...parseA1(r)) : rangeOf(r, c, nr || 1, nc || 1)),
        getConditionalFormatRules: () => (s.cf || []).map((rule) => strict({
          getRanges: () => rule.ranges.map((x) => sheetApi.getRange(x)),
          getBooleanCondition: () => strict({ getCriteriaType: () => rule.type, getCriteriaValues: () => [rule.formula] }, 'BooleanCondition', env),
        }, 'ConditionalFormatRule', env)),
        getCharts: () => (s.charts || []).map((ch) => strict({
          getRanges: () => ch.sources.map((x) => {
            const t = spec.sheets.find((q) => q.name === x.sheet);
            return strict({ getSheet: () => strict({ getName: () => t.name }, 'Sheet', env) }, 'Range', env);
          }),
          getContainerInfo: () => strict({ getAnchorRow: () => ch.row || 1, getAnchorColumn: () => ch.col || 1 }, 'ContainerInfo', env),
        }, 'Chart', env)),
        getPivotTables: () => [],
        getFilter: () => null,
        getBandings: () => [],
        getSlicers: () => [],
        getImages: () => [],
        getDrawings: () => (s.drawings || []).map((d) => strict({
          getOnAction: () => d.action || '',
          getContainerInfo: () => strict({ getAnchorRow: () => d.row || 1, getAnchorColumn: () => d.col || 1 }, 'ContainerInfo', env),
        }, 'Drawing', env)),
      },
      'TargetSheet(' + s.name + ')',
      env
    );
    return sheetApi;
  }

  // ---------- GAS のサービス ----------
  const ui = {
    ButtonSet: { OK: 'OK', OK_CANCEL: 'OK_CANCEL', YES_NO: 'YES_NO' },
    Button: { OK: 'OK', CANCEL: 'CANCEL', YES: 'YES', NO: 'NO' },
    alert: (title, msg, buttons) => {
      env.alerts.push({ title, msg: msg === undefined ? title : msg });
      return buttons === 'YES_NO' ? env.answer : 'OK';
    },
    prompt: (title, msg) => {
      if (!env.prompts.length) throw new Error('MOCK: 用意していない入力欄 ' + title);
      const text = env.prompts.shift();
      return { getSelectedButton: () => (text === null ? 'CANCEL' : 'OK'), getResponseText: () => text || '' };
    },
    createMenu: () => {
      const m = { addItem: () => m, addSeparator: () => m, addToUi: () => {} };
      return m;
    },
  };
  env.uiAvailable = true;
  const SpreadsheetApp = {
    getActive: () => toolSs,
    getActiveSpreadsheet: () => toolSs,
    openById: (id) => {
      tick(1000);
      if (!env.targets[id]) throw new Error('MOCK: ファイルが見つからない ' + id);
      return env.targets[id];
    },
    getUi: () => {
      if (!env.uiAvailable) throw new Error('Cannot call SpreadsheetApp.getUi() from this context.');
      return ui;
    },
    SheetType: { GRID: 'GRID', OBJECT: 'OBJECT', DATASOURCE: 'DATASOURCE' },
    BooleanCriteria: { CUSTOM_FORMULA: 'CUSTOM_FORMULA' },
    ProtectionType: { SHEET: 'SHEET', RANGE: 'RANGE' },
  };
  const props = {
    getProperty: (k) => (k in env.props ? env.props[k] : null),
    setProperty: (k, v) => {
      if (String(v).length > 9000) throw new Error('MOCK: プロパティの値が 9KB を超えた ' + k + ' ' + String(v).length);
      env.props[k] = String(v);
      return props;
    },
    setProperties: (o) => {
      Object.keys(o).forEach((k) => props.setProperty(k, o[k]));
      return props;
    },
    deleteProperty: (k) => {
      delete env.props[k];
      return props;
    },
  };
  const ScriptApp = {
    newTrigger: (fn) => {
      const b = {
        timeBased: () => b,
        after: (ms) => ((b.at = env.now + ms), b),
        create: () => {
          const t = { id: ++env.triggerSeq, fn, at: b.at };
          env.triggers.push(t);
          return t;
        },
      };
      return b;
    },
    getProjectTriggers: () => env.triggers.map((t) => ({ getHandlerFunction: () => t.fn, getUniqueId: () => String(t.id), _t: t })),
    deleteTrigger: (t) => (env.triggers = env.triggers.filter((x) => x.id !== t._t.id)),
  };
  const pad = (n) => String(n).padStart(2, '0');
  const Utilities = {
    formatDate: (d, tz, fmt) => {
      const j = new Date(d.getTime() + 9 * 3600 * 1000);
      return fmt
        .replace('yyyy', j.getUTCFullYear())
        .replace('MM', pad(j.getUTCMonth() + 1))
        .replace('dd', pad(j.getUTCDate()))
        .replace('HH', pad(j.getUTCHours()))
        .replace('mm', pad(j.getUTCMinutes()))
        .replace('ss', pad(j.getUTCSeconds()));
    },
  };
  const RealDate = Date;
  class FakeDate extends RealDate {
    constructor(...a) {
      if (a.length === 0) super(env.now);
      else super(...a);
    }
    static now() {
      return env.now;
    }
  }

  const ctx = {
    SpreadsheetApp,
    PropertiesService: { getScriptProperties: () => props },
    LockService: {
      getScriptLock: () => ({
        tryLock: () => (env.lockHeld ? false : (env.lockHeld = true)),
        releaseLock: () => (env.lockHeld = false),
      }),
    },
    ScriptApp,
    Utilities,
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
    console: { log: (s) => env.logs.push(String(s)) },
    Date: FakeDate,
    Math,
    JSON,
    Object,
    Number,
    String,
    Array,
    RegExp,
    Map,
    Set,
    Error,
    isFinite,
    isNaN,
  };
  vm.createContext(ctx);

  env.ctx = ctx;
  env.tool = tool;
  env.addTarget = (spec) => (env.targets[spec.id] = targetSs(spec));
  env.load = (path) => vm.runInContext(fs.readFileSync(path, 'utf8') + '\nthis.LT = LT; this.S = S;', ctx, { filename: path });
  env.sheet = (name) => {
    const s = tool.sheets.find((x) => x.getName() === name);
    return s ? s._grid : null;
  };
  env.rows = (name) => {
    const s = env.sheet(name);
    return s ? s.grid.filter((r) => r && r.some((v) => v !== '' && v !== undefined)) : [];
  };
  env.fireDueTriggers = () => {
    const due = env.triggers.slice().sort((a, b) => a.at - b.at)[0];
    if (!due) return false;
    env.now = Math.max(env.now, due.at);
    env.uiAvailable = false;
    try {
      ctx[due.fn]();
    } finally {
      env.uiAvailable = true;
    }
    return true;
  };
  return env;
}

module.exports = { createEnv, colLetter, a1 };
