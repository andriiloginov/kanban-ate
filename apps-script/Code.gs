/**
 * Канбан API для таблиці «Design 2025». Інструкція — README.md.
 * GET  ?action=tabs|list|all&tab=...&key=...   (all — усі вкладки-місяці одразу, для статистики)
 * POST {action: "update"|"create", key, ...}  (Content-Type: text/plain)
 */

const HIDDEN_TABS = ['зведена таблиця']; // платіжні реквізити — ніколи не читати
const FIELDS = {
  title: ['задача'],
  hours: ['годин', 'години'],
  project: ['проєкт'],
  status: ['статус'],
  assignee: ['відповідальний'],
  doneDate: ['дата', 'дата виконання'],
  deadline: ['дедлайн'],
  comment: ['коментар'],
  briefLink: ['посилання тз'],
  fileLink: ['посилання файл'],
};
const EDITABLE = ['title', 'hours', 'project', 'status', 'assignee', 'comment', 'deadline', 'doneDate'];
const DEFAULT_STATUSES = ['Not started', 'In progress', 'Done'];

/** Запустіть один раз вручну: згенерує ключ доступу і виведе його в журнал. */
function setup() {
  const key = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('ADMIN_KEY', key);
  Logger.log('ADMIN_KEY: ' + key);
}

function doGet(e) {
  return respond(e && e.parameter, ['tabs', 'list', 'all']);
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json({ error: 'bad_request' });
  }
  return respond(body, ['update', 'create']);
}

function respond(p, allowed) {
  p = p || {};
  let res;
  try {
    const key = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
    if (!key || p.key !== key) res = { error: 'unauthorized' };
    else if (allowed.indexOf(p.action) < 0) res = { error: 'unknown_action' };
    else res = ACTIONS[p.action](p);
  } catch (err) {
    res = { error: 'server_error', message: String((err && err.message) || err) };
  }
  return json(res);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

const ACTIONS = {
  tabs: () => ({ tabs: boardSheets().map((s) => ({ name: s.getName() })) }),

  all: () => ({ tabs: boardSheets().map((s) => ACTIONS.list({ tab: s.getName() })) }),

  list: (p) => {
    const b = getBoard(p.tab);
    if (!b) return { error: 'tab_not_found' };
    const { sheet, header } = b;
    const first = header.row + 1, last = sheet.getLastRow(), items = [];
    if (last >= first) {
      const range = sheet.getRange(first, 1, last - first + 1, header.width);
      const values = range.getValues(), display = range.getDisplayValues();
      values.forEach((v, i) => {
        const item = toItem(header.cols, v, display[i], first + i);
        if (item) items.push(item);
      });
    }
    return {
      tab: sheet.getName(),
      columns: Object.keys(FIELDS).filter((f) => header.cols[f] != null),
      statuses: rawStatuses(sheet, header).map((s) => s.trim()),
      items,
    };
  },

  update: (p) => withLock(() => {
    const b = getBoard(p.tab);
    if (!b) return { error: 'tab_not_found' };
    const { sheet, header } = b, cols = header.cols, row = Number(p.row);
    if (!(row > header.row && row <= sheet.getLastRow())) return { error: 'conflict' };
    const current = String(sheet.getRange(row, cols.title + 1).getDisplayValue()).trim();
    if (current !== String(p.expectedTitle == null ? '' : p.expectedTitle).trim()) return { error: 'conflict' };

    const patch = Object.assign({}, p.patch);
    if (isDone(patch.status) && cols.doneDate != null && !('doneDate' in patch) &&
        !String(sheet.getRange(row, cols.doneDate + 1).getDisplayValue()).trim()) {
      patch.doneDate = todayIso();
    }
    const writes = cellWrites(sheet, header, patch);
    if (writes.error) return writes;
    writeCells(sheet, row, writes);
    SpreadsheetApp.flush();
    return { item: readRow(sheet, header, row) };
  }),

  create: (p) => withLock(() => {
    const b = getBoard(p.tab);
    if (!b) return { error: 'tab_not_found' };
    const { sheet, header } = b, cols = header.cols, last = sheet.getLastRow();
    const item = Object.assign({}, p.item);
    if (!String(item.title == null ? '' : item.title).trim()) return { error: 'title_required' };
    if (!item.status) item.status = rawStatuses(sheet, header)[0];

    // Рядок одразу після останньої задачі; якщо він чимось зайнятий (напр. підсумок) — вставляємо новий.
    let row = header.row + 1, maxNum = 0;
    if (last > header.row) {
      const values = sheet.getRange(header.row + 1, 1, last - header.row, header.width).getValues();
      values.forEach((v, i) => {
        if (String(v[cols.title]).trim()) row = header.row + 2 + i;
        if (cols.num != null) maxNum = Math.max(maxNum, toNumber(v[cols.num]) || 0);
      });
      const next = values[row - header.row - 1];
      if (next && next.some((x) => String(x).trim())) sheet.insertRowAfter(row - 1);
    }
    const patch = {};
    EDITABLE.forEach((f) => { if (item[f] != null && item[f] !== '') patch[f] = item[f]; });
    if (isDone(patch.status) && !patch.doneDate) patch.doneDate = todayIso();
    const writes = cellWrites(sheet, header, patch);
    if (writes.error) return writes;
    if (cols.num != null) writes.push([cols.num, maxNum + 1]);
    writeCells(sheet, row, writes);
    SpreadsheetApp.flush();
    return { item: readRow(sheet, header, row) };
  }),
};

function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

const norm = (s) => String(s == null ? '' : s).trim().toLowerCase();
const isHidden = (name) => HIDDEN_TABS.indexOf(norm(name)) >= 0;
const isDone = (status) => norm(status) === 'done';
const isDate = (v) => Object.prototype.toString.call(v) === '[object Date]';
const pad = (n) => (n < 10 ? '0' : '') + n;

function boardSheets() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().filter((s) => !isHidden(s.getName()) && readHeader(s));
}

function getBoard(tab) {
  if (!tab || isHidden(tab)) return null;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(String(tab));
  const header = sheet && readHeader(sheet);
  return header ? { sheet, header } : null;
}

/** Рядок заголовків — перший із рядків 1..5, де є клітинка «Задача». */
function readHeader(sheet) {
  const width = sheet.getLastColumn(), n = Math.min(5, sheet.getLastRow());
  if (!width || !n) return null;
  const rows = sheet.getRange(1, 1, n, width).getDisplayValues();
  for (let r = 0; r < n; r++) {
    const cells = rows[r].map(norm);
    if (cells.indexOf('задача') < 0) continue;
    const cols = {};
    cells.forEach((c, i) => {
      if (c === '№') cols.num = i;
      Object.keys(FIELDS).forEach((f) => {
        if (cols[f] == null && FIELDS[f].indexOf(c) >= 0) cols[f] = i;
      });
    });
    // «Задача», об'єднана з порожньою клітинкою праворуч (як у «September 25»): ліворуч №, праворуч назва.
    if (cols.num == null && cols.title + 1 < width && !cells[cols.title + 1]) {
      cols.num = cols.title;
      cols.title++;
    }
    return { row: r + 1, cols, width };
  }
  return null;
}

/** Допустимі статуси з випадного списку колонки «Статус» (як є, з пробілами), інакше стандартні. */
function rawStatuses(sheet, header) {
  if (header.cols.status == null) return DEFAULT_STATUSES;
  const rule = sheet.getRange(header.row + 1, header.cols.status + 1).getDataValidation();
  if (rule) {
    const C = SpreadsheetApp.DataValidationCriteria, type = rule.getCriteriaType(), args = rule.getCriteriaValues();
    let list = null;
    if (type === C.VALUE_IN_LIST) list = args[0];
    else if (type === C.VALUE_IN_RANGE) list = [].concat.apply([], args[0].getValues());
    list = (list || []).map(String).filter((s) => s.trim());
    if (list.length) return list;
  }
  return DEFAULT_STATUSES;
}

function toItem(cols, values, display, row) {
  const text = (f) => (cols[f] == null ? '' : String(display[cols[f]]).trim());
  const title = text('title');
  if (!title) return null;
  const item = { row, title, hours: cols.hours == null ? null : toNumber(values[cols.hours]) };
  ['project', 'status', 'assignee', 'doneDate', 'deadline', 'comment', 'briefLink', 'fileLink']
    .forEach((f) => { item[f] = text(f); });
  if (Number.isNaN(item.hours)) item.hours = null;
  item.deadlineIso = cols.deadline == null ? '' : toIso(values[cols.deadline]);
  item.doneDateIso = cols.doneDate == null ? '' : toIso(values[cols.doneDate]);
  return item;
}

function readRow(sheet, header, row) {
  const range = sheet.getRange(row, 1, 1, header.width);
  return toItem(header.cols, range.getValues()[0], range.getDisplayValues()[0], row);
}

/** Перетворює patch на [[колонка, значення, зЧасом?]]; значення вже в типах таблиці (число, Date, рядок). */
function cellWrites(sheet, header, patch) {
  const writes = [];
  for (const f of EDITABLE) {
    if (!(f in patch) || header.cols[f] == null) continue;
    let v = patch[f] == null ? '' : patch[f];
    if (f === 'hours') {
      v = toNumber(v);
      if (Number.isNaN(v)) return { error: 'bad_hours' };
      if (v == null) v = '';
    } else if (f === 'deadline' || f === 'doneDate') {
      const m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}))?$/);
      if (!m && String(v).trim()) return { error: 'bad_date' };
      v = m ? new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0)) : '';
      if (m && m[4]) { writes.push([header.cols[f], v, true]); continue; }
    } else if (f === 'status') {
      const raw = rawStatuses(sheet, header).filter((s) => norm(s) === norm(v))[0];
      v = raw || String(v).trim();
    } else {
      v = String(v).trim();
      if (f === 'title' && !v) return { error: 'title_required' };
    }
    writes.push([header.cols[f], v]);
  }
  return writes;
}

/** null — порожньо, NaN — не число. Приймає і 4.5, і "4,5". */
function toNumber(x) {
  if (typeof x === 'number') return x;
  const s = String(x == null ? '' : x).trim().replace(',', '.');
  return s === '' ? null : Number(s);
}

/** Записує клітинки; якщо дата з часом, а формат клітинки час не показує — додає час до формату. */
function writeCells(sheet, row, writes) {
  writes.forEach(([col, value, withTime]) => {
    const cell = sheet.getRange(row, col + 1);
    if (withTime) {
      const fmt = String(cell.getNumberFormat() || '');
      if (!/h/i.test(fmt)) cell.setNumberFormat(/[dmy]/i.test(fmt) ? fmt + ' HH:mm' : 'dd.mm.yyyy HH:mm');
    }
    cell.setValue(value);
  });
}

/** "yyyy-mm-dd", або "yyyy-mm-ddTHH:MM", якщо в даті є час. */
function toIso(v) {
  if (!isDate(v)) return '';
  const date = v.getFullYear() + '-' + pad(v.getMonth() + 1) + '-' + pad(v.getDate());
  return v.getHours() || v.getMinutes() ? date + 'T' + pad(v.getHours()) + ':' + pad(v.getMinutes()) : date;
}

function todayIso() {
  return toIso(new Date()).split('T')[0];
}
