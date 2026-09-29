// Локальний мок Apps Script: виконує apps-script/Code.gs з фейковими сервісами Google над фікстурою.
// Самостійний запуск: `node test/mock-apps-script.mjs` → http://localhost:8787/#key=test-key
import http from 'node:http';
import vm from 'node:vm';
import { randomUUID } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { pathToFileURL } from 'node:url';

const root = new URL('../', import.meta.url);
export const KEY = 'test-key';
const isDate = (v) => Object.prototype.toString.call(v) === '[object Date]';
const d = (y, m, day) => new Date(y, m - 1, day);
const STATUS_LIST = ['Not started', 'In progress', 'Hold', 'Review', 'Done'];

export function makeFixture() {
  return {
    locks: 0,
    tabs: [
      {
        name: 'Sep 26',
        validation: { col: 3, values: STATUS_LIST },
        rows: [
          ['OVERALL DESIGN SUPPORT'],
          ['Задача', 'Проєкт', 'Годин', 'Статус', 'Відповідальний', 'Дедлайн', 'Дата виконання', 'Коментар', 'Посилання ТЗ', 'Посилання файл'],
          ['Банер для сайту', 'ATE', 3, 'In progress ', 'Олена', d(2026, 10, 3), '', 'Дві версії', 'https://example.com/tz', 'https://example.com/file'],
          ['Іконки для застосунку', 'Mobile', 6, 'Not started', 'Андрій', d(2026, 10, 10), '', '', '', ''],
          ['Логотип', 'Brand', 2.5, 'Not started', 'Олена', '', '', '', '', ''],
          ['', '', '', '', '', '', '', '', '', ''],
          ['Презентація для інвесторів з довгою назвою, яка займає кілька рядків', 'ATE', 8, 'Done', 'Марія', '', d(2026, 9, 12), '', 'https://example.com/tz2', ''],
          ['Пости для соцмереж', 'SMM', 1.5, 'Not started', '', d(2026, 9, 30), '', '', '', ''],
        ],
      },
      {
        name: 'Aug 26',
        validation: { col: 1, values: STATUS_LIST },
        rows: [
          ['OVERALL DESIGN SUPPORT'],
          ['Проєкт', 'Статус', 'Задача', 'Години', 'Дата', 'Відповідальний', 'Коментар'],
          ['ATE', 'Done', 'Серпневий лендінг', 4, d(2026, 8, 20), 'Олена', ''],
          ['Brand', 'Done', 'Гайдлайн кольорів', '4,5', d(2026, 8, 25), 'Андрій', 'старий формат'],
          ['', '', 'Overtime extimate (x1.5) = 3', 2, '', '', ''], // службовий рядок, як у реальній таблиці
        ],
      },
      { name: 'Зведена таблиця', rows: [['Задача', 'IBAN'], ['Оплата', 'UA000000000000000000000000000']] },
      {
        name: 'September 25',
        rows: [
          ['OVERALL DESIGN SUPPORT'],
          ['Задача', '', 'Годин', 'Статус ', 'Проєкт'], // як у реальній вкладці: «Задача» над №, назви — без заголовка
          [1, 'Старий макет', 2, 'On hold', 'ATE'],
          [2, 'Архівна задача', 1, 'Done ', 'Brand'],
          ['', '', '', '', ''],
          ['', '', 3, '', ''], // підсумок без назви — не задача
        ],
      },
    ],
  };
}

// Спрощено: формат клітинки впливає лише на те, чи показувати час у даті.
function display(v, fmt = '') {
  const p2 = (n) => String(n).padStart(2, '0');
  if (isDate(v)) return `${p2(v.getDate())}.${p2(v.getMonth() + 1)}.${v.getFullYear()}` + (/h/i.test(fmt) ? ` ${p2(v.getHours())}:${p2(v.getMinutes())}` : '');
  if (typeof v === 'number') return String(v).replace('.', ',');
  return v == null ? '' : String(v);
}
const empty = (v) => v === '' || v == null;

function fakeServices(fx) {
  const props = { ADMIN_KEY: KEY }, cache = new Map(); // ponytail: кеш без TTL — у тестах він не потрібен
  const sheet = (tab) => {
    const lastRow = () => tab.rows.reduce((last, r, i) => (r.some((v) => !empty(v)) ? i + 1 : last), 0);
    const lastCol = () => Math.max(0, ...tab.rows.map((r) => r.reduce((l, v, i) => (empty(v) ? l : i + 1), 0)));
    const range = (r, c, nr = 1, nc = 1) => {
      tab.formats ??= {};
      const grid = (fn) => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => fn((tab.rows[r - 1 + i] || [])[c - 1 + j] ?? '', tab.formats[`${r + i},${c + j}`])));
      return {
        getValues: () => grid((v) => v),
        getDisplayValues: () => grid(display),
        getValue: () => grid((v) => v)[0][0],
        getDisplayValue: () => grid(display)[0][0],
        getNumberFormat: () => tab.formats[`${r},${c}`] ?? 'dd.mm.yyyy',
        setNumberFormat(f) { tab.formats[`${r},${c}`] = f; return this; },
        setValue(v) {
          while (tab.rows.length < r) tab.rows.push([]);
          tab.rows[r - 1][c - 1] = v;
          return this;
        },
        getDataValidation: () => (tab.validation && tab.validation.col === c - 1
          ? { getCriteriaType: () => 'VALUE_IN_LIST', getCriteriaValues: () => [tab.validation.values, true] }
          : null),
      };
    };
    return {
      getName: () => tab.name,
      getLastRow: lastRow,
      getLastColumn: lastCol,
      getRange: range,
      insertRowAfter: (r) => tab.rows.splice(r, 0, []),
    };
  };
  const ss = {
    getSheets: () => fx.tabs.map(sheet),
    getSheetByName: (n) => { const t = fx.tabs.find((t) => t.name === n); return t ? sheet(t) : null; },
  };
  return {
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss,
      getActive: () => ss,
      flush() {},
      DataValidationCriteria: { VALUE_IN_LIST: 'VALUE_IN_LIST', VALUE_IN_RANGE: 'VALUE_IN_RANGE' },
    },
    LockService: { getScriptLock: () => ({ waitLock: () => fx.locks++, releaseLock() {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => props[k] ?? null, setProperty: (k, v) => { props[k] = v; } }) },
    CacheService: { getScriptCache: () => ({ get: (k) => cache.get(k) ?? null, put: (k, v) => { cache.set(k, v); } }) },
    ContentService: {
      MimeType: { JSON: 'application/json' },
      createTextOutput: (s) => ({ setMimeType() { return this; }, getContent: () => s }),
    },
    Utilities: { getUuid: randomUUID },
    Logger: { log: console.log },
  };
}

export function loadScript(fx) {
  const ctx = vm.createContext(fakeServices(fx));
  vm.runInContext(readFileSync(new URL('apps-script/Code.gs', root), 'utf8'), ctx, { filename: 'Code.gs' });
  return ctx;
}

/** Сторінка — на http://localhost:PORT/, API — на http://127.0.0.1:PORT/exec (інше походження, як у проді). */
export function startMock({ port = 0, fixture = makeFixture() } = {}) {
  const gs = loadScript(fixture);
  const server = http.createServer(async (req, res) => {
    const url = new URL(req.url, 'http://x');
    if (url.pathname === '/exec' && (req.method === 'GET' || req.method === 'POST')) {
      let body = '';
      for await (const chunk of req) body += chunk;
      const out = req.method === 'GET'
        ? vm.runInContext('doGet', gs)({ parameter: Object.fromEntries(url.searchParams) })
        : vm.runInContext('doPost', gs)({ postData: { contents: body } });
      res.writeHead(200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
      return res.end(out.getContent());
    }
    if (url.pathname === '/' && req.method === 'GET') {
      const api = `http://127.0.0.1:${server.address().port}/exec`;
      const html = readFileSync(new URL('index.html', root), 'utf8').replace(/const API_URL = "[^"]*"/, `const API_URL = "${api}"`);
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      return res.end(html);
    }
    res.writeHead(req.method === 'OPTIONS' ? 405 : 404).end(); // жодного CORS preflight
  });
  // onEdit() — імітує ручну правку таблиці (простий тригер скидає кеш).
  const onEdit = () => vm.runInContext('onEdit', gs)();
  return new Promise((ok) => server.listen(port, () => ok({ server, fixture, onEdit, port: server.address().port })));
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  const { port } = await startMock({ port: 8787 });
  console.log(`Мок: http://localhost:${port}/#key=${KEY}`);
}
