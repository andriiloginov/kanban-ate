// npm test: мок Apps Script + Playwright (Chromium). Скриншоти — у test/screenshots/.
import assert from 'node:assert/strict';
import { mkdirSync } from 'node:fs';
import { chromium } from 'playwright';
import { startMock, KEY } from './mock-apps-script.mjs';

const { server, fixture, onEdit, port } = await startMock();
const API = `http://127.0.0.1:${port}/exec`;
const PAGE = `http://localhost:${port}/`;
const tab = (name) => fixture.tabs.find((t) => t.name === name);
const rowOf = (name, title) => tab(name).rows.find((r) => r.some((v) => String(v).trim() === title));
const isDate = (v) => Object.prototype.toString.call(v) === '[object Date]';
const get = (q) => fetch(`${API}?${new URLSearchParams(q)}`).then((r) => r.json());
const post = (body) => fetch(API, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(body) }).then((r) => r.json());
async function until(fn, ms = 5000) {
  for (const end = Date.now() + ms; Date.now() < end; await new Promise((r) => setTimeout(r, 50))) if (fn()) return;
  assert.fail(`не дочекались: ${fn}`);
}

let failed = 0;
async function test(name, fn) {
  try { await fn(); console.log(`✓ ${name}`); } catch (e) { failed++; console.log(`✗ ${name}\n  ${e.stack}`); }
}

const browser = await chromium.launch();
const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
const page = await context.newPage();
page.on('pageerror', (e) => { failed++; console.log('✗ pageerror', e); });
const card = (title) => page.locator('article', { hasText: title });
const column = (status) => page.locator(`[data-list="${status}"]`);

await test('unauthorized без ключа (GET і POST)', async () => {
  assert.deepEqual(await get({ action: 'tabs' }), { error: 'unauthorized' });
  assert.deepEqual(await get({ action: 'tabs', key: 'wrong' }), { error: 'unauthorized' });
  assert.deepEqual(await post({ action: 'create', tab: 'Sep 26', item: { title: 'x' } }), { error: 'unauthorized' });
});

await test('«Зведена таблиця» не видно через API', async () => {
  const { tabs } = await get({ action: 'tabs', key: KEY });
  assert.deepEqual(tabs.map((t) => t.name), ['Sep 26', 'Aug 26', 'September 25']);
  assert.deepEqual(await get({ action: 'list', tab: 'Зведена таблиця', key: KEY }), { error: 'tab_not_found' });
});

await test('list: колонки за назвою, числа, trim статусів', async () => {
  const aug = await get({ action: 'list', tab: 'Aug 26', key: KEY });
  const it = aug.items.find((i) => i.title === 'Гайдлайн кольорів');
  assert.equal(it.hours, 4.5);
  assert.equal(it.doneDate, '25.08.2026');
  assert.equal(it.status, 'Done');
  const sep = await get({ action: 'list', tab: 'Sep 26', key: KEY });
  assert.equal(sep.items.length, 5, 'порожній рядок пропущено');
  assert.equal(sep.items[0].status, 'In progress');
  assert.equal(sep.items[0].deadlineIso, '2026-10-03');
});

await test('кеш: ручна правка без onEdit не видна, fresh=1 і onEdit — видно', async () => {
  const q = { action: 'list', tab: 'Aug 26', key: KEY };
  const title = (r) => r.items.find((i) => i.row === 3).title;
  assert.equal(title(await get(q)), 'Серпневий лендінг');
  tab('Aug 26').rows[2][2] = 'Серпневий лендінг 2';
  assert.equal(title(await get(q)), 'Серпневий лендінг', 'з кешу');
  assert.equal(title(await get({ ...q, fresh: 1 })), 'Серпневий лендінг 2');
  tab('Aug 26').rows[2][2] = 'Серпневий лендінг';
  onEdit();
  assert.equal(title(await get(q)), 'Серпневий лендінг');
});

await test('conflict при зміненому рядку (API)', async () => {
  const r = await post({ action: 'update', key: KEY, tab: 'Sep 26', row: 3, expectedTitle: 'Інша назва', patch: { status: 'Done' } });
  assert.deepEqual(r, { error: 'conflict' });
  assert.equal(rowOf('Sep 26', 'Банер для сайту')[3], 'In progress ');
});

await test('завантаження дошки з #key, ключ прибрано з адреси', async () => {
  await page.goto(`${PAGE}#key=${KEY}`);
  await card('Банер для сайту').waitFor();
  assert.equal(new URL(page.url()).hash, '');
  assert.equal(await page.evaluate(() => localStorage.getItem('kanbanKey')), KEY);
  for (const s of ['Not started', 'In progress', 'Done']) assert.equal(await column(s).count(), 1);
  assert.equal(await column('Not started').locator('article').count(), 3);
  assert.match(await page.textContent('#total'), /21 год/);
});

await test('фільтр за проєктом', async () => {
  await page.selectOption('#projectFilter', 'ATE');
  assert.equal(await page.locator('article').count(), 2);
  assert.match(await page.textContent('#total'), /11 год/);
  await page.selectOption('#projectFilter', '');
});

await test('перемикання місяця + невідомий статус окремою колонкою', async () => {
  await page.selectOption('#tabSelect', 'Aug 26');
  await card('Гайдлайн кольорів').waitFor();
  await page.selectOption('#tabSelect', 'September 25');
  await column('On hold').locator('article', { hasText: 'Старий макет' }).waitFor();
  await page.selectOption('#tabSelect', 'Sep 26');
  await card('Банер для сайту').waitFor();
});

await test('перетягування в Done записує статус і дату', async () => {
  const from = await card('Іконки для застосунку').boundingBox();
  const to = await column('Done').boundingBox();
  await page.mouse.move(from.x + 20, from.y + 15);
  await page.mouse.down();
  await page.mouse.move(from.x + 60, from.y + 30, { steps: 5 });
  await page.mouse.move(to.x + to.width / 2, to.y + to.height - 10, { steps: 15 });
  await page.mouse.up();
  const row = rowOf('Sep 26', 'Іконки для застосунку');
  await until(() => row[3] === 'Done');
  assert.ok(isDate(row[6]), 'дата виконання — Date');
  assert.equal(row[6].toDateString(), new Date().toDateString());
  await column('Done').locator('article', { hasText: 'Іконки для застосунку' }).waitFor();
});

await test('редагування годин "4,5" → число 4.5', async () => {
  await page.waitForTimeout(350); // клік одразу після drop ігнорується навмисно
  await card('Банер для сайту').locator('p').click();
  await page.fill('#editForm [name=hours]', '4,5');
  await page.click('#editForm button[value=save]');
  const row = rowOf('Sep 26', 'Банер для сайту');
  await until(() => row[2] === 4.5);
  assert.equal(typeof row[2], 'number');
  await card('Банер для сайту').getByText('4,5 год').waitFor();
});

await test('дедлайн з часом → Date з годинами, час видно в картці', async () => {
  await card('Пости для соцмереж').locator('p').click();
  await page.fill('#editForm [name=deadline]', '2026-10-01');
  await page.fill('#editForm [name=deadlineTime]', '14:30');
  await page.click('#editForm button[value=save]');
  const row = rowOf('Sep 26', 'Пости для соцмереж');
  await until(() => isDate(row[5]) && row[5].getHours() === 14);
  assert.equal(row[5].getMinutes(), 30);
  await card('Пости для соцмереж').getByText('до 01.10.2026 14:30').waitFor();
  await card('Пости для соцмереж').locator('p').click();
  assert.equal(await page.inputValue('#editForm [name=deadlineTime]'), '14:30');
  await page.click('#editForm button[value=cancel]');
});

await test('додавання задачі (з №, не затираючи підсумок)', async () => {
  await page.selectOption('#tabSelect', 'September 25');
  await card('Старий макет').waitFor();
  await page.locator('[data-add="Not started"]').click();
  await page.fill('[data-add-form] input', 'Нова задача');
  await page.press('[data-add-form] input', 'Enter');
  const rows = tab('September 25').rows;
  await until(() => rows.some((r) => r[1] === 'Нова задача'));
  assert.deepEqual(rows[4].slice(0, 4), [3, 'Нова задача', '', 'Not started']);
  assert.equal(rows[5][2], 3, 'рядок підсумку не зачеплено');
  await column('Not started').locator('article[data-row="5"]', { hasText: 'Нова задача' }).waitFor();
});

await test('conflict у UI: дошка перезавантажується', async () => {
  await page.selectOption('#tabSelect', 'Sep 26');
  await card('Логотип').waitFor();
  rowOf('Sep 26', 'Логотип')[0] = 'Логотип v2';
  onEdit();
  await card('Логотип').locator('select').selectOption('In progress');
  await page.locator('#toast', { hasText: 'змінився' }).waitFor();
  await column('Not started').locator('article', { hasText: 'Логотип v2' }).waitFor();
  assert.equal(rowOf('Sep 26', 'Логотип v2')[3], 'Not started');
});

await test('статистика: усі місяці одним запитом, Overtime не рахується', async () => {
  const { tabs } = await get({ action: 'all', key: KEY });
  assert.deepEqual(tabs.map((t) => t.tab), ['Sep 26', 'Aug 26', 'September 25']);
  const expected = tabs.flatMap((t) => t.items).filter((i) => !/^overtime/i.test(i.title)).reduce((s, i) => s + (i.hours || 0), 0);
  await page.click('[data-view=stats]');
  await page.getByText('Найбільші задачі').waitFor();
  assert.equal(await page.locator('#board').isHidden(), true);
  const totalEl = page.locator('#stats p', { hasText: 'Всього годин' }).locator('xpath=following-sibling::p[1]');
  const shown = async () => Number((await totalEl.textContent()).replace(/\s/g, '').replace(',', '.'));
  for (let i = 0; i < 50 && (await shown()) !== expected; i++) await page.waitForTimeout(100); // спершу — з кешу, потім свіже
  assert.equal(await shown(), expected);
  assert.equal(await page.getByText('Overtime extimate').count(), 0);
  await page.getByText('Вер 25 — Вер 26').waitFor();
  await page.reload();
  await page.getByText('Найбільші задачі').waitFor(); // вид запам'ятовується
  await page.click('[data-view=board]');
  await page.locator('article').first().waitFor();
});

await test('записи під LockService', () => assert.ok(fixture.locks >= 4));

await test('екран ключа і невірний ключ', async () => {
  const p = await browser.newPage();
  await p.goto(PAGE);
  await p.getByText('Введіть ключ доступу').waitFor();
  await p.fill('#keyInput', 'wrong');
  await p.click('#keyForm button');
  await p.locator('#keyError', { hasText: 'Невірний ключ' }).waitFor();
  await p.close();
});

await test('скриншоти 1280/390 × світла/темна', async () => {
  mkdirSync('test/screenshots', { recursive: true });
  for (const scheme of ['light', 'dark']) {
    for (const [w, h] of [[1280, 900], [1024, 768], [390, 844]]) {
      const p = await browser.newPage({ viewport: { width: w, height: h }, colorScheme: scheme });
      await p.goto(`${PAGE}#key=${KEY}`);
      await p.locator('article').first().waitFor();
      await p.screenshot({ path: `test/screenshots/board-${w}-${scheme}.png` });
      const overflow = await p.evaluate(() => [document.documentElement, document.querySelector('#board')].some((el) => el.scrollWidth > el.clientWidth));
      assert.equal(overflow, false, `горизонтальний скрол на ${w}px`);
      if (w === 390) {
        await p.locator('article').first().click();
        await p.screenshot({ path: `test/screenshots/edit-${w}-${scheme}.png` });
      }
      await p.close();
    }
  }
});

await browser.close();
server.close();
console.log(failed ? `\n${failed} тест(и) впали` : '\nУсі тести пройшли');
process.exit(failed ? 1 : 0);
