// Запуск: node tests/e2e.mjs  (поднимает свой статический сервер)
import { createRequire } from 'node:module';
import http from 'node:http';
import fs from 'node:fs';
import path from 'node:path';
import assert from 'node:assert/strict';
const { chromium } = createRequire(import.meta.url)('playwright');

const root = path.resolve(import.meta.dirname, '..');
const types = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.webmanifest': 'application/manifest+json', '.png': 'image/png' };
const server = http.createServer((req, res) => {
  const p = path.join(root, req.url.split('?')[0] === '/' ? 'index.html' : req.url.split('?')[0]);
  fs.readFile(p, (err, buf) => {
    if (err) { res.writeHead(404); return res.end(); }
    res.writeHead(200, { 'content-type': types[path.extname(p)] || 'application/octet-stream' });
    res.end(buf);
  });
}).listen(0);
const url = `http://localhost:${server.address().port}/`;

const browser = await chromium.launch();
const ctx = await browser.newContext({ viewport: { width: 390, height: 844 }, acceptDownloads: true });
const page = await ctx.newPage();
const errors = [];
page.on('pageerror', e => errors.push(e.message));
page.on('console', m => m.type() === 'error' && errors.push(m.text()));

await page.goto(url);
await page.evaluate(() => navigator.serviceWorker.ready);
await page.reload(); // чтобы SW взял контроль

// офлайн
await ctx.setOffline(true);
await page.reload();
for (const t of ['one', 'two', 'three', 'four', 'five']) {
  await page.fill('#add-input', t);
  await page.press('#add-input', 'Enter');
}
await page.waitForFunction(() => document.querySelectorAll('#open-list .item').length === 5, null, { timeout: 3000 });

// выполнить
await page.locator('#open-list .item').first().locator('.check').click();
await page.waitForFunction(() => document.querySelectorAll('#open-list .item').length === 4);
assert.equal(await page.textContent('#done-count'), '1');

// перезапуск — данные на месте
await page.reload();
assert.equal(await page.locator('#open-list .item').count(), 4, 'после перезапуска');
assert.equal(await page.textContent('#done-count'), '1');
assert.match(await page.textContent('#status'), /офлайн/);
await ctx.setOffline(false);

// экспорт
await page.click('details.data summary');
const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#export-btn')]);
const file = await dl.path();
const exported = JSON.parse(fs.readFileSync(file, 'utf8'));
assert.equal(exported.events.length, 5);

// удалить всё, импортировать обратно
for (let i = 4; i > 0; i--) {
  await page.locator('#open-list .item .del').first().click();
  await page.waitForFunction(n => document.querySelectorAll('#open-list .item').length === n, i - 1);
}
await page.locator('#done-block summary').click();
await page.locator('#done-list .item .del').first().click();
await page.waitForFunction(() => document.querySelectorAll('.item').length === 0);
await page.setInputFiles('#import-input', file);
await page.waitForFunction(() => document.querySelectorAll('#open-list .item').length === 4);
assert.equal(await page.textContent('#done-count'), '1', 'импорт восстановил');

// манифест
const man = await (await ctx.request.get(url + 'manifest.webmanifest')).json();
assert.equal(man.display, 'standalone');

assert.deepEqual(errors, [], 'нет ошибок в консоли: ' + errors.join('; '));
await browser.close(); server.close();
console.log('E2E OK');
