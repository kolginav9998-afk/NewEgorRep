import test from 'node:test';
import assert from 'node:assert/strict';
import { makeTask, toggleTask, splitTasks, serialize, parseImport } from '../events.js';

test('makeTask: обрезает пробелы, ставит provenance', () => {
  const t = makeTask('  купить молоко  ', 1000);
  assert.equal(t.payload.text, 'купить молоко');
  assert.equal(t.source, 'manual');
  assert.equal(t.confidence, 1);
  assert.deepEqual(t.evidence, []);
});
test('makeTask: пустая задача — ошибка', () => {
  assert.throws(() => makeTask('   '));
  assert.throws(() => makeTask(undefined));
});
test('toggleTask: туда и обратно', () => {
  const t = makeTask('a', 1);
  const d = toggleTask(t, 50);
  assert.equal(d.payload.done, true);
  assert.equal(d.payload.doneAt, 50);
  const back = toggleTask(d, 60);
  assert.equal(back.payload.done, false);
  assert.equal(back.payload.doneAt, null);
});
test('splitTasks: порядок и разделение', () => {
  const a = makeTask('a', 1), b = makeTask('b', 2), c = toggleTask(makeTask('c', 3), 10), d = toggleTask(makeTask('d', 4), 20);
  const { open, done } = splitTasks([a, b, c, d, { id: 'x', type: 'note', ts: 1, payload: {} }]);
  assert.deepEqual(open.map(e => e.payload.text), ['b', 'a']);
  assert.deepEqual(done.map(e => e.payload.text), ['d', 'c']);
});
test('экспорт → импорт без потерь', () => {
  const evs = [makeTask('a', 1), toggleTask(makeTask('b', 2), 5)];
  const { events, skipped } = parseImport(serialize(evs, 9));
  assert.equal(skipped, 0);
  assert.deepEqual(events, evs);
});
test('импорт: битые записи пропускаются, чужой файл отклоняется', () => {
  const good = makeTask('ok', 1);
  const text = JSON.stringify({ app: 'lifeos', events: [good, { id: 1 }, null, { id: 'z', type: 'task' }] });
  const { events, skipped } = parseImport(text);
  assert.equal(events.length, 1);
  assert.equal(skipped, 3);
  assert.throws(() => parseImport('{"foo":1}'), /не резервная копия/);
  assert.throws(() => parseImport('не json'), /не JSON/);
});
