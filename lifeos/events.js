// Ядро данных: журнал событий. Задача — это Event типа "task".
// Поля source/confidence/evidence есть с первого дня (см. план: контроль достоверности).

export const SCHEMA_VERSION = 1;

export function newId() {
  if (globalThis.crypto?.randomUUID) return crypto.randomUUID();
  return 'id-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 10);
}

export function makeTask(text, now = Date.now()) {
  const clean = String(text ?? '').trim();
  if (!clean) throw new Error('empty task');
  return {
    id: newId(),
    type: 'task',
    ts: now,
    payload: { text: clean, done: false, doneAt: null },
    source: 'manual',   // manual | voice | ai | notification | import
    confidence: 1,      // 0..1
    evidence: [],
  };
}

export function toggleTask(ev, now = Date.now()) {
  const done = !ev.payload.done;
  return { ...ev, payload: { ...ev.payload, done, doneAt: done ? now : null } };
}

export function splitTasks(events) {
  const tasks = events.filter(e => e.type === 'task');
  const open = tasks.filter(e => !e.payload.done).sort((a, b) => b.ts - a.ts);
  const done = tasks.filter(e => e.payload.done).sort((a, b) => b.payload.doneAt - a.payload.doneAt);
  return { open, done };
}

export function serialize(events, now = Date.now()) {
  return JSON.stringify({ app: 'lifeos', schema: SCHEMA_VERSION, exportedAt: now, events }, null, 2);
}

// Возвращает только валидные события; мусор отбрасывает и считает.
export function parseImport(text) {
  let data;
  try { data = JSON.parse(text); } catch { throw new Error('Это не JSON-файл'); }
  if (!data || data.app !== 'lifeos' || !Array.isArray(data.events)) {
    throw new Error('Это не резервная копия Life OS');
  }
  const valid = [];
  let skipped = 0;
  for (const e of data.events) {
    const ok = e && typeof e.id === 'string' && typeof e.type === 'string'
      && typeof e.ts === 'number' && e.payload && typeof e.payload === 'object';
    if (ok) valid.push({ source: 'import', confidence: 1, evidence: [], ...e });
    else skipped++;
  }
  return { events: valid, skipped };
}
