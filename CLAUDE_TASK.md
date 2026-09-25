# CLAUDE TASK — CORE PHASE 1: RELIABILITY SKELETON

## ВАЖНО

Это первое production-задание только после явного разрешения пользователя.

До такого разрешения НЕ выполнять задачу.

### GitHub
Без отдельного разрешения пользователя в сообщении запрещены:
- commit;
- push;
- PR;
- создание/переключение/удаление веток;
- изменение main.

Если пользователь разрешает код, но не GitHub-действия — писать изменения только в рабочей копии и показать diff/report.

## Перед началом

Прочитать:
1. WMS_MASTER_SPEC v0.3
2. WMS_DECISIONS v0.3
3. CLAUDE_REPORT.md
4. CLAUDE_REPORT_P1_P4.md

Ничего не переносить из одноразового прототипа вслепую. Использовать его только как подтверждённый технический образец.

## Цель этапа

Создать production-скелет надёжности, на который позже будут опираться приход, выдача, возврат и исправления.

На этом этапе НЕ делать полноценные бизнес-листы, отчёты, аналитику, акты, комплекты и миграцию.

## Реализовать

### 1. Структура исходников
Текстовые Basic-модули в репозитории:
- WmsCore
- WmsJournal
- WmsRecovery
- WmsLock
- WmsBackup
- WmsConfig
- WmsDiagnostics

Плюс минимальные тестовые модули, если они не смешиваются с production.

### 2. `_SYS`
Минимальные служебные значения:
- schema/version
- instance ID
- LAST_SEQ
- NEXT_EI
- NEXT_NO
- NEXT_RET
- journal file/offset/seq
- registered workbook path
- current transaction marker
- last clean save marker

### 3. Lock
- lock-файл в WMS_Journal;
- владелец lock;
- корректное снятие по владельцу;
- после аварии — явная команда `Снять блокировку WMS`;
- второй экземпляр не проводит.

### 4. Journal
Формат:
`J1;seq;timestamp;instance;type;fields...;END;length`

- UTF-8
- LF
- append only
- monthly
- проверка ожидаемого размера до append
- seq strictly monotonic
- проверка saved position по содержимому
- fallback full read

### 5. Two-phase transaction
Одна функция применения операции.

Фаза 1:
- validate
- normalize
- calculate
- no mutation

Фаза 2:
- STARTED + printable before-image
- apply mutations
- append journal
- COMMITTED + LAST_SEQ

До journal commit → rollback.
После journal commit → finish/recover.

### 6. Startup fail-closed
Startup должен различать:
- clean
- stale lock
- unsaved journal tail
- old backup opened as working book
- journal mismatch
- saved without WMS/macros
- corrupt transaction marker

При любой проблеме проведение заблокировано.

Явные recovery actions:
- Восстановить
- Отложить хвост журнала
- Снять блокировку WMS
- Сделать рабочим файлом

### 7. Backup
- daily backup once
- backup before upgrade/migration function hooks
- rotation
- manual backup
- never use backup as substitute for journal

### 8. Undo
Единый wrapper учётной записи:
- lock Undo
- operation
- clear Undo
- unlock safely

### 9. Diagnostics
Минимальный self-check:
- journal continuity
- marker state
- lock ownership
- registered path
- counters sanity
- ODS format
- macros/session marker

### 10. Locale-independent quantity parser
Production helper:
- accepts `1,5`, `1.5`
- max 3 decimals
- rejects date-like, exponent, negative, ambiguous grouping, fractions/AutoCorrect values
- returns canonical numeric value

### 11. AutoInput management
- remember current setting
- disable at WMS startup
- restore at clean close where technically safe
- document that setting is global to LibreOffice

## Tests required before completion

Automated/synthetic tests for:
- normal operation;
- error before journal;
- error after journal;
- kill/reopen equivalent where feasible;
- journal LF/CRLF conversion safety;
- stale lock;
- second instance;
- old backup;
- journal gap;
- invalid saved offset;
- Ctrl+Z;
- quantity parser ru/en cases;
- save/reopen;
- ODS-only guard.

No user business operation may silently proceed after startup error.

## Acceptance

Return:
- list of created/changed files;
- architecture notes;
- test results;
- known limitations;
- timing results;
- exact Git diff summary;
- recommendation whether Phase 2 `Выдачи` can begin.

Do not expand scope.

Do not implement Phase 2 in the same task.

---

# Дополнение пользователя от 25.09.2026 — проверка перед commit Core Phase 1

> Текст дополнения перенесён дословно. После переноса документов:
> - `CLAUDE_REPORT.md` из задания выше лежит в `docs/reviews/CLAUDE_REPORT_ARCHITECTURE.md`;
> - `CLAUDE_REPORT_P1_P4.md` — в `docs/reviews/CLAUDE_REPORT_P1_P4.md`;
> - WMS MASTER SPEC v0.3 — в `WMS_MASTER_SPEC.md`;
> - WMS DECISIONS v0.3 — в `WMS_DECISIONS.md`.

Core Phase 1 предварительно принят.
Перед commit необходимо сделать ещё одну обязательную проверку.

## 1. Добавить регрессионный тест Undo

Ты указал, что при просмотре кода обнаружил и исправил сценарий:
операция уже попала в журнал, но завершение commit в книге завершилось ошибкой → Undo мог остаться заблокированным до перезапуска.
При этом ты отдельно отметил, что этот путь тестом не покрыт.
Добавь отдельный автоматический регрессионный тест именно на этот сценарий.
Тест должен доказать:

1. операция дошла до journal commit point;
2. после этого искусственно возникает ошибка на этапе завершения операции в книге;
3. recovery/доведение приводит состояние к корректному;
4. UndoManager после завершения recovery не остаётся заблокированным;
5. последующий обычный ручной ввод в Calc и следующая складская операция работают;
6. операция существует ровно один раз;
7. журнал, LAST_SEQ и книга согласованы.

После этого снова запусти полный suite.
Критерий:
0 FAIL, 0 SKIP.
Если новый тест выявит проблему — сначала исправить её и снова прогнать полный suite.

## 2. После успешного теста разрешаю Git-действия

Если полный финальный suite проходит:
РАЗРЕШАЮ:

* сделать один commit Core Phase 1;
* push этого commit только в существующую ветку
`claude/festive-lamport-l4f8f1`.

НЕ РАЗРЕШАЮ:

* создавать PR;
* merge;
* менять `main`;
* force push;
* создавать другую ветку;
* удалять ветки;
* начинать Phase 2 в этом же задании.

## 3. Что включить в commit

Обязательно:
Production
`src/basic/`

* WmsCore
* WmsJournal
* WmsRecovery
* WmsLock
* WmsBackup
* WmsConfig
* WmsDiagnostics

Tests
Все актуальные тесты Core Phase 1, включая новый Undo regression test.
Build
`tools/build_ods.py`
Project memory
В корне репозитория должен существовать один текущий источник истины:

* `WMS_MASTER_SPEC.md` — актуальное содержание v0.3;
* `WMS_DECISIONS.md` — актуальные решения;
* `WMS_STATE.md`;
* `TEST_REPORT.md`;
* `CLAUDE_TASK.md`;
* `CLAUDE_REPORT_CORE_PHASE1.md`;
* `.gitignore`.

`CLAUDE_REPORT_P1_P4.md` тоже сохранить как архитектурную историю, желательно в:
`docs/reviews/CLAUDE_REPORT_P1_P4.md`
Если предыдущий архитектурный `CLAUDE_REPORT.md` нужен для истории:
`docs/reviews/CLAUDE_REPORT_ARCHITECTURE.md`
Не создавать несколько одновременно «актуальных» файлов вроде:
`WMS_MASTER_SPEC_v0.1.md`
`WMS_MASTER_SPEC_v0.2.md`
`WMS_MASTER_SPEC_v0.3.md`
в корне.
В корне должен быть один:
`WMS_MASTER_SPEC.md`
с указанной внутри версией `v0.3`.
Старые версии при необходимости хранить только в `docs/history/`.
То же правило для `WMS_DECISIONS.md`.

## 4. Перед commit проверить

* `git status`;
* никаких временных ODS;
* никаких журналов тестовых запусков;
* никаких backup-файлов;
* никаких профилей LibreOffice;
* никаких кэшей;
* никаких реальных рабочих данных;
* никаких абсолютных путей к временной машине;
* никаких секретов;
* тестовый Basic не встраивается в production ODS;
* production ODS собирается только из разрешённых production-модулей.

## 5. Commit

Один логический commit для Phase 1.
Подходящее сообщение:
`core: add WMS reliability skeleton with recovery and journal`
После push ничего больше не менять.
Верни мне:

1. commit hash;
2. ветку;
3. полный `git status`;
4. список файлов commit;
5. результат последнего полного suite;
6. отдельно результат нового Undo regression test;
7. краткий diff-stat;
8. подтверждение, что PR и merge не делались;
9. подтверждение, что Phase 2 ещё не начинался.

После этого остановись.
