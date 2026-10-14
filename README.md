# NewEgorRep
Poke no know

## WMS на LibreOffice Calc

Складской учёт в одной книге LibreOffice Calc: каждое движение проводится своей операцией, записывается во внешний журнал и защищено от случайных правок; отдельные инструменты работают со снимком склада.

- Руководство кладовщика — [`docs/OPERATOR_GUIDE.md`](docs/OPERATOR_GUIDE.md)
- Проверка складского ПК и принтера — [`docs/PC_VALIDATION.md`](docs/PC_VALIDATION.md)
- Инструменты `WMS_TOOLBOX` — [`docs/TOOLBOX.md`](docs/TOOLBOX.md)
- Перенос старой таблицы «Заказы» в WMS — [`docs/LEGACY_TRANSFER.md`](docs/LEGACY_TRANSFER.md)
- Резервные копии и восстановление — [`docs/BACKUP_RECOVERY.md`](docs/BACKUP_RECOVERY.md)
- Примечания к выпуску — [`docs/RELEASE_NOTES.md`](docs/RELEASE_NOTES.md)
- Состояние проекта — [`WMS_STATE.md`](WMS_STATE.md), техническое состояние — [`FINAL_HANDOFF.md`](FINAL_HANDOFF.md)

Текущий выпуск — **0.8.0 PRIME**: новая версия ставится **из рабочей книги 0.7 со всеми данными** (книга `WMS_UPGRADE.ods`, старая книга не изменяется), исправления позиции заказа и переноса, живые «Дашборд», «Поиск» с карточкой ЕИ и «Карта склада» по вашей топологии, учёт по документу и выдача частями («1 кан», «1,5 м») и в процентах, несколько мест у одной позиции ([`docs/RELEASE_NOTES.md`](docs/RELEASE_NOTES.md)). До этого 0.7.2 — перенос старой таблицы «Заказы» (`WMS_LEGACY_TRANSFER.ods`) с пониманием старых статусов и дат; 0.7.0 — журнал машин «Приход авто», аналитика с закономерностями, «Контроль дня». Выпуск собирается из исходников: `python3 tools/build_release.py ПАПКА --zip` (нужны LibreOffice и `python3-uno`); обновление книги 0.7 — мастер `WMS_UPGRADE.ods` ([`docs/OPERATOR_GUIDE.md`](docs/OPERATOR_GUIDE.md), раздел 11), книги 0.6.0 — сначала на месте `tools/upgrade.py` ([`docs/BACKUP_RECOVERY.md`](docs/BACKUP_RECOVERY.md), раздел 7).

- Контракт снимка и пакетов для инструментов — [`docs/EXPORT_CONTRACT.md`](docs/EXPORT_CONTRACT.md)
