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

Текущий выпуск — **0.7.2** (книга `WMS_LEGACY_TRANSFER.ods`: перенос старой таблицы «Заказы» в WMS — заказы, открытые позиции, старые ЕИ, текущие остатки; 0.7.2 понимает старые статусы и даты так, как их писали люди, — «Размещен», «Ожидаем», «1.9.26», а просрочки как статуса больше нет: прошедшая ожидаемая дата только подсвечивается и считается в контроле сроков; до этого 0.7.0 — журнал машин «Приход авто», аналитика с закономерностями, «Контроль дня»). Выпуск собирается из исходников: `python3 tools/build_release.py ПАПКА --zip` (нужны LibreOffice и `python3-uno`); книга 0.6.0 обновляется на месте — `tools/upgrade.py` ([`docs/BACKUP_RECOVERY.md`](docs/BACKUP_RECOVERY.md), раздел 7).

- Контракт снимка и пакетов для инструментов — [`docs/EXPORT_CONTRACT.md`](docs/EXPORT_CONTRACT.md)
