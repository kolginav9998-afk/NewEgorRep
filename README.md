# NewEgorRep
Poke no know

## WMS на LibreOffice Calc

Складской учёт в одной книге LibreOffice Calc: каждое движение проводится своей операцией, записывается во внешний журнал и защищено от случайных правок; отдельные инструменты работают со снимком склада.

- Руководство кладовщика — [`docs/OPERATOR_GUIDE.md`](docs/OPERATOR_GUIDE.md)
- Проверка складского ПК и принтера — [`docs/PC_VALIDATION.md`](docs/PC_VALIDATION.md)
- Инструменты `WMS_TOOLBOX` — [`docs/TOOLBOX.md`](docs/TOOLBOX.md)
- Резервные копии и восстановление — [`docs/BACKUP_RECOVERY.md`](docs/BACKUP_RECOVERY.md)
- Примечания к выпуску — [`docs/RELEASE_NOTES.md`](docs/RELEASE_NOTES.md)
- Состояние проекта — [`WMS_STATE.md`](WMS_STATE.md), техническое состояние — [`FINAL_HANDOFF.md`](FINAL_HANDOFF.md)

Выпуск собирается из исходников: `python3 tools/build_release.py ПАПКА` (нужны LibreOffice и `python3-uno`).
