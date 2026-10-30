# Инженерная память проекта SolidWorks

Вход для нового Codex: [корневой AGENTS](../AGENTS.md) → [правила координатора](AGENTS.md) → [каталог знаний](topics/README.md) → только нужный проект и функции.

- [topics](topics/README.md) — **пять предметных разделов**: виды, связанные заметки, PDF, размеры, VBA/COM. Это не отдельные агенты.
- [agents](agents/) — **три роли**: `coder` (исходники), `api_expert` (документация API), `debugger` (ошибки/тесты).
- [конфигурация Codex](../.codex/config.toml) — маршруты трёх ролей, максимум четыре потока с координатором.
- [реестр проверок](topics/evidence.md) — точные статусы версий, ограничения и источники.
- [снимок AutoSizeDrilling](references/autosizedrilling/README.md) — сохранённые исходники и результаты проверок.

Проекты: [drawingmerger](../drawingmerger/README.md) (V16.21 подтверждена), [weldment_drilling](../weldment_drilling/README_RU.md) (R7 подтверждена, R17 кандидат), [autosizedrilling](../autosizedrilling/README.md) (C# кандидат), [DrillingView](../DrillingView/README.md). Архивы и диагностика остаются на своих местах.

Проверка маршрутов, ссылок и неизменности архивов: `python3 engineering-knowledge/tools/check_agent_knowledge.py`. Её успех не означает запуска макросов в SolidWorks.
