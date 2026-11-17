# Автоматизация SolidWorks

- [drawingmerger](drawingmerger/) — макрос объединения V16.29:
  основной модуль и код формы. Папка `runs` не создаётся; ошибки и время работы
  записываются в `<имя чертежа>_merge.log` рядом с чертежом. Изменение проверено
  статически; запуск V16.29 в SolidWorks ещё не проверен.
- [autosizedrilling](autosizedrilling/README.md) — исходная DLL, C#-кандидат и тесты размеров.
- [DrillingView](DrillingView/README.md) — подготовленная папка проекта.
- [engineering-knowledge](engineering-knowledge/README.md) — накопленные знания и инструкции агентов.

Папки, сохранённые без перемещения и изменения:

- [diagnostics](diagnostics/) — диагностика заметок.
- [docs/agent-knowledge](docs/agent-knowledge/) — существующее исследование сварных тел.
- [weldment_drilling](weldment_drilling/README_RU.md) — отдельный проект сверловки.

В `drawingmerger` оставлена только последняя версия по просьбе пользователя.
Предыдущие версии доступны в истории Git.

[Инструкции для Codex](AGENTS.md).
