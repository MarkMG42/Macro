# Инженерная память и инструкции агентов

Это единое место для накопленных знаний и правил их использования.
Дата реорганизации: 29 октября 2026 года.

- [AGENTS.md](AGENTS.md) — инструкции координатора: выбор областей, делегирование,
  проверка доказательств и обновление знаний после значимой задачи.
- [topics/README.md](topics/README.md) — каталог пяти технических специализаций.
- [agents/](agents) — переиспользуемые TOML-инструкции профильных агентов.
- [references/autosizedrilling/README.md](references/autosizedrilling/README.md) —
  неизменный снимок источника, SHA256 и результаты проверки.
- [tools/check_agent_knowledge.py](tools/check_agent_knowledge.py) — проверка
  ссылок, маршрутов ролей, снимка и сохранённых комплектов.

Из корня репозитория:

```sh
python3 engineering-knowledge/tools/check_agent_knowledge.py
```

Корневой [AGENTS.md](../AGENTS.md) — короткий вход для автоматического обнаружения
инструкций Codex. [.codex/config.toml](../.codex/config.toml) регистрирует роли,
ссылаясь на инструкции в этой папке. Эти два файла оставлены на стандартных местах,
чтобы перенос базы не ломал новый чат и конфигурацию CLI.

Проекты рядом: [drawingmerger](../drawingmerger/README.md),
[autosizedrilling](../autosizedrilling/README.md), [DrillingView](../DrillingView/README.md).
Все выпуски объединения и их архивы теперь в drawingmerger; основной рабочий —
V16.21. Остальные сохранены как история с прежними статусами. По уточнению
пользователя diagnostics, docs/agent-knowledge и weldment_drilling оставлены в корне
без изменения файлов. Ссылки на них ведут в исходные папки.
