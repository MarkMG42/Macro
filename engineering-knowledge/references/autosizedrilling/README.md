# Неизменный комплект AutoSizeDrilling

Источник: [MarkMG42/Macro-improvement, autosizedrilling](https://github.com/MarkMG42/Macro-improvement/tree/69b28526b022503ec8a1c4951a02ec97a48a8619/autosizedrilling).
Срез: 2026-10-28, commit `69b28526b022503ec8a1c4951a02ec97a48a8619`.

[Локальный ZIP-снимок](autosizedrilling-69b28526b022.zip) сохраняет **все 21
отслеживаемый файл папки**: исходную DLL/PDB, interop DLL, C#-исходники, тесты,
кандидат 1.1.2, прежние download-архивы 1.1.1/1.1.2 и документы анализа.
Файлы побайтово совпадают с указанным commit; это архив справочного материала,
он не подключается к макросу объединения и не меняет его сборку.
[manifest.json](manifest.json) фиксирует SHA256 ZIP и каждого вложенного файла.
Документация/код не требуют доступности внешнего GitHub: ZIP включён локально.

## Какой вариант брать

| Комплект внутри ZIP | Назначение / статус |
|---|---|
| autosizedrilling/AutoSizeDrilling.dll + .pdb | Исходный VB.NET net48/x64 1.0.0.0, неизменный вариант для сравнения/отката. Есть документированные ограничения и жалоба на отдельные размеры. Нет общего подтверждения нативной корректности/совместимости в доступных материалах |
| autosizedrilling/src/ | Поддерживаемая восстановленная C#-реализация 1.1.2, а не исходный VB.NET-проект. Геометрические алгоритмы покрыты тестами |
| autosizedrilling/candidate/ | DLL 1.1.2 и её собственный PDB. Кандидат для SW, **не подтверждённая рабочая замена оригинала** |
| autosizedrilling/tests/ | 23 исполняемых теста общей Geometry.cs без SolidWorks |
| autosizedrilling/downloads/ | Исторические упакованные 1.1.1 и 1.1.2, без повышения их статуса только из-за наличия архива |

Не смешивать PDB разных сборок. Root SHA256SUMS и candidate/SHA256SUMS проверяются
из своих каталогов. SHA подтверждает целостность, не корректность алгоритма.
Оригинального VSTA-проекта нет. DLL не standalone EXE; загрузку через поддерживаемый
VSTA-шаблон конкретной установки ещё требуется проверить. Не подменять оригинал
кандидатом без испытания на копии чертежа.

## Чтение и повторное использование

Из корня Macro распаковать в отдельный каталог (без изменения текущих макросов):

```sh
unzip engineering-knowledge/references/autosizedrilling/autosizedrilling-69b28526b022.zip -d /tmp/autosizedrilling-reference
cd /tmp/autosizedrilling-reference
# .NET SDK 8.0; источники interop DLL уже внутри комплекта:
dotnet run --project autosizedrilling/tests/GeometryTests.csproj -c Release
dotnet build autosizedrilling/src/AutoSizeDrilling.csproj -c Release
```

Тесты target net8.0, рабочая сборка target net48/x64; NuGet reference assemblies
нужны для сборки, COM/SolidWorks для запуска DLL. Используемые API проверяются
компиляцией с предоставленными interop, не нативным исполнением.

Ключевые исходники с привязкой к commit (также лежат в ZIP по тем же путям):

- [Geometry.cs](https://github.com/MarkMG42/Macro-improvement/blob/69b28526b022503ec8a1c4951a02ec97a48a8619/autosizedrilling/src/Geometry.cs): численная геометрия, допуски, группы отверстий.
- [SolidWorksMacro.cs](https://github.com/MarkMG42/Macro-improvement/blob/69b28526b022503ec8a1c4951a02ec97a48a8619/autosizedrilling/src/SolidWorksMacro.cs): COM, размеры, сверка и отмена.
- [Program.cs](https://github.com/MarkMG42/Macro-improvement/blob/69b28526b022503ec8a1c4951a02ec97a48a8619/autosizedrilling/tests/Program.cs): регрессии без копирования реализации.
- [CHANGES.md](https://github.com/MarkMG42/Macro-improvement/blob/69b28526b022503ec8a1c4951a02ec97a48a8619/autosizedrilling/CHANGES.md), [REVIEW.md](https://github.com/MarkMG42/Macro-improvement/blob/69b28526b022503ec8a1c4951a02ec97a48a8619/autosizedrilling/REVIEW.md): история гипотез/дефектов и статический разбор оригинала.

Сжатые инженерные выводы — [dimensions.md](../../topics/dimensions.md).
Снимок не обновлять молча: новый источник — новый commit/архив/manifest с явным
статусом, не перезапись этого комплекта. Проверка целостности и свежие результаты
автоматической сборки/тестов — [validation.json](validation.json).

Повторная проверка здесь: SDK8.0.408, 23/23 теста; net48/x64 build — 0 ошибок,
0 предупреждений. Свежие DLL/PDB **не побайтово равны** сохранённому candidate;
оба хеша записаны в validation.json. Причина расхождения не установлена,
тождество готовой DLL и повторной сборки не заявляется. Автотесты относятся
к сохранённому исходному Geometry.cs. Исходные candidate DLL/PDB не заменялись.

29.10.2026: комплект также распакован в соседний проект
[autosizedrilling](../../../autosizedrilling/README.md). Все21 исходных файла там
побайтово сверяются с manifest; архив остаётся неизменным контрольным снимком.
