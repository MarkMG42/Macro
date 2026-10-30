# VBA/COM SolidWorks — короткие правила

## C1. Успех API подтверждается состоянием
- COM Boolean может иметь неклассическое VBA-представление; не использовать неосторожный `Not` как критерий успеха. Предпочесть `If condition Then Exit Sub` и явный отказ.
- VBA `And`/`Or` не short-circuit: проверять `Nothing` **до** обращения к свойству отдельным условием.
- После `ActivateSheet` сверять `GetCurrentSheet`; после `ActivateDoc3` — полный путь ActiveDoc; после Paste — реальные объекты. Не полагаться лишь на один Boolean.
- Реализация: `drawingmerger/V16_21/tools/merge_core.vba`, `Require`, `EnsureSheetActive`, `RequireActive`.

## C2. API, тесты и статус
- Сигнатуры проверять по [официальной справке SolidWorks 2025](https://help.solidworks.com/2025/english/api/sldworksapi/Welcome.htm) и фактическому interop. Не выдумывать `SetBaseView`, `EditPaste`, `GetReadOnlyState`.
- Проверка ANTLR/grammar, interop 32.1.0, Python-тестов и ZIP означает только статическую проверку. Нужны `Debug → Compile` и запуск в SW2025 на **копии**.
- При ошибке сохранять этап, `Err.Number/Source/Description` **до** очистки. Проверять `SAFEARRAY`, `Nothing`, `GetOutline`, `Position` перед индексацией. Известные зарезервированные имена в VBA: `circle`, `scale`.
- Нативный механизм проверки не подменять отчётом JSON с прежним SHA.

## C3. Поставка
- Текстовый модуль для вставки — TXT; для импорта BAS — CP1251/CRLF с `Attribute VB_Name`. Не оставлять два `main`. Форма merger: `frmDrawingOrder` и `FORM_CODE.txt`.
- `.swp` собирать/сохранять в SolidWorks VBA Editor: source-only перестроенный OLE/SWP ранее **не открылся** у пользователя.
- Текущая сверловка R7+ — один VBA-модуль, без регистрации DLL, COM-моста и прав администратора. Прежний R6 — история.
- Не отключать Defender и не делать выводов о причине блокировки без точного детекта.

## C4. C# и AutoSizeDrilling
- Оригинал VB.NET DLL — неизменный; C# 1.1.2 с Geometry.cs и тестами — кандидат, не подтверждённый native-заменитель VSTA. VSTA-проект оригинала отсутствует.
- `RunForIntegration` — строгая историческая ветка R6; C# и VBA имеют разные правила bool/interop. См. [dimensions.md](dimensions.md).

## C5. Сверловка: важные ошибки
- `ReferencedDocument` **не** гарантирует `ActiveView`/открытое окно модели: вероятный путь Error91, но конкретная строка не доказана.
- Error457 у R3 был на этапе снимка видов до Paste. `GetUniqueName` тоже дал совпадения (R4). Не делать словарь по имени вида и не угадывать идентичность по индексу.
- Сравнение `Is` и `IsSame` требует **положительного** `Same`; `Unsupported` — не успех. Проверять сохранность всех прежних видов и результат частичной вставки.
- Удалённую COM-ссылку не разыменовывать; после удаления проверять число видов и выжившие объекты.
- Старый SWP включал другие пользовательские модули; при обновлении заменить только `InsertWeldWithHoles`, не удалять соседние без проверки.

**Минимальный контроль нового изменения:** API → грамматика/тесты → Windows Compile → native запуск на копии → проверка листов, размеров, связей, сохранения/PDF. При отсутствии последнего писать «Не проверено в SW». [Реестр](evidence.md), [виды](drawings.md), [инструкция сверловки](../../weldment_drilling/README_RU.md).
