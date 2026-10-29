# Компоненты, свойства листов и связанные заметки

## N1. Источник свойств листа — отдельно от компонента заметки

Задача: после объединения штамп и заметки получают свойства исходной модели,
а не основной сборки. Реализация: [merge_core.vba](../../drawingmerger/V16_21/tools/merge_core.vba)
`ApplySheetPropertySource`, `EffectivePropertyView`, `LocalPropertySourceView`,
`PropertyViewReadbackMatches`; [source_properties.vba](../../drawingmerger/V16_21/tools/source_properties.vba)
`LocalizeTargetPropertyView`. Статус: поведение V13/V14 частично проверено журналами,
позднее включено в подтверждённые V15/V21; условия — [evidence.md](evidence.md).

API: Sheet.GetProperties2/SetProperties2 (флаг наследования SameCustomProp),
CustomPropertyView, ModelDocExtension.GetUserPreferenceToggle
(swDrawingSheetsMatchCustomPropVals, swDetailingNoOptionSpecified) и
GetUserPreferenceInteger(swDrawingSheetCustomPropSheetNo, swDetailingNoOptionSpecified),
View.ReferencedConfiguration, GetReferencedModelName. Номер листа — 1..N;
PropertySheetNo — ключ снимка, не метод DrawingDoc. Снимать эффективный источник,
отключать нежелательное наследование, выбирать локальный эквивалент по модели
**и конфигурации**. Readback «По умолчанию»/Default допустим лишь если фактически
разрешённый локальный вид соответствует обеим исходным характеристикам.
Строковое совпадение имени вида недостаточно; произвольный другой именованный вид
не принимать. Источник с другого листа может быть отвергнут SW.

V17 добавила локальный скрытый helper-view через CreateDrawViewFromModelView3,
ReferencedConfiguration, SetVisible(False,False), HiddenViewsVisible=False,
масштаб 0.01. Он тоже входит в ожидаемый снимок, не заменяет видимые виды деталей
и не доказывает исправления $COMP-заметок. Границы этого метода — модель/конфигурация
и поддерживаемая исходная ориентация. Для дальнейшего создания видов есть API-
отправная точка, но нет универсально проверенного алгоритма размещения.

## N2. Не терять живую связь «Связать со свойством → Наименование»

Задача: выноска остаётся, но текст после переноса становится пустым. Статус:
V21 Подтверждено пользователем 28.10; offline fixtures Проверено автоматически.
Реализация: [linked_notes.vba](../../drawingmerger/V16_21/tools/linked_notes.vba),
`CaptureLinkedNotes`, `NoteAttachmentSignature`, `RemapComponentViews`,
`RestoreViewLinkedNotes`, `SectionComponentLink`, `VerifyLinkedNotes`.

Снимать `Note.PropertyLinkedText` вместе с GetText, координатами Annotation,
слоем, видимостью, GetAttachedEntities3 и количеством привязок. Текущий алгоритм
адресует только выражения, содержащие **оба** `$PRPSMODEL:` и `$COMP:"`.
Другие классы заметок не объявлять автоматически покрытыми этим ремонтом.

Для Entity сначала `GetDrawingComponent(view).Component`, затем fallback
GetComponent. Контекст разреза отличается от сырого assembly-контекста Entity.
Идентичность привязки: нормализованный путь модели + конфигурация + Name2;
сортировка сохраняет кратность. Сравнивать после установки PropertyLinkedText:
старая ссылка до ремонта могла разрешаться в другом контексте. Неизвестные
объекты/неоднозначное совпадение — остановка, а не выбор похожей детали.
Не заменять выражение значением GetText/SetText: это уничтожает связь со свойством.

Общий путь: remap только токенов @view/ внутри $COMP из исходной строки без
каскадных замен; сопоставить заметку однозначно по выражению и позиции,
восстановить слой/видимость. Финально проверить число, координаты, выражение,
привязки и непустой исходный текст. Допуск координат GEOM_TOL=1e-6 м.

## N3. Нативный адрес SectionAssembly

Условия специального ремонта V21: исходная ссылка содержит SectionAssembly,
исходный текст непуст, текущий пуст, идентичность привязки совпала; ровно одна
привязка и одно $COMP. Это узкое решение, не универсальный парсер всех формул.

`SectionComponentLink` берёт **DrawingComponent.Name прикреплённой Entity**, а не
RootDrawingComponent.Component.Name2: полная цепочка включает генерируемую SW
SectionAssembly с номерами экземпляров. Имя вида — **View.GetUniqueName**,
а не видимая подпись GetName2 «Разрез А-А». Строить:

```text
<SectionAssembly-instance>@<view-unique-name>/
<assembly-instance>@<parent-model-name>/<part-instance>@<parent-model-name>
```

В коде это одна строка. `InstanceModelName` снимает только последний проверенный
суффикс «-число»; не обрезать произвольные дефисы названия. Проверить identity,
точное чтение установленной формулы, сохранить ResolvedLink для ожидаемого снимка.
Разные разрезы могут иметь разные внутренние имена и разные экземпляры детали.

V20 **Не работает** в выявленном случае: root Name2 терял номер экземпляра,
а отображаемое имя разреза не являлось внутренним адресом; обычное переименование
токенов не обновляло генерируемую SectionAssembly. Выноска с пустым текстом могла
проходить проверки количества заметок. Диагностика указала эти различия; V21
восстановила native-цепочку, затем пользователь подтвердил результат.
Для смежных сборочных задач различать DrawingComponent и Component2 контексты;
это не алгоритм изменения сборки или универсальный persistent ID.

Диагностика без сохранения/rebuild: [linked_notes](../../diagnostics/linked_notes/README_RU.md)
и [linked_notes_context](../../diagnostics/linked_notes_context/README_RU.md).
TSV UTF-16 содержит личные пути/тексты — не публиковать. Предпочитать пару отчётов
исходник/результат вместо нового долгого объединения только ради диагностики.

Для геометрических баз и нанесения размеров на вид см. [dimensions.md](dimensions.md).
Контекст View.SelectEntity полезен и при повторном использовании одного ребра
модели в нескольких видах, но не заменяет проверку привязки заметок N2/N3.
