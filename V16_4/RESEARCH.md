> Текущая корректировка V16.4: FIX_RU.md. Далее исторические сведения.

> Текущая корректировка V16.3: FIX_RU.md. Далее исторические сведения.

> Текущая версия V16.2: см. FIX_RU.md. Далее исторические сведения.

> Историческое описание V16. Для текущей корректировки см. FIX_RU.md.

# Исследование переноса зависимых видов — 27.10.2026

## Причина дефекта

В V12 каждый лист копировался отдельно. На проблемном листе оба разреза ссылаются
через `GetBaseView` на `Чертежный вид4` другого листа. Сообщённый результат:
`PasteSheet=True`, но вместо двух видов на вставленном листе — ноль.
V14 обходил дефект пропуском разрезов; это не выполняет поставленную задачу.
V15 содержал эксперимент с `EditPaste`; у `IModelDoc2` действительный метод — `Paste`.

## Документация SOLIDWORKS

1. [Copying and Pasting Views](https://help.solidworks.com/2023/english/SolidWorks/sldworks/c_copy_and_paste_views.htm):
   для переноса section/detail между документами родитель копируется первым либо
   вместе с зависимым видом. Также описан перенос видов между листами одного документа.
2. [IDrawingDoc.PasteSheet](https://help.solidworks.com/2025/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IDrawingDoc~PasteSheet.html):
   выбирается лист, выполняется `EditCopy`, затем вставка. Возвращаемый Boolean сам
   по себе не подтверждает наличие всех видов и их ассоциативность.
3. [IModelDoc2.Paste](https://help.solidworks.com/2025/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDoc2~Paste.html)
   и [EditCut](https://help.solidworks.com/2019/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelDoc2~EditCut.html):
   используются для переноса нативного вида внутри документа.
4. [IView.GetBaseView](https://help.solidworks.com/2017/English/api/sldworksapi/SOLIDWORKS.Interop.sldworks~SOLIDWORKS.Interop.sldworks.IView~GetBaseView.html):
   чтение родительского вида; произвольный `SetBaseView` не предусмотрен.
5. [IDrawingDoc.GetViews](https://help.solidworks.com/2010/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IDrawingDoc~GetViews.html):
   порядок массивов листов не определён, первый элемент каждого — служебный вид.
   V16 не сопоставляет массивы с листами по индексу.
6. [IView.SetName2](https://help.solidworks.com/2025/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IView~SetName2.html):
   переименование section/detail этим методом не поддерживается. Поэтому V16
   не строит восстановление связей на переименовании разрезов.

## Первичные исходники GitHub

- [xarial/xcad — SwSheet.cs](https://github.com/xarial/xcad/blob/master/src/SolidWorks/Documents/SwSheet.cs):
  проверка `PasteSheet` и повтор вставки при отказе. В V16 повтор разрешён лишь
  если количество листов не изменилось, чтобы не продублировать частичную вставку.
- [xarial/codestack — propagate-configurations-sheets](https://github.com/xarial/codestack/blob/master/solidworks-api/document/drawing/propagate-configurations-sheets/Macro.vba):
  практическое использование копирования листов через API.
- [JayaVeeramani/GD-Macros — HideShowForm.frm](https://github.com/JayaVeeramani/GD-Macros/blob/master/CutListAndBeamDetails/HideShowForm.frm):
  `CreateSheetAndMoveDrawingViews` использует выбор видов, `EditCut` и вставку на
  другом листе одного чертежа. Это пример использования, не испытание нашего случая.

Готового доказанно работающего исправления именно описанного многолистового
случая в просмотренных исходниках не найдено.

## Принятое решение

Сбор семейства на одном листе перед `PasteSheet`, затем распределение внутри
целевого документа — инженерный вывод из документированных операций.
Это не готовый рецепт SOLIDWORKS и не результат нативного испытания V16.
Нативные разрезы не заменяются картинками или заново построенной геометрией.

Работа выполняется только на побайтовых копиях сохранённых файлов; исходный
документ открывается для чтения/разрешения моделей и сравнения копии. `SaveAs`
для создания копии источника не используется. Изменённые исходники не сохраняются.
Подлинные ссылки на модели и родительские виды сверяются до и после переноса.

Имена и число параметров API дополнительно проверены по метаданным оригинальных
NuGet-пакетов `SolidWorks.Interop.sldworks` и `SolidWorks.Interop.swconst` 32.1.0.
Это ABI SOLIDWORKS 2024; используемые операции сверены с документацией 2025.
Проверка DLL не заменяет компиляцию VBA и выполнение COM в SOLIDWORKS 2025.
