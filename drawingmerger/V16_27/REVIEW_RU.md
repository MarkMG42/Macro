# V16.27 — эксперимент UI на базе V16.25

Добавлен обратимый guard обновлений интерфейса по полному нормализованному пути документа. Реестр создаётся до записи флагов; запись сохраняет модель, доступные менеджеры, прежние значения и отдельные held-флаги. Перед контролируемым CloseDoc выполняется восстановление закрываемого документа и всех отслеживаемых недокументных моделей (IsDrawing=False), потому что закрытие чертежа может выгрузить ReferencedDocument. После подтверждённого восстановления запись удаляется. Неразрешённые ошибки оставляют запись для повторной попытки; все свойства/документы по возможности обрабатываются независимо.

Guard вызывается после проверенного RequireActive и при чтении уже загруженного ReferencedDocument. RestoreAllUI сначала выключает uiEnabled, поэтому финальный CaptureDrawing и cleanup-активации не включают подавление повторно. Состояния восстанавливаются до аудита/Save3/PDF, при ошибке и в Finished. Начало следующего запуска не стирает непогашенные записи. Таймер отражает UI_RESTORE, остальные метрики прежние. При отсутствии/недоступности менеджера оптимизация пропускается с явной записью; отсутствие ConfigurationManager у drawing не считается отказом макроса.

Сама геометрия и её обновление не отключаются. RebuildDrawing, AppendDrawing, CloneInput, PrepareSource, MapCopiedSheet, VerifyDrawing, VerifyStructure, PDF и логика источников взяты из V16.25. В CaptureDrawing/CaptureStructure добавлен только вызов guard после положительной проверки refDoc. Данные конфигураций и их выбор сохранены. V16.26 не является основой.

API-контракты проверены по первичным источникам и наличию соответствующих членов в доступном interop (это не native SW2025-тест):
- [FeatureManager properties](https://help.solidworks.com/2024/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IFeatureManager_properties.html): EnableFeatureTree и EnableFeatureTreeWindow, Boolean get/set.
- [EnableFeatureTreeWindow](https://help.solidworks.com/2023/english/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IFeatureManager~EnableFeatureTreeWindow.html): временная блокировка интерактивного устаревшего дерева, обязательное возвращение управления.
- [ConfigurationManager SW2025](https://help.solidworks.com/2025/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IConfigurationManager_members.html): EnableConfigurationTree.
- [EnableGraphicsUpdate SW2025](https://help.solidworks.com/2025/English/api/sldworksapi/SolidWorks.Interop.sldworks~SolidWorks.Interop.sldworks.IModelView~EnableGraphicsUpdate.html): обновление модельного окна.

Страницы отдельных свойств SW2025 частично недоступны в Web Help; контракт FeatureManager сверялся с более ранней официальной справкой и interop. VBA-парсер не доказывает работоспособность late-bound CallByName в реальном документе. Нужны Compile VBAProject, нормальный запуск, принудительно вызванный контролируемый отказ, проверки восстановления и тест выделения листов при выключенном окне FeatureManager.

Рабочие архивы V16.25 и V16.26 неизменны; рекомендуемая версия остаётся V16.25 до пользовательского испытания. Native-запуск и измерение ускорения V16.27 не выполнялись.
