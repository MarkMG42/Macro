Attribute VB_Name = "MergeDrawingsV16_21"
Option Explicit

Private Type OPENFILENAMEW
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    lpstrFilter As LongPtr
    lpstrCustomFilter As LongPtr
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As LongPtr
    nMaxFile As Long
    lpstrFileTitle As LongPtr
    nMaxFileTitle As Long
    lpstrInitialDir As LongPtr
    lpstrTitle As LongPtr
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As LongPtr
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As LongPtr
    pvReserved As LongPtr
    dwReserved As Long
    FlagsEx As Long
End Type

Private Declare PtrSafe Function GetOpenFileNameW _
    Lib "comdlg32.dll" (ByRef p As OPENFILENAMEW) As Long

Private Declare PtrSafe Function GetSaveFileNameW _
    Lib "comdlg32.dll" (ByRef p As OPENFILENAMEW) As Long


Private Const OFN_OVERWRITEPROMPT As Long = &H2
Private Const OFN_HIDEREADONLY As Long = &H4
Private Const OFN_PATHMUSTEXIST As Long = &H800
Private Const OFN_FILEMUSTEXIST As Long = &H1000
Private Const OFN_ALLOWMULTISELECT As Long = &H200
Private Const OFN_EXPLORER As Long = &H80000

Private Const INSERT_MOVE_TO_END As Long = 2
Private Const RENAME_DUPLICATES As Long = 1

Private swApp As SldWorks.SldWorks
Private targetModel As SldWorks.ModelDoc2
Private targetDraw As SldWorks.DrawingDoc
Private targetPath As String

' MERGE_DRAWINGS_V16_21_NATIVE_SECTION_ADDRESS_20261028
' SOLIDWORKS 2025 x64 / VBA7. Entry point: main.
' Source files are read-only inputs. Only owned result/working drawings are edited.
' Native section/detail features are transferred with their complete parent chain.
' A successful PasteSheet is never accepted without verifying the view graph.

Private Const BUILD_ID As String = "MERGE_DRAWINGS_V16_21_NATIVE_SECTION_ADDRESS_20261028"
Private Const VERIFY_AFTER_SAVE As Boolean = False ' True: slower reopen/rebuild verification
Private Const GEOM_TOL As Double = 0.000001 ' metres; 0.001 mm
Private Const VALUE_TOL As Double = 0.0000001
Private Const MAX_ITEMS As Long = 100000
Private runFolder As String
Private publishedDrawingPath As String
Private runToken As String
Private logStream As Object
Private fsoRun As Object
Private expectedResult As Object
Private runBusy As Boolean

Public Sub main()
    Dim mainPath As String, resultPath As String, pdfPath As String, resultPdfPath As String
    Dim sourceFiles As Collection, sourcePath As Variant
    Dim choice As VbMsgBoxResult, added As Long, fileIndex As Long
    Dim errors As Long, warnings As Long, saved As Boolean
    Dim actual As Object, failure As String, failureNumber As Long
    Dim verificationText As String
    Dim initialActive As SldWorks.ModelDoc2
    On Error GoTo Failed
    If runBusy Then Exit Sub
    runBusy = True
    Set swApp = Application.SldWorks
    Set initialActive = swApp.ActiveDoc
    Set targetModel = Nothing
    Set targetDraw = Nothing
    Set expectedResult = Nothing
    targetPath = ""
    runFolder = ""
    publishedDrawingPath = ""
    Set fsoRun = CreateObject("Scripting.FileSystemObject")

    mainPath = PickOneDrawing("Выберите ОСНОВНОЙ чертеж", "")
    If Len(mainPath) = 0 Then GoTo Finished
    CheckSavedInput mainPath
    InitRunFolder mainPath
    LogLine "BUILD_ID: " & BUILD_ID
    LogLine "SOLIDWORKS: " & swApp.RevisionNumber
    LogLine "MAIN: " & mainPath
    LogLine "OUTPUT folder: " & fsoRun.GetParentFolderName(mainPath)

    choice = MsgBox("Как выбрать исходные чертежи?" & vbCrLf & vbCrLf & _
        "ДА - все .slddrw из папки основного чертежа." & vbCrLf & _
        "НЕТ - выбрать вручную." & vbCrLf & "ОТМЕНА - выйти.", _
        vbYesNoCancel + vbQuestion, "Объединение V16")
    If choice = vbCancel Then GoTo Finished
    If choice = vbYes Then
        Set sourceFiles = GetAllDrawingsInFolder(GetFolder(mainPath), mainPath)
    Else
        Set sourceFiles = PickMultipleDrawings("Выберите ИСХОДНЫЕ чертежи", GetFolder(mainPath))
    End If
    If sourceFiles Is Nothing Then GoTo Finished
    Set sourceFiles = GetOrderedSourceFiles(sourceFiles, mainPath)
    If sourceFiles Is Nothing Then GoTo Finished
    If sourceFiles.Count = 0 Then GoTo Finished
    For Each sourcePath In sourceFiles
        CheckSavedInput CStr(sourcePath)
    Next sourcePath

    ' Allocate one new result file. Never edit or SaveAs the selected original.
    resultPath = fsoRun.BuildPath(fsoRun.GetParentFolderName(mainPath), "MERGED_" & runToken & ".slddrw")
    resultPdfPath = fsoRun.BuildPath(fsoRun.GetParentFolderName(mainPath), "MERGED_" & runToken & ".pdf")
    Require fsoRun.FileExists(resultPath) = False, "Итоговый файл уже существует: " & resultPath
    Require fsoRun.FileExists(resultPdfPath) = False, "Итоговый PDF уже существует: " & resultPdfPath
    publishedDrawingPath = resultPath
    targetPath = resultPath
    LogLine "MODE direct result; verify_after_save=" & CStr(VERIFY_AFTER_SAVE)
    Set targetModel = CloneInput(mainPath, targetPath, expectedResult, True)
    Set targetDraw = targetModel
    LogLine "MAIN sheets=" & expectedResult("Sheets").Count & _
        "; views=" & expectedResult("Views").Count

    For Each sourcePath In sourceFiles
        fileIndex = fileIndex + 1
        added = added + AppendDrawing(CStr(sourcePath), fileIndex, actual)
    Next sourcePath

    RequireActive targetModel
    ' Each main/imported sheet has already been rebuilt once and the last append
    ' verified the complete drawing. Do not force another full rebuild here.
    Require Not actual Is Nothing, "Нет проверенного снимка результата."
    VerifyDrawing expectedResult, actual, "FINAL before save (last verified snapshot)"
    targetModel.ClearSelection2 True
    LogLine "SAVE result begin: " & targetModel.GetPathName
    saved = targetModel.Save3(swSaveAsOptions_Silent, errors, warnings)
    LogLine "SAVE result: " & saved & "; errors=" & errors & "; warnings=" & warnings
    Require saved And errors = 0, "Не удалось сохранить итоговый чертеж: " & errors

    If VERIFY_AFTER_SAVE Then
        CloseOwnedDrawing targetModel
        Set targetDraw = Nothing
        Set targetModel = OpenWorkingDrawing(resultPath)
        Set targetDraw = targetModel
        RebuildDrawing targetModel
        Set actual = CaptureDrawing(targetModel)
        VerifyDrawing expectedResult, actual, "FINAL after reopen (strict)"
        LogLine "SAVE strict result begin: " & targetModel.GetPathName
        saved = targetModel.Save3(swSaveAsOptions_Silent, errors, warnings)
        LogLine "SAVE strict result: " & saved & "; errors=" & errors & "; warnings=" & warnings
        Require saved And errors = 0, "Не сохранён результат строгой проверки: " & errors
        verificationText = "Проверены виды, связи, заметки и повторное открытие."
    Else
        ' Save3 may update notes or other state: read and compare after saving.
        ' This is an in-memory check, not a persistence/reopen test.
        Set actual = CaptureDrawing(targetModel)
        VerifyDrawing expectedResult, actual, "FINAL after save (memory)"
        verificationText = "Проверены виды, связи и заметки в открытом документе после Save3. Повторное открытие не выполнялось."
    End If
    WriteManifest actual, runFolder & "\verified_views.tsv"
    pdfPath = runFolder & "\WORK_PDF.pdf"
    RequireActive targetModel
    targetModel.ClearSelection2 True
    saved = ExportDrawingToPDF(targetModel, pdfPath, errors, warnings)
    LogLine "PDF: " & saved & "; errors=" & errors & "; warnings=" & warnings
    Require saved And errors = 0, "Чертеж проверен, но экспорт PDF не выполнен. Код: " & errors
    Require fsoRun.FileExists(pdfPath), "API сообщил об экспорте, но PDF отсутствует."
    Require fsoRun.GetFile(pdfPath).Size > 0, "Получен пустой файл PDF."
    fsoRun.CopyFile pdfPath, resultPdfPath, False
    LogLine "RESULT drawing: " & resultPath
    LogLine "RESULT PDF: " & resultPdfPath
    LogLine "SUCCESS sheets=" & targetDraw.GetSheetCount & "; added=" & added
    WriteStatus "SUCCESS", verificationText
    MsgBox "Объединение завершено." & vbCrLf & _
        "Добавлено листов: " & added & vbCrLf & _
        "Чертеж:" & vbCrLf & resultPath & vbCrLf & _
        "PDF:" & vbCrLf & resultPdfPath & vbCrLf & _
        "Журнал:" & vbCrLf & runFolder, vbInformation, "V16"
    GoTo Finished

Failed:
    failureNumber = Err.Number
    failure = Err.Description
    ' Capture the actual error before calling anything that can reset Err.
    On Error Resume Next
    LogLine "FAILED " & failureNumber & ": " & failure
    WriteStatus "FAILED", failure
    SaveFailedCheckpoint targetModel, failure
    Set targetDraw = Nothing
    If Not initialActive Is Nothing Then RequireActive initialActive
    MsgBox "Объединение остановлено: " & failure & vbCrLf & _
        "Исходные чертежи не перезаписаны." & vbCrLf & _
        "Рабочий результат оставлен открытым; он НЕ прошёл проверку." & vbCrLf & _
        "Журнал: " & runFolder & "\merge.log", vbCritical, "V16: проверка не пройдена"
    On Error GoTo 0
Finished:
    On Error Resume Next
    If Not logStream Is Nothing Then logStream.Close
    Set logStream = Nothing
    runBusy = False
    On Error GoTo 0
End Sub

Private Function AppendDrawing(ByVal inputPath As String, ByVal fileIndex As Long, _
    ByRef validatedSnapshot As Object) As Long
    Dim source As SldWorks.ModelDoc2, drawing As SldWorks.DrawingDoc
    Dim original As Object, staged As Object, current As Object
    Dim viewMap As Object, sheetMap As Object, rec As Object, copied As Object
    Dim propertyMap As Object, helperNames As Object, helperKey As Variant
    Dim directSource As Boolean, sourceOpened As Boolean, sourceSheet As String
    Dim order As Collection, key As Variant, sheetKey As Variant, s As Object
    Dim newSheet As String, destSheet As String, name As String
    Dim view As SldWorks.View, dest As SldWorks.Sheet, item As Object
    Dim i As Long, n As Long, e As Long, message As String
    On Error GoTo Failed
    LogLine "SOURCE " & fileIndex & ": " & inputPath
    Set source = PrepareSource(inputPath, fileIndex, original, directSource, sourceOpened, sourceSheet)
    Set drawing = source
    Set order = DependencyOrder(original("Views"))
    Set viewMap = NewDictionary()
    Set sheetMap = NewDictionary()
    Set propertyMap = NewDictionary()
    Set helperNames = NewDictionary()

    ' Keep each root view on its original sheet. Bring its descendants to it.
    ' Ordinary families already on one sheet are left untouched.
    For Each key In order
        Set rec = original("Views")(CStr(key))
        Set view = FindView(source, CStr(key))
        Require Not view Is Nothing, "Потерян вид источника: " & CStr(key)
        If Not SameText(ViewSheetName(view), CStr(rec("Carrier"))) Then
            MoveView source, CStr(key), CStr(rec("Carrier")), CDbl(rec("Scale"))
        End If
    Next key
    If directSource Then
        Set staged = original
    Else
        RebuildDrawing source
        Set staged = CaptureDrawing(source)
    End If
    VerifyStructure original, staged, "SOURCE consolidated"
    ' Parent and child must now be on the SAME sheet before PasteSheet.
    For Each key In staged("Views").Keys
        Set rec = staged("Views")(CStr(key))
        If Len(CStr(rec("Parent"))) > 0 Then
            Set item = staged("Views")(CStr(rec("Parent")))
            Require SameText(CStr(rec("Sheet")), CStr(item("Sheet"))), _
                "Не удалось собрать родительскую цепочку: " & CStr(key)
        End If
    Next key

    ' Transfer complete sheets in ORIGINAL order. The temporarily empty sheets
    ' carry their original format, sheet notes, tables, scale and paper settings.
    For Each sheetKey In original("Order")
        newSheet = CopyWholeSheet(source, CStr(sheetKey))
        sheetMap.Add CStr(sheetKey), newSheet
        Set current = CaptureDrawing(targetModel, newSheet)
        MapCopiedSheet staged, CStr(sheetKey), current, newSheet, viewMap
    Next sheetKey
    Require viewMap.Count = original("Views").Count, "Не все виды сопоставлены после вставки."

    ' Move leaves first when spreading a family back across target sheets.
    ' This moves native features WITHIN a drawing, preserving their parent links.
    For i = order.Count To 1 Step -1
        key = CStr(order(i))
        Set rec = original("Views")(CStr(key))
        name = CStr(viewMap(CStr(key)))
        destSheet = CStr(sheetMap(CStr(rec("Sheet"))))
        Set view = FindView(targetModel, name)
        Require Not view Is Nothing, "Не найден вставленный вид: " & name
        If Not SameText(ViewSheetName(view), destSheet) Then
            MoveView targetModel, name, destSheet, CDbl(rec("Scale"))
        End If
    Next i

    ' Restore exact scale mode/position in parent-first order. No invented
    ' SetBaseView, no detached images, no recreated or flattened section geometry.
    For Each key In order
        Set rec = original("Views")(CStr(key))
        name = CStr(viewMap(CStr(key)))
        Set view = FindView(targetModel, name)
        Require Not view Is Nothing, "Не найден вид для восстановления: " & name
        RequireActive targetModel
        EnsureSheetActive targetModel, CStr(sheetMap(CStr(rec("Sheet")))), "RestoreView"
        RestoreView view, rec
    Next key
    For Each sheetKey In original("Order")
        Set s = original("Sheets")(CStr(sheetKey))
        destSheet = CStr(sheetMap(CStr(sheetKey)))
        ' Apply sheet settings in the verified active-sheet context.
        EnsureSheetActive targetModel, destSheet, "Restore CustomPropertyView"
        targetModel.ClearSelection2 True
        Set dest = targetDraw.GetCurrentSheet
        name = MappedPropertyView(original, CStr(sheetKey), viewMap)
        LogLine "PROPERTY VIEW before: sheet=" & destSheet & "; source=[" & _
            CStr(s("PropertyView")) & "]; requested=[" & name & "]; current=[" & dest.CustomPropertyView & "]"
        name = LocalizeTargetPropertyView(original, CStr(sheetKey), destSheet, name, helperNames)
        propertyMap.Add CStr(sheetKey), name
        ApplySheetPropertySource dest, name, False
        LogLine "PROPERTY VIEW readback: sheet=" & destSheet & "; requested=[" & _
            name & "]; actual=[" & dest.CustomPropertyView & "]"
    Next sheetKey
    RestoreImportedLinkedNotes original, sheetMap, viewMap
    RebuildDrawing targetModel, sheetMap

    ' Build the expected final drawing from the ORIGINAL snapshots, not from
    ' whatever happened to survive the paste. Only generated names are mapped.
    Set current = CaptureDrawing(targetModel)
    For Each sheetKey In original("Order")
        Set rec = CloneRecord(original("Sheets")(CStr(sheetKey)))
        name = CStr(sheetMap(CStr(sheetKey)))
        rec("Name") = name
        rec("PropertyView") = CStr(propertyMap(CStr(sheetKey)))
        Set rec("LinkedNotes") = MappedLinkedNotes(rec("LinkedNotes"), viewMap)
        ' Imported sheets must not inherit the main drawing's property model.
        rec("SameCustomProp") = False
        expectedResult("Sheets").Add name, rec
        expectedResult("Order").Add name
    Next sheetKey
    For Each key In original("Views").Keys
        Set rec = CloneRecord(original("Views")(CStr(key)))
        name = CStr(viewMap(CStr(key)))
        rec("Name") = name
        Set rec("LinkedNotes") = MappedLinkedNotes(rec("LinkedNotes"), viewMap)
        rec("Sheet") = CStr(sheetMap(CStr(rec("Sheet"))))
        If Len(CStr(rec("Parent"))) > 0 Then rec("Parent") = CStr(viewMap(CStr(rec("Parent"))))
        ' PasteSheet(rename=yes) legitimately changes conflicting section labels.
        ' Keep the native new label, but never accept a changed parent identity.
        Set copied = current("Views")(name)
        rec("Label") = CStr(copied("Label"))
        expectedResult("Views").Add name, rec
    Next key
    For Each helperKey In helperNames.Keys
        Set copied = current("Views")(CStr(helperKey))
        Require SameText(CStr(copied("Model")), CStr(helperNames(helperKey)("PropertyModel"))), "Изменилась модель служебного вида."
        Require CStr(copied("Configuration")) = CStr(helperNames(helperKey)("PropertyConfiguration")), "Изменилась конфигурация служебного вида."
        Require CBool(copied("Visible")) = False, "Служебный вид должен быть скрыт."
        Require Len(CStr(copied("Parent"))) = 0, "У служебного вида появился родитель."
        expectedResult("Views").Add CStr(helperKey), CloneRecord(copied)
    Next helperKey
    VerifyDrawing expectedResult, current, "AFTER source " & fileIndex
    Set validatedSnapshot = current
    n = original("Sheets").Count
    ReleaseSource source, directSource, sourceOpened, sourceSheet
    AppendDrawing = n
    Exit Function
Failed:
    e = Err.Number
    message = Err.Description
    On Error Resume Next
    ReleaseSource source, directSource, sourceOpened, sourceSheet
    On Error GoTo 0
    Err.Raise e, "AppendDrawing", inputPath & vbCrLf & message
End Function

Private Sub ApplySheetPropertySource(ByVal sheet As SldWorks.Sheet, _
    ByVal viewName As String, ByVal sameAsDocument As Boolean)
    Dim props As Variant, afterProps As Variant
    props = sheet.GetProperties2
    Require IsArray(props), "Не прочитаны свойства листа " & sheet.GetName
    Require UBound(props) >= 7, "Не прочитан режим свойств листа " & sheet.GetName
    LogLine "PROPERTY MODE before: sheet=" & sheet.GetName & "; inherited=" & _
        CStr(FlagValue(CBool(props(7)))) & "; requested=" & CStr(FlagValue(sameAsDocument))
    ' The inheritance flag controls whether CustomPropertyView can take effect.
    ' Keep all other paper/template/scale settings unchanged.
    sheet.SetProperties2 CLng(props(0)), CLng(props(1)), CDbl(props(2)), CDbl(props(3)), _
        FlagValue(CBool(props(4))), CDbl(props(5)), CDbl(props(6)), FlagValue(sameAsDocument)
    afterProps = sheet.GetProperties2
    Require IsArray(afterProps), "Не прочитан результат установки режима свойств листа."
    Require UBound(afterProps) >= 7, "Нет флага режима свойств после установки."
    Require FlagValue(CBool(afterProps(7))) = FlagValue(sameAsDocument), _
        "Не восстановлен режим источника свойств листа " & sheet.GetName
    If sameAsDocument = False Then sheet.CustomPropertyView = viewName
    LogLine "PROPERTY MODE readback: sheet=" & sheet.GetName & "; inherited=" & _
        CStr(FlagValue(CBool(afterProps(7)))) & "; view=[" & sheet.CustomPropertyView & "]"
End Sub

Private Function LocalPropertySourceView(ByVal snapshot As Object, ByVal sheetName As String) As String
    Dim sheet As Object, view As Object, key As Variant, name As String
    Set sheet = snapshot("Sheets")(sheetName)
    name = CStr(sheet("EffectivePropertyView"))
    Require Len(name) > 0, "Нет фактического вида-источника свойств листа " & sheetName
    Require snapshot("Views").Exists(name), "Потерян вид-источник свойств " & name
    Set view = snapshot("Views")(name)
    ' Prefer the original source when it belongs to this sheet.
    If SameText(CStr(view("Sheet")), sheetName) Then
        LocalPropertySourceView = name
        Exit Function
    End If
    ' A different local view is equivalent only for the exact same model/configuration.
    For Each key In snapshot("Views").Keys
        Set view = snapshot("Views")(CStr(key))
        If SameText(CStr(view("Sheet")), sheetName) And _
            SameText(CStr(view("Model")), CStr(sheet("PropertyModel"))) And _
            CStr(view("Configuration")) = CStr(sheet("PropertyConfiguration")) Then
            LocalPropertySourceView = CStr(key)
            Exit Function
        End If
    Next key
    ' Some sheets have no local view of the property model. Keep the original source;
    ' the normal readback and post-rebuild checks must confirm cross-sheet support.
    LogLine "PROPERTY SOURCE cross-sheet required: sheet=" & sheetName & "; view=" & name
    LocalPropertySourceView = name
End Function

Private Function MappedPropertyView(ByVal snapshot As Object, ByVal sheetName As String, _
    ByVal viewMap As Object) As String
    Dim rec As Object, name As String, originalMode As String
    Set rec = snapshot("Sheets")(sheetName)
    name = CStr(rec("PropertyView"))
    originalMode = name
    ' Freeze the original effective source before joining documents with different defaults.
    If CBool(rec("SameCustomProp")) Then
        name = LocalPropertySourceView(snapshot, sheetName)
        Require viewMap.Exists(name), "Не сопоставлен источник свойств листа " & sheetName & ": " & name
        MappedPropertyView = CStr(viewMap(name))
        LogLine "PROPERTY SOURCE localize: sourceSheet=" & sheetName & "; sourceView=[" & name & _
            "]; mappedView=[" & MappedPropertyView & "]; inherited=False"
        Exit Function
    End If
    ' A real view name takes precedence over a localized special label.
    If Not viewMap.Exists(name) Then
        If SameText(name, "По умолчанию") Or SameText(name, "Default") Then
            name = CStr(rec("FirstView"))
            Require Len(name) > 0, "Нет первого вида для источника свойств листа " & sheetName
            Require viewMap.Exists(name), "Не сопоставлен первый вид свойств листа " & sheetName & ": " & name
            Require SameText(CStr(snapshot("Views")(name)("Sheet")), sheetName), _
                "Первый вид свойств относится к другому листу: " & name
            LogLine "PROPERTY VIEW resolve default: sourceSheet=" & sheetName & _
                "; mode=[" & originalMode & "]; sourceView=[" & name & "]; mappedView=[" & CStr(viewMap(name)) & "]"
        End If
    End If
    If viewMap.Exists(name) Then name = CStr(viewMap(name))
    MappedPropertyView = name
End Function

Private Function EffectivePropertyView(ByVal snapshot As Object, ByVal sheetName As String) As String
    Dim rec As Object, sourceSheet As String, name As String, sheetNo As Long
    Set rec = snapshot("Sheets")(sheetName)
    sourceSheet = sheetName
    If CBool(rec("SameCustomProp")) And CBool(snapshot("MatchPropertySheets")) Then
        sheetNo = CLng(snapshot("PropertySheetNo"))
        Require sheetNo >= 1 And sheetNo <= snapshot("Order").Count, _
            "Некорректный номер листа свойств документа: " & sheetNo
        sourceSheet = CStr(snapshot("Order")(sheetNo))
        Set rec = snapshot("Sheets")(sourceSheet)
    End If
    name = CStr(rec("PropertyView"))
    If snapshot("Views").Exists(name) = False Then
        Select Case LCase$(Trim$(name))
            Case "", "default", "по умолчанию", _
                "совпадает с листом, указанным в свойствах документа", _
                "same as sheet specified in document properties"
                name = CStr(rec("FirstView"))
            Case Else
                Require False, "Не распознан источник свойств листа " & sourceSheet & ": [" & name & "]"
        End Select
    End If
    If Len(name) > 0 Then
        Require snapshot("Views").Exists(name), "Потерян фактический вид свойств: " & name
    End If
    EffectivePropertyView = name
End Function

Private Function CaptureDrawing(ByVal model As SldWorks.ModelDoc2, Optional ByVal onlySheet As String = "") As Object
    Dim result As Object, sheets As Object, views As Object, rec As Object, sheetRec As Object
    Dim order As Collection, names As Variant, props As Variant, p As Variant, bounds As Variant
    Dim drawing As SldWorks.DrawingDoc, sheet As SldWorks.Sheet
    Dim view As SldWorks.View, parent As SldWorks.View, refDoc As SldWorks.ModelDoc2
    Dim section As SldWorks.DrSection
    Dim key As Variant, i As Long, guard As Long, activeSheet As String
    RequireActive model
    Set drawing = model
    Set sheet = drawing.GetCurrentSheet
    activeSheet = sheet.GetName
    names = drawing.GetSheetNames
    Require IsArray(names), "Не получен список листов: " & model.GetTitle
    Set result = NewDictionary()
    result.Add "MatchPropertySheets", FlagValue(model.Extension.GetUserPreferenceToggle( _
        swDrawingSheetsMatchCustomPropVals, swDetailingNoOptionSpecified))
    result.Add "PropertySheetNo", model.Extension.GetUserPreferenceInteger( _
        swDrawingSheetCustomPropSheetNo, swDetailingNoOptionSpecified)
    LogLine "PROPERTY DOCUMENT: doc=" & model.GetPathName & "; match=" & _
        CStr(result("MatchPropertySheets")) & "; sheetNo=" & CStr(result("PropertySheetNo"))
    Set sheets = NewDictionary()
    Set views = NewDictionary()
    Set order = New Collection
    For i = LBound(names) To UBound(names)
        If Len(onlySheet) > 0 Then
            If Not SameText(CStr(names(i)), onlySheet) Then GoTo NextCaptureSheet
        End If
        EnsureSheetActive model, CStr(names(i)), "CaptureDrawing"
        drawing.EditSheet
        Set sheet = drawing.GetCurrentSheet
        Require SameText(sheet.GetName, CStr(names(i))), "Активирован неверный лист."
        props = sheet.GetProperties2
        Require IsArray(props), "Не прочитаны параметры листа " & sheet.GetName
        Require UBound(props) >= 7, "Не получен флаг источника свойств листа " & sheet.GetName
        Set sheetRec = NewDictionary()
        sheetRec.Add "Name", sheet.GetName
        sheetRec.Add "Width", CDbl(props(5))
        sheetRec.Add "Height", CDbl(props(6))
        sheetRec.Add "Scale1", CDbl(props(2))
        sheetRec.Add "Scale2", CDbl(props(3))
        sheetRec.Add "FirstAngle", FlagValue(CBool(props(4)))
        sheetRec.Add "PropertyView", sheet.CustomPropertyView
        sheetRec.Add "SameCustomProp", FlagValue(CBool(props(7)))
        sheetRec.Add "FormatVisible", FlagValue(sheet.SheetFormatVisible)
        Set view = drawing.GetFirstView
        Require Not view Is Nothing, "Не получен служебный вид листа " & sheet.GetName
        Require view.Type = swDrawingSheet, "GetFirstView не является служебным видом листа."
        AddAnnotationStats sheetRec, view
        sheets.Add sheet.GetName, sheetRec
        order.Add sheet.GetName
        Set view = view.GetNextView ' Exclude the sheet pseudo-view exactly once.
        sheetRec.Add "FirstView", ""
        If Not view Is Nothing Then sheetRec("FirstView") = view.GetName2
        Do While Not view Is Nothing
            guard = guard + 1
            Require guard <= MAX_ITEMS, "Зацикливание обхода видов."
            Set rec = NewDictionary()
            rec.Add "Name", view.GetName2
            rec.Add "Sheet", sheet.GetName
            rec.Add "Type", CLng(view.Type)
            rec.Add "Model", NormalizePath(view.GetReferencedModelName)
            rec.Add "Configuration", view.ReferencedConfiguration
            If Len(CStr(rec("Model"))) > 0 Then
                Set refDoc = view.ReferencedDocument
                Require Not refDoc Is Nothing, "Не загружена модель вида " & view.GetName2
                rec("Model") = NormalizePath(refDoc.GetPathName)
            End If
            Set parent = view.GetBaseView
            rec.Add "Parent", ""
            If Not parent Is Nothing Then rec("Parent") = parent.GetName2
            If view.Type = swDrawingSectionView Or view.Type = swDrawingDetailView Then
                Require Not parent Is Nothing, "У разреза/выносного вида отсутствует родитель: " & view.GetName2
            End If
            p = view.Position
            bounds = view.GetOutline
            Require IsArray(p) And IsArray(bounds), "Не получена геометрия вида " & view.GetName2
            rec.Add "X", CDbl(p(0))
            rec.Add "Y", CDbl(p(1))
            rec.Add "BoxWidth", CDbl(bounds(2)) - CDbl(bounds(0))
            rec.Add "BoxHeight", CDbl(bounds(3)) - CDbl(bounds(1))
            rec.Add "Scale", CDbl(view.ScaleDecimal)
            rec.Add "SheetScale", CLng(view.UseSheetScale)
            rec.Add "ParentScale", FlagValue(view.UseParentScale)
            rec.Add "Angle", CDbl(view.Angle)
            rec.Add "Orientation", view.GetOrientationName
            rec.Add "Locked", FlagValue(view.PositionLocked)
            rec.Add "Alignment", CLng(view.GetAlignment)
            rec.Add "Suppressed", CLng(view.SuppressState)
            rec.Add "Visible", FlagValue(view.GetVisible)
            rec.Add "Label", ""
            rec.Add "Reversed", False
            rec.Add "Partial", False
            rec.Add "SurfaceOnly", False
            rec.Add "Hatches", CLng(view.GetFaceHatchCount)
            If view.Type = swDrawingSectionView Then
                Set section = view.GetSection
                Require Not section Is Nothing, "Не прочитан объект разреза " & view.GetName2
                rec("Label") = section.GetLabel
                rec("Reversed") = FlagValue(section.GetReversedCutDirection)
                rec("Partial") = FlagValue(section.GetPartialSection)
                rec("SurfaceOnly") = FlagValue(section.GetDisplayOnlySurfaceCut)
            End If
            AddAnnotationStats rec, view
            Require Not views.Exists(CStr(rec("Name"))), "Неоднозначное имя вида: " & CStr(rec("Name"))
            views.Add CStr(rec("Name")), rec
            Set view = view.GetNextView
        Loop
NextCaptureSheet:
    Next i
    For Each key In views.Keys
        Set rec = views(CStr(key))
        If Len(CStr(rec("Parent"))) > 0 Then
            Require views.Exists(CStr(rec("Parent"))), "Родитель отсутствует в документе: " & CStr(key)
        End If
    Next key
    result.Add "Sheets", sheets
    result.Add "Views", views
    result.Add "Order", order
    For Each key In sheets.Keys
        Set sheetRec = sheets(CStr(key))
        sheetRec.Add "EffectivePropertyView", ""
        If Len(onlySheet) = 0 Then sheetRec("EffectivePropertyView") = EffectivePropertyView(result, CStr(key))
        sheetRec.Add "PropertyModel", ""
        sheetRec.Add "PropertyConfiguration", ""
        If Len(CStr(sheetRec("EffectivePropertyView"))) > 0 Then
            Set rec = views(CStr(sheetRec("EffectivePropertyView")))
            sheetRec("PropertyModel") = CStr(rec("Model"))
            sheetRec("PropertyConfiguration") = CStr(rec("Configuration"))
        End If
    Next key
    EnsureSheetActive model, activeSheet, "CaptureDrawing restore"
    Set CaptureDrawing = result
End Function

Private Sub AddAnnotationStats(ByVal rec As Object, ByVal view As SldWorks.View)
    Dim annotation As SldWorks.Annotation, dangling As Long, guard As Long
    rec.Add "Dimensions", CLng(view.GetDimensionCount4)
    rec.Add "Annotations", CLng(view.GetAnnotationCount)
    Set annotation = view.GetFirstAnnotation3
    Do While Not annotation Is Nothing
        guard = guard + 1
        Require guard <= MAX_ITEMS, "Зацикливание обхода аннотаций."
        If annotation.IsDangling Then dangling = dangling + 1
        Set annotation = annotation.GetNext3
    Loop
    rec.Add "Dangling", dangling
    rec.Add "LinkedNotes", CaptureLinkedNotes(view)
End Sub

Private Function DependencyOrder(ByVal views As Object) As Collection
    Dim result As Collection, key As Variant, node As Object, parent As String
    Dim rec As Object, depth As Long, maxDepth As Long, level As Long
    Set result = New Collection
    For Each key In views.Keys
        Set rec = views(CStr(key))
        Set node = rec
        depth = 0
        parent = CStr(node("Parent"))
        Do While Len(parent) > 0
            depth = depth + 1
            Require depth <= views.Count, "Цикл в родительских связях: " & CStr(key)
            Require views.Exists(parent), "Не найден родитель " & parent
            Set node = views(parent)
            parent = CStr(node("Parent"))
        Loop
        rec("Depth") = depth
        rec("Carrier") = CStr(node("Sheet"))
        If depth > maxDepth Then maxDepth = depth
    Next key
    For level = 0 To maxDepth
        For Each key In views.Keys
            Set rec = views(CStr(key))
            If CLng(rec("Depth")) = level Then result.Add CStr(key)
        Next key
    Next level
    Set DependencyOrder = result
End Function

Private Sub MoveView(ByVal model As SldWorks.ModelDoc2, ByVal viewName As String, _
    ByVal sheetName As String, ByVal desiredScale As Double)
    Dim drawing As SldWorks.DrawingDoc, view As SldWorks.View
    Dim fromSheet As String, before As Object, after As Object, props As Variant
    Dim p As Variant, locked As Boolean
    RequireOwnedPath model.GetPathName
    RequireActive model
    Set drawing = model
    Set view = FindView(model, viewName)
    Require Not view Is Nothing, "Не найден переносимый вид " & viewName
    fromSheet = ViewSheetName(view)
    If SameText(fromSheet, sheetName) Then Exit Sub
    Set before = CaptureDrawing(model)
    EnsureSheetActive model, fromSheet, "MoveView source"
    Set view = FindView(model, viewName)
    locked = FlagValue(view.PositionLocked)
    p = view.Position
    view.PositionLocked = False
    ' A child temporarily on a sheet with another scale must not change size.
    view.UseParentScale = False
    view.UseSheetScale = 0
    view.ScaleDecimal = desiredScale
    SelectViewForMove model, view
    model.EditCut
    EnsureSheetActive model, sheetName, "MoveView destination"
    drawing.EditSheet
    model.ClearSelection2 True
    ' Paste a drawing view into the verified active sheet, with no selection.
    LogLine "PASTE VIEW begin: view=" & viewName & "; destination=" & drawing.GetCurrentSheet.GetName
    model.Paste
    LogLine "PASTE VIEW returned: view=" & viewName
    Set view = FindView(model, viewName)
    Require Not view Is Nothing, "Вырезанный вид не вставлен: " & viewName
    Require SameText(ViewSheetName(view), sheetName), "Вид вставлен не на тот лист: " & viewName
    ' Native paste may use the current cursor location. Retain coordinates so
    ' different sections are not inadvertently stacked at the paste location.
    view.PositionLocked = False
    view.Position = p
    view.PositionLocked = locked
    Set after = CaptureDrawing(model)
    VerifyStructure before, after, "MOVE " & viewName
    LogLine "MOVE [" & viewName & "] " & fromSheet & " -> " & sheetName
End Sub

Private Sub SelectViewForMove(ByVal model As SldWorks.ModelDoc2, ByVal view As SldWorks.View)
    Dim objects(0) As Object, count As Long, byName As Boolean
    Dim viewName As String, sheetName As String
    viewName = view.GetName2
    sheetName = ViewSheetName(view)
    EnsureSheetActive model, sheetName, "SelectViewForMove"
    model.ClearSelection2 True
    Set objects(0) = view
    count = model.Extension.MultiSelect2(objects, False, Nothing)
    LogLine "SELECT VIEW object: name=" & viewName & "; sheet=" & sheetName & "; count=" & count
    If SelectedViewMatches(model, viewName, sheetName) Then Exit Sub
    ' Compatibility fallback; never cut solely on a Boolean selection result.
    model.ClearSelection2 True
    byName = model.Extension.SelectByID2(viewName, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
    LogLine "SELECT VIEW name: name=" & viewName & "; api=" & CStr(byName)
    If SelectedViewMatches(model, viewName, sheetName) Then Exit Sub
    model.ClearSelection2 True
    Require False, "Не выбран переносимый вид " & viewName & "; лист=" & sheetName & _
        "; MultiSelect2=" & count & "; SelectByID2=" & CStr(byName)
End Sub

Private Function SelectedViewMatches(ByVal model As SldWorks.ModelDoc2, _
    ByVal viewName As String, ByVal sheetName As String) As Boolean
    Dim manager As SldWorks.SelectionMgr, selectedObject As Object
    Dim selectedView As SldWorks.View, count As Long, objectType As Long
    Set manager = model.SelectionManager
    count = manager.GetSelectedObjectCount2(-1)
    LogLine "SELECT CHECK count=" & count
    If count <> 1 Then Exit Function
    objectType = manager.GetSelectedObjectType3(1, -1)
    LogLine "SELECT CHECK type=" & objectType
    Set selectedObject = manager.GetSelectedObject6(1, -1)
    If selectedObject Is Nothing Then Exit Function
    If Not TypeOf selectedObject Is SldWorks.View Then Exit Function
    Set selectedView = selectedObject
    LogLine "SELECT CHECK view=" & selectedView.GetName2 & "; sheet=" & ViewSheetName(selectedView)
    SelectedViewMatches = SameText(selectedView.GetName2, viewName) And _
        SameText(ViewSheetName(selectedView), sheetName)
End Function

Private Sub RestoreView(ByVal view As SldWorks.View, ByVal rec As Object)
    Dim p(1) As Double, actual As Variant
    view.PositionLocked = False
    ' Reapply scale only when necessary: avoid perturbing ordinary pasted views.
    If FlagValue(view.UseParentScale) <> CBool(rec("ParentScale")) Or _
       view.UseSheetScale <> CLng(rec("SheetScale")) Or _
       Not NearValue(view.ScaleDecimal, CDbl(rec("Scale")), VALUE_TOL) Then
        view.UseParentScale = False
        view.UseSheetScale = 0
        view.ScaleDecimal = CDbl(rec("Scale"))
        If CBool(rec("ParentScale")) Then
            view.UseParentScale = True
        ElseIf CLng(rec("SheetScale")) <> 0 Then
            view.UseSheetScale = 1
        End If
    End If
    If Not NearValue(view.Angle, CDbl(rec("Angle")), VALUE_TOL) Then view.Angle = CDbl(rec("Angle"))
    p(0) = CDbl(rec("X"))
    p(1) = CDbl(rec("Y"))
    actual = view.Position
    If Not NearValue(CDbl(actual(0)), p(0), GEOM_TOL) Or _
       Not NearValue(CDbl(actual(1)), p(1), GEOM_TOL) Then view.Position = p
    view.PositionLocked = CBool(rec("Locked"))
    ' Alignment is never removed to force a view into place. The final audit
    ' checks the original position AND alignment flags; an incompatible move fails.
End Sub

Private Sub SelectSheetForCopy(ByVal model As SldWorks.ModelDoc2, ByVal sheet As SldWorks.Sheet)
    Dim objects(0) As Object, count As Long, ok As Boolean
    Dim sheetName As String, props As Variant, x As Double, y As Double
    sheetName = sheet.GetName
    EnsureSheetActive model, sheetName, "SelectSheetForCopy"
    model.ClearSelection2 True
    Set objects(0) = sheet
    count = model.Extension.MultiSelect2(objects, False, Nothing)
    LogLine "SELECT SHEET object: sheet=" & sheetName & "; count=" & count
    If SelectedSheetMatches(model, sheetName) Then Exit Sub
    ' Preserve the previous selection routes, but verify their actual result.
    props = sheet.GetProperties2
    x = CDbl(props(5)) / 2#
    y = CDbl(props(6)) / 2#
    model.ClearSelection2 True
    ok = model.Extension.SelectByID2(sheetName, "SHEET", x, y, 0, False, 0, Nothing, 0)
    LogLine "SELECT SHEET name: api=" & CStr(ok)
    If SelectedSheetMatches(model, sheetName) Then Exit Sub
    model.ClearSelection2 True
    ok = model.Extension.SelectByID2("", "SHEET", x, y, 0, False, 0, Nothing, 0)
    LogLine "SELECT SHEET centre: api=" & CStr(ok)
    If SelectedSheetMatches(model, sheetName) Then Exit Sub
    model.ClearSelection2 True
    ok = model.Extension.SelectByID2("", "SHEET", 0, 0, 0, False, 0, Nothing, 0)
    LogLine "SELECT SHEET origin: api=" & CStr(ok)
    If SelectedSheetMatches(model, sheetName) Then Exit Sub
    model.ClearSelection2 True
    Require False, "Не выбран лист для копирования " & sheetName
End Sub

Private Function SelectedSheetMatches(ByVal model As SldWorks.ModelDoc2, ByVal sheetName As String) As Boolean
    Dim manager As SldWorks.SelectionMgr, selectedObject As Object
    Dim selectedSheet As SldWorks.Sheet, count As Long
    Set manager = model.SelectionManager
    count = manager.GetSelectedObjectCount2(-1)
    LogLine "SELECT SHEET CHECK count=" & count
    If count <> 1 Then Exit Function
    LogLine "SELECT SHEET CHECK type=" & manager.GetSelectedObjectType3(1, -1)
    Set selectedObject = manager.GetSelectedObject6(1, -1)
    If selectedObject Is Nothing Then Exit Function
    If Not TypeOf selectedObject Is SldWorks.Sheet Then Exit Function
    Set selectedSheet = selectedObject
    LogLine "SELECT SHEET CHECK name=" & selectedSheet.GetName
    SelectedSheetMatches = SameText(selectedSheet.GetName, sheetName)
End Function

Private Function CopyWholeSheet(ByVal source As SldWorks.ModelDoc2, ByVal sheetName As String) As String
    Dim drawing As SldWorks.DrawingDoc, sheet As SldWorks.Sheet
    Dim oldNames As Object, names As Variant, key As Variant
    Dim countBefore As Long, countAfter As Long, attempt As Long, found As Long
    Dim ok As Boolean
    Set drawing = source
    RequireActive source
    EnsureSheetActive source, sheetName, "CopyWholeSheet"
    drawing.EditSheet
    Set sheet = drawing.GetCurrentSheet
    Set oldNames = NewDictionary()
    names = targetDraw.GetSheetNames
    For Each key In names
        oldNames.Add CStr(key), True
    Next key
    countBefore = targetDraw.GetSheetCount
    SelectSheetForCopy source, sheet
    source.EditCopy
    ' Nothing that changes the source selection occurs between select and copy.
    RequireActive targetModel
    Set targetDraw = targetModel
    targetModel.ClearSelection2 True
    For attempt = 1 To 3
        ok = targetDraw.PasteSheet(INSERT_MOVE_TO_END, RENAME_DUPLICATES)
        countAfter = targetDraw.GetSheetCount
        If ok Then Exit For
        ' Retrying after a partial mutation could duplicate sheets. Never do that.
        Require countAfter = countBefore, "PasteSheet вернул False, но изменил число листов."
    Next attempt
    Require ok And countAfter = countBefore + 1, "PasteSheet не добавил ровно один лист: " & sheetName
    names = targetDraw.GetSheetNames
    For Each key In names
        If Not oldNames.Exists(CStr(key)) Then
            CopyWholeSheet = CStr(key)
            found = found + 1
        End If
    Next key
    Require found = 1, "Нельзя однозначно определить вставленный лист."
    Require SameText(CStr(names(UBound(names))), CopyWholeSheet), "Новый лист не добавлен в конец."
    LogLine "SHEET [" & sheetName & "] -> [" & CopyWholeSheet & "]"
End Function

Private Sub MapCopiedSheet(ByVal source As Object, ByVal sourceSheet As String, _
    ByVal target As Object, ByVal targetSheet As String, ByVal mapping As Object)
    Dim order As Collection, used As Object, rec As Object, candidate As Object
    Dim key As Variant, candidateKey As Variant, mappedParent As String
    Dim match As String, count As Long, expectedCount As Long, actualCount As Long
    Set order = DependencyOrder(source("Views"))
    Set used = NewDictionary()
    For Each candidateKey In target("Views").Keys
        Set candidate = target("Views")(CStr(candidateKey))
        If SameText(CStr(candidate("Sheet")), targetSheet) Then actualCount = actualCount + 1
    Next candidateKey
    For Each key In order
        Set rec = source("Views")(CStr(key))
        If SameText(CStr(rec("Sheet")), sourceSheet) Then
            expectedCount = expectedCount + 1
            mappedParent = ""
            If Len(CStr(rec("Parent"))) > 0 Then
                Require mapping.Exists(CStr(rec("Parent"))), "Родитель не сопоставлен до дочернего вида."
                mappedParent = CStr(mapping(CStr(rec("Parent"))))
            End If
            count = 0
            match = ""
            For Each candidateKey In target("Views").Keys
                Set candidate = target("Views")(CStr(candidateKey))
                If SameText(CStr(candidate("Sheet")), targetSheet) And Not used.Exists(CStr(candidateKey)) Then
                    If SameText(CStr(candidate("Parent")), mappedParent) Then
                        If SameViewFingerprint(rec, candidate) Then
                            count = count + 1
                            match = CStr(candidateKey)
                        End If
                    End If
                End If
            Next candidateKey
            Require count = 1, "Неоднозначный или потерянный вид [" & CStr(key) & _
                "]; совпадений=" & count & "; лист=" & targetSheet
            mapping.Add CStr(key), match
            used.Add match, True
            LogLine "MAP [" & CStr(key) & "] -> [" & match & "]; parent=[" & mappedParent & "]"
        End If
    Next key
    Require actualCount = expectedCount, "Изменилось число видов на листе " & targetSheet & _
        ": ожидалось=" & expectedCount & "; получено=" & actualCount
End Sub

Private Function SameViewFingerprint(ByVal a As Object, ByVal b As Object) As Boolean
    SameViewFingerprint = False
    If CLng(a("Type")) <> CLng(b("Type")) Then Exit Function
    If Not SameText(CStr(a("Model")), CStr(b("Model"))) Then Exit Function
    If StrComp(CStr(a("Configuration")), CStr(b("Configuration")), vbBinaryCompare) <> 0 Then Exit Function
    If CStr(a("Orientation")) <> CStr(b("Orientation")) Then Exit Function
    If Not NearValue(CDbl(a("X")), CDbl(b("X")), GEOM_TOL) Then Exit Function
    If Not NearValue(CDbl(a("Y")), CDbl(b("Y")), GEOM_TOL) Then Exit Function
    If Not NearValue(CDbl(a("Scale")), CDbl(b("Scale")), VALUE_TOL) Then Exit Function
    If Not NearValue(CDbl(a("Angle")), CDbl(b("Angle")), VALUE_TOL) Then Exit Function
    If CLng(a("Dimensions")) <> CLng(b("Dimensions")) Then Exit Function
    If CBool(a("Reversed")) <> CBool(b("Reversed")) Then Exit Function
    If CBool(a("Partial")) <> CBool(b("Partial")) Then Exit Function
    If CBool(a("SurfaceOnly")) <> CBool(b("SurfaceOnly")) Then Exit Function
    SameViewFingerprint = True
End Function

Private Sub VerifyStructure(ByVal expected As Object, ByVal actual As Object, ByVal stage As String)
    Dim key As Variant, a As Object, b As Object
    Require expected("Views").Count = actual("Views").Count, stage & ": изменилось общее число видов."
    For Each key In expected("Views").Keys
        Require actual("Views").Exists(CStr(key)), stage & ": потерян вид " & CStr(key)
        Set a = expected("Views")(CStr(key))
        Set b = actual("Views")(CStr(key))
        Require CLng(a("Type")) = CLng(b("Type")), stage & ": изменился тип " & CStr(key)
        Require SameText(CStr(a("Parent")), CStr(b("Parent"))), stage & ": потеряна связь родителя " & CStr(key)
        Require SameText(CStr(a("Model")), CStr(b("Model"))), stage & ": изменена модель " & CStr(key)
        Require CStr(a("Configuration")) = CStr(b("Configuration")), stage & ": изменена конфигурация " & CStr(key)
        Require CLng(a("Dimensions")) = CLng(b("Dimensions")), stage & ": потеряны размеры " & CStr(key)
    Next key
End Sub

Private Function PropertyViewReadbackMatches(ByVal expectedSheet As Object, _
    ByVal actualSheet As Object, ByVal actualViews As Object) As Boolean
    Dim desired As String, resolved As String, desiredView As Object, resolvedView As Object
    ' Only a local default label may replace a requested explicit view.
    If CBool(expectedSheet("SameCustomProp")) Or CBool(actualSheet("SameCustomProp")) Then Exit Function
    Select Case LCase$(Trim$(CStr(actualSheet("PropertyView"))))
        Case "", "default", "по умолчанию"
        Case Else
            Exit Function
    End Select
    desired = CStr(expectedSheet("PropertyView"))
    resolved = CStr(actualSheet("EffectivePropertyView"))
    If actualViews.Exists(desired) = False Then Exit Function
    If actualViews.Exists(resolved) = False Then Exit Function
    Set desiredView = actualViews(desired)
    Set resolvedView = actualViews(resolved)
    If SameText(CStr(desiredView("Model")), CStr(expectedSheet("PropertyModel"))) = False Then Exit Function
    If CStr(desiredView("Configuration")) <> CStr(expectedSheet("PropertyConfiguration")) Then Exit Function
    If SameText(CStr(resolvedView("Model")), CStr(expectedSheet("PropertyModel"))) = False Then Exit Function
    If CStr(resolvedView("Configuration")) <> CStr(expectedSheet("PropertyConfiguration")) Then Exit Function
    If SameText(CStr(actualSheet("PropertyModel")), CStr(expectedSheet("PropertyModel"))) = False Then Exit Function
    If CStr(actualSheet("PropertyConfiguration")) <> CStr(expectedSheet("PropertyConfiguration")) Then Exit Function
    PropertyViewReadbackMatches = True
End Function

Private Sub VerifyDrawing(ByVal expected As Object, ByVal actual As Object, ByVal stage As String)
    Dim key As Variant, field As Variant, a As Object, b As Object, i As Long
    VerifyStructure expected, actual, stage
    Require expected("Sheets").Count = actual("Sheets").Count, stage & ": число листов не совпадает."
    For i = 1 To expected("Order").Count
        Require SameText(CStr(expected("Order")(i)), CStr(actual("Order")(i))), stage & ": нарушен порядок листов."
    Next i
    For Each key In expected("Sheets").Keys
        Require actual("Sheets").Exists(CStr(key)), stage & ": потерян лист " & CStr(key)
        Set a = expected("Sheets")(CStr(key))
        Set b = actual("Sheets")(CStr(key))
        For Each field In Array("Width", "Height", "Scale1", "Scale2")
            Require NearValue(CDbl(a(field)), CDbl(b(field)), VALUE_TOL), stage & ": параметры листа " & CStr(key)
        Next field
        Require CBool(a("FirstAngle")) = CBool(b("FirstAngle")), stage & ": система проекций листа " & CStr(key)
        Require CBool(a("SameCustomProp")) = CBool(b("SameCustomProp")), _
            stage & ": режим источника свойств листа " & CStr(key)
        LogLine "PROPERTY SOURCE CHECK [" & stage & "]: sheet=" & CStr(key) & _
            "; inherited=" & CStr(a("SameCustomProp")) & "; expectedModel=[" & CStr(a("PropertyModel")) & _
            "]; actualModel=[" & CStr(b("PropertyModel")) & "]; expectedConfig=[" & _
            CStr(a("PropertyConfiguration")) & "]; actualConfig=[" & CStr(b("PropertyConfiguration")) & "]"
        Require SameText(CStr(a("PropertyModel")), CStr(b("PropertyModel"))), _
            stage & ": изменилась модель источника свойств листа " & CStr(key)
        Require CStr(a("PropertyConfiguration")) = CStr(b("PropertyConfiguration")), _
            stage & ": изменилась конфигурация источника свойств листа " & CStr(key)
        If CBool(a("SameCustomProp")) = False And _
            Not SameText(CStr(a("PropertyView")), CStr(b("PropertyView"))) Then
            If PropertyViewReadbackMatches(a, b, actual("Views")) Then
                LogLine "PROPERTY VIEW equivalent default [" & stage & "]: sheet=" & CStr(key) & _
                    "; requested=[" & CStr(a("PropertyView")) & "]; readback=[" & CStr(b("PropertyView")) & _
                    "]; effective=[" & CStr(b("EffectivePropertyView")) & "]"
            Else
                LogLine "PROPERTY VIEW mismatch [" & stage & "]: sheet=" & CStr(key) & _
                    "; expected=[" & CStr(a("PropertyView")) & "]; actual=[" & CStr(b("PropertyView")) & "]"
                Require False, stage & ": вид свойств листа " & CStr(key) & _
                    "; ожидалось=[" & CStr(a("PropertyView")) & "]; фактически=[" & CStr(b("PropertyView")) & "]"
            End If
        End If
        Require CBool(a("FormatVisible")) = CBool(b("FormatVisible")), stage & ": видимость формата " & CStr(key)
        VerifyAnnotations a, b, stage & ": лист " & CStr(key)
    Next key
    For Each key In expected("Views").Keys
        Set a = expected("Views")(CStr(key))
        Set b = actual("Views")(CStr(key))
        Require SameText(CStr(a("Sheet")), CStr(b("Sheet"))), stage & ": вид на неверном листе " & CStr(key)
        Require SameViewFingerprint(a, b), stage & ": геометрия/настройки вида не совпадают: " & CStr(key)
        For Each field In Array("SheetScale", "ParentScale", "Locked", "Alignment", "Suppressed", "Hatches", "Visible")
            If CLng(a(field)) <> CLng(b(field)) Then
                LogLine "VIEW FIELD mismatch [" & stage & "]: view=" & CStr(key) & _
                    "; sheet=" & CStr(b("Sheet")) & "; field=" & CStr(field) & _
                    "; expected=" & CLng(a(field)) & "; actual=" & CLng(b(field))
            End If
            Require CLng(a(field)) = CLng(b(field)), stage & ": " & CStr(field) & " / " & CStr(key) & _
                "; ожидалось=" & CLng(a(field)) & "; фактически=" & CLng(b(field))
        Next field
        If CDbl(a("BoxWidth")) > GEOM_TOL And CDbl(a("BoxHeight")) > GEOM_TOL Then
            Require CDbl(b("BoxWidth")) > GEOM_TOL And CDbl(b("BoxHeight")) > GEOM_TOL, _
                stage & ": пустая геометрия вида " & CStr(key)
        End If
        Require CStr(a("Label")) = CStr(b("Label")), stage & ": изменилось обозначение разреза " & CStr(key)
        VerifyAnnotations a, b, stage & ": вид " & CStr(key)
        If Len(CStr(a("Parent"))) > 0 Then LogLine "LINK OK [" & CStr(key) & "] -> [" & CStr(b("Parent")) & "]"
    Next key
    LogLine "AUDIT OK " & stage & "; sheets=" & actual("Sheets").Count & "; views=" & actual("Views").Count
End Sub

Private Sub VerifyAnnotations(ByVal a As Object, ByVal b As Object, ByVal context As String)
    Require CLng(a("Dimensions")) = CLng(b("Dimensions")), context & ": число размеров не совпадает."
    Require CLng(a("Annotations")) = CLng(b("Annotations")), context & ": число аннотаций не совпадает."
    Require CLng(b("Dangling")) <= CLng(a("Dangling")), context & ": появились висячие аннотации."
    VerifyLinkedNotes a("LinkedNotes"), b("LinkedNotes"), context
End Sub

Private Function FindView(ByVal model As SldWorks.ModelDoc2, ByVal name As String) As SldWorks.View
    Dim drawing As SldWorks.DrawingDoc, groups As Variant, group As Variant
    Dim view As SldWorks.View, i As Long, j As Long
    Set drawing = model
    groups = drawing.GetViews
    If Not IsArray(groups) Then Exit Function
    For i = LBound(groups) To UBound(groups)
        group = groups(i)
        If IsArray(group) Then
            For j = LBound(group) + 1 To UBound(group)
                Set view = group(j)
                If SameText(view.GetName2, name) Then
                    Set FindView = view
                    Exit Function
                End If
            Next j
        End If
    Next i
End Function

Private Function ViewSheetName(ByVal view As SldWorks.View) As String
    Dim sheet As SldWorks.Sheet
    Set sheet = view.Sheet
    Require Not sheet Is Nothing, "Не определен лист вида " & view.GetName2
    ViewSheetName = sheet.GetName
End Function

Private Sub RebuildDrawing(ByVal model As SldWorks.ModelDoc2, Optional ByVal selectedSheets As Variant)
    Dim drawing As SldWorks.DrawingDoc, names As Variant, name As Variant
    Dim sheet As SldWorks.Sheet, originalSheet As String
    Dim forced As Boolean, rebuilt As Boolean
    RequireOwnedPath model.GetPathName
    RequireActive model
    Set drawing = model
    Set sheet = drawing.GetCurrentSheet
    originalSheet = sheet.GetName
    names = drawing.GetSheetNames
    For Each name In names
        If Not IsMissing(selectedSheets) Then
            If Not ContainsSheetValue(selectedSheets, CStr(name)) Then GoTo NextRebuildSheet
        End If
        EnsureSheetActive model, CStr(name), "RebuildDrawing"
        model.ClearSelection2 True
        LogLine "REBUILD begin: doc=" & model.GetPathName & "; sheet=" & CStr(name)
        forced = model.ForceRebuild3(False)
        LogLine "REBUILD ForceRebuild3=" & CStr(forced) & "; raw=" & CLng(forced) & "; sheet=" & CStr(name)
        If forced = False Then
            ' Use the normal rebuild API as an independent fallback, not a retry.
            rebuilt = model.EditRebuild3()
            LogLine "REBUILD EditRebuild3=" & CStr(rebuilt) & "; raw=" & CLng(rebuilt) & "; sheet=" & CStr(name)
            Require rebuilt, "Перестроение не подтверждено: лист=" & CStr(name) & _
                "; ForceRebuild3=False; EditRebuild3=False; документ=" & model.GetPathName
        End If
        Set sheet = drawing.GetCurrentSheet
        Require SameText(sheet.GetName, CStr(name)), "После перестроения изменился активный лист."
NextRebuildSheet:
    Next name
    EnsureSheetActive model, originalSheet, "RebuildDrawing restore"
End Sub

Private Sub CheckSavedInput(ByVal path As String)
    Dim model As SldWorks.ModelDoc2
    Require fsoRun.FileExists(path), "Файл не найден: " & path
    Require LCase$(fsoRun.GetExtensionName(path)) = "slddrw", "Нужен файл .slddrw: " & path
    Set model = swApp.GetOpenDocumentByName(path)
    If Not model Is Nothing Then
        Require model.GetSaveFlag = False, "Сначала сохраните открытый чертеж с изменениями: " & path
    End If
End Sub

Private Function CloneInput(ByVal path As String, ByVal copyPath As String, _
    Optional ByRef validatedSnapshot As Object = Nothing, _
    Optional ByVal rebuildBeforeSnapshot As Boolean = False) As SldWorks.ModelDoc2
    Dim original As SldWorks.ModelDoc2, copy As SldWorks.ModelDoc2
    Dim originalDrawing As SldWorks.DrawingDoc, originalSheet As String
    Dim wasOpen As Boolean, baseline As Object, actual As Object
    Dim errors As Long, warnings As Long, e As Long, message As String
    Dim oldFolder As String, opened As Boolean
    On Error GoTo Failed
    CheckSavedInput path
    RequireOwnedPath copyPath
    Require Not fsoRun.FileExists(copyPath), "Рабочая копия уже существует: " & copyPath
    Set original = swApp.GetOpenDocumentByName(path)
    wasOpen = Not original Is Nothing
    If Not wasOpen Then
        Set original = swApp.OpenDoc6(path, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errors, warnings)
        opened = Not original Is Nothing
    End If
    Require Not original Is Nothing, "Не открыт исходный чертеж: " & path & "; код=" & errors
    Set originalDrawing = original
    originalSheet = originalDrawing.GetCurrentSheet.GetName
    LogLine "CLONE baseline: " & path
    Set baseline = CaptureDrawing(original)
    ' A byte copy preserves the full native graph and avoids changing any source
    ' references through SaveAs. Keep referenced models loaded from the original.
    LogLine "CLONE file copy begin: " & copyPath
    fsoRun.CopyFile path, copyPath, False
    LogLine "CLONE file copy complete: bytes=" & fsoRun.GetFile(copyPath).Size
    SetAttr copyPath, vbNormal
    oldFolder = swApp.GetCurrentWorkingDirectory
    EnsureWorkingDirectory fsoRun.GetParentFolderName(path), "CloneInput source"
    Set copy = OpenWorkingDrawing(copyPath)
    EnsureWorkingDirectory oldFolder, "CloneInput restore"
    oldFolder = ""
    LogLine "CLONE copy: " & copyPath
    If rebuildBeforeSnapshot Then RebuildDrawing copy
    Set actual = CaptureDrawing(copy)
    VerifyDrawing baseline, actual, "CLONE " & path
    Set validatedSnapshot = actual
    If opened Then
        swApp.CloseDoc original.GetTitle
    Else
        EnsureSheetActive original, originalSheet, "CloneInput restore"
    End If
    Set CloneInput = copy
    Exit Function
Failed:
    e = Err.Number
    message = Err.Description
    On Error Resume Next
    If Len(oldFolder) > 0 Then EnsureWorkingDirectory oldFolder, "CloneInput failure restore"
    If opened Then
        If Not original Is Nothing Then swApp.CloseDoc original.GetTitle
    ElseIf Len(originalSheet) > 0 Then
        EnsureSheetActive original, originalSheet, "CloneInput restore"
    End If
    CloseOwnedDrawing copy
    On Error GoTo 0
    Err.Raise e, "CloneInput", message
End Function

Private Function OpenWorkingDrawing(ByVal path As String) As SldWorks.ModelDoc2
    Dim model As SldWorks.ModelDoc2, errors As Long, warnings As Long
    RequireOwnedPath path
    LogLine "OPEN begin: " & path
    Set model = swApp.OpenDoc6(path, swDocDRAWING, swOpenDocOptions_Silent, "", errors, warnings)
    LogLine "OPEN " & path & "; errors=" & errors & "; warnings=" & warnings
    Require Not model Is Nothing, "Не открыта рабочая копия. Код: " & errors
    Require model.GetType = swDocDRAWING, "Рабочая копия не является чертежом."
    Require SameText(NormalizePath(model.GetPathName), NormalizePath(path)), "SOLIDWORKS вернул другой документ."
    Set OpenWorkingDrawing = model
End Function

Private Sub EnsureWorkingDirectory(ByVal folder As String, ByVal context As String)
    Dim requested As String, beforeFolder As String, afterFolder As String
    Dim apiResult As Boolean
    Require Len(folder) > 0, "Пустая рабочая папка [" & context & "]."
    Require fsoRun.FolderExists(folder), "Рабочая папка не существует [" & context & "]: " & folder
    requested = fsoRun.GetFolder(folder).Path
    beforeFolder = swApp.GetCurrentWorkingDirectory
    If SameFolder(beforeFolder, requested) Then
        LogLine "CWD CURRENT [" & context & "] folder=" & beforeFolder
        Exit Sub
    End If
    apiResult = swApp.SetCurrentWorkingDirectory(requested)
    afterFolder = swApp.GetCurrentWorkingDirectory
    LogLine "CWD SET [" & context & "] requested=" & requested & _
        "; before=" & beforeFolder & "; api=" & CStr(apiResult) & "; current=" & afterFolder
    Require SameFolder(afterFolder, requested), "Не задана рабочая папка [" & context & _
        "]; requested=" & requested & "; before=" & beforeFolder & _
        "; current=" & afterFolder & "; API=" & CStr(apiResult)
End Sub

Private Function SameFolder(ByVal first As String, ByVal second As String) As Boolean
    ' GetFolder resolves equivalent existing folder spellings without changing CWD.
    If Len(first) = 0 Or Len(second) = 0 Then Exit Function
    If Not fsoRun.FolderExists(first) Then Exit Function
    If Not fsoRun.FolderExists(second) Then Exit Function
    SameFolder = SameText(fsoRun.GetFolder(first).Path, fsoRun.GetFolder(second).Path)
End Function

Private Sub EnsureSheetActive(ByVal model As SldWorks.ModelDoc2, ByVal sheetName As String, ByVal context As String)
    Dim drawing As SldWorks.DrawingDoc, sheet As SldWorks.Sheet
    Dim beforeName As String, afterName As String, apiResult As Boolean
    RequireActive model
    Set drawing = model
    Set sheet = drawing.GetCurrentSheet
    If Not sheet Is Nothing Then beforeName = sheet.GetName
    ' ActivateSheet is unnecessary if the requested sheet is already current.
    If SameText(beforeName, sheetName) Then
        drawing.EditSheet
        LogLine "SHEET CURRENT [" & context & "] doc=" & model.GetPathName & "; sheet=" & beforeName
        Exit Sub
    End If
    drawing.EditSheet
    apiResult = drawing.ActivateSheet(sheetName)
    Set sheet = drawing.GetCurrentSheet
    If Not sheet Is Nothing Then afterName = sheet.GetName
    LogLine "SHEET ACTIVATE [" & context & "] doc=" & model.GetPathName & _
        "; requested=" & sheetName & "; before=" & beforeName & _
        "; api=" & CStr(apiResult) & "; current=" & afterName
    ' Verify the actual postcondition even when the API reports success.
    Require SameText(afterName, sheetName), "Не активирован лист " & sheetName & _
        " [" & context & "]; текущий=" & afterName & "; API=" & CStr(apiResult) & _
        "; документ=" & model.GetPathName
    drawing.EditSheet
End Sub

Private Sub RequireActive(ByVal model As SldWorks.ModelDoc2)
    Dim active As SldWorks.ModelDoc2, errors As Long
    Require Not model Is Nothing, "Документ не задан."
    Set active = swApp.ActiveDoc
    If Not active Is Nothing Then
        If SameText(NormalizePath(active.GetPathName), NormalizePath(model.GetPathName)) Then Exit Sub
    End If
    Set active = swApp.ActivateDoc3(model.GetTitle, False, swDontRebuildActiveDoc, errors)
    Require Not active Is Nothing, "Не удалось активировать " & model.GetTitle & "; код=" & errors
    Require SameText(NormalizePath(active.GetPathName), NormalizePath(model.GetPathName)), _
        "Активирован другой документ с таким же именем."
End Sub

Private Sub CloseOwnedDrawing(ByRef model As SldWorks.ModelDoc2)
    If model Is Nothing Then Exit Sub
    RequireOwnedPath model.GetPathName
    swApp.CloseDoc model.GetTitle
    Set model = Nothing
End Sub

Private Sub RequireOwnedPath(ByVal path As String)
    Dim root As String, full As String
    Require Len(runFolder) > 0, "Рабочая папка не создана."
    root = NormalizePath(fsoRun.GetAbsolutePathName(runFolder)) & "\"
    full = NormalizePath(fsoRun.GetAbsolutePathName(path))
    If Len(publishedDrawingPath) > 0 Then
        If SameText(full, NormalizePath(fsoRun.GetAbsolutePathName(publishedDrawingPath))) Then Exit Sub
    End If
    Require Left$(full, Len(root)) = root, "Запрещено изменять файл вне папки запуска: " & path
End Sub

Private Sub InitRunFolder(ByVal mainPath As String)
    Dim base As String, parent As String, token As String, suffix As Long
    base = fsoRun.GetParentFolderName(mainPath)
    Require Len(base) > 0, "Не определена папка основного чертежа."
    Require fsoRun.FolderExists(base), "Не найдена папка основного чертежа: " & base
    parent = fsoRun.BuildPath(base, "runs")
    If fsoRun.FolderExists(parent) = False Then fsoRun.CreateFolder parent
    token = Format$(Now, "yyyymmdd_hhnnss")
    runToken = token
    runFolder = fsoRun.BuildPath(parent, runToken)
    Do While fsoRun.FolderExists(runFolder) Or _
        fsoRun.FileExists(fsoRun.BuildPath(base, "MERGED_" & runToken & ".slddrw")) Or _
        fsoRun.FileExists(fsoRun.BuildPath(base, "MERGED_" & runToken & ".pdf"))
        suffix = suffix + 1
        runToken = token & "_" & CStr(suffix)
        runFolder = fsoRun.BuildPath(parent, runToken)
    Loop
    fsoRun.CreateFolder runFolder
    Set logStream = fsoRun.CreateTextFile(runFolder & "\merge.log", False, True)
End Sub

Private Sub WriteStatus(ByVal status As String, ByVal message As String)
    Dim stream As Object
    If Len(runFolder) = 0 Then Exit Sub
    Set stream = fsoRun.CreateTextFile(runFolder & "\STATUS.txt", True, True)
    stream.WriteLine BUILD_ID
    stream.WriteLine status
    stream.WriteLine message
    stream.Close
End Sub

Private Sub WriteManifest(ByVal snapshot As Object, ByVal path As String)
    Dim stream As Object, key As Variant, rec As Object
    RequireOwnedPath path
    Set stream = fsoRun.CreateTextFile(path, False, True)
    stream.WriteLine "sheet" & vbTab & "view" & vbTab & "type" & vbTab & "parent" & vbTab & _
        "model" & vbTab & "configuration" & vbTab & "x_m" & vbTab & "y_m" & vbTab & "scale"
    For Each key In snapshot("Views").Keys
        Set rec = snapshot("Views")(CStr(key))
        stream.WriteLine CStr(rec("Sheet")) & vbTab & CStr(key) & vbTab & CStr(rec("Type")) & vbTab & _
            CStr(rec("Parent")) & vbTab & CStr(rec("Model")) & vbTab & CStr(rec("Configuration")) & vbTab & _
            CStr(rec("X")) & vbTab & CStr(rec("Y")) & vbTab & CStr(rec("Scale"))
    Next key
    stream.Close
End Sub

Private Sub LogLine(ByVal message As String)
    Debug.Print message
    If Not logStream Is Nothing Then logStream.WriteLine Format$(Now, "hh:nn:ss") & " " & message
End Sub

Private Function FlagValue(ByVal value As Boolean) As Boolean
    ' Return a canonical VBA Boolean (-1/0), even if COM supplied +1 for True.
    If value Then
        FlagValue = True
    Else
        FlagValue = False
    End If
End Function

Private Sub Require(ByVal condition As Boolean, ByVal message As String)
    ' Test truth directly: Not is a bitwise operation on raw COM Boolean values.
    If condition Then Exit Sub
    Err.Raise vbObjectError + 1600, BUILD_ID, message
End Sub

Private Function NewDictionary() As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    Set NewDictionary = result
End Function

Private Function CloneRecord(ByVal source As Object) As Object
    Dim result As Object, key As Variant
    Set result = NewDictionary()
    For Each key In source.Keys
        result.Add CStr(key), source(key)
    Next key
    Set CloneRecord = result
End Function

Private Function SameText(ByVal a As String, ByVal b As String) As Boolean
    SameText = (StrComp(a, b, vbTextCompare) = 0)
End Function

Private Function NormalizePath(ByVal value As String) As String
    NormalizePath = LCase$(Replace(value, "/", "\"))
End Function

Private Function NearValue(ByVal a As Double, ByVal b As Double, ByVal tolerance As Double) As Boolean
    NearValue = (Abs(a - b) <= tolerance)
End Function

' Only sources whose parent chains cross sheets need an editable private copy.
Private Function PrepareSource(ByVal path As String, ByVal fileIndex As Long, _
    ByRef snapshot As Object, ByRef direct As Boolean, ByRef opened As Boolean, _
    ByRef initialSheet As String) As SldWorks.ModelDoc2
    Dim model As SldWorks.ModelDoc2, drawing As SldWorks.DrawingDoc
    Dim key As Variant, order As Collection, needsCopy As Boolean
    Dim errors As Long, warnings As Long, number As Long, message As String
    On Error GoTo Failed
    direct = True
    opened = False
    CheckSavedInput path
    Set model = swApp.GetOpenDocumentByName(path)
    If model Is Nothing Then
        Set model = swApp.OpenDoc6(path, swDocDRAWING, swOpenDocOptions_Silent Or swOpenDocOptions_ReadOnly, "", errors, warnings)
        opened = Not model Is Nothing
    End If
    Require Not model Is Nothing, "Не открыт источник: " & path & "; код=" & errors
    Set drawing = model
    initialSheet = drawing.GetCurrentSheet.GetName
    Set snapshot = CaptureDrawing(model)
    Set order = DependencyOrder(snapshot("Views"))
    For Each key In order
        If Not SameText(CStr(snapshot("Views")(CStr(key))("Sheet")), _
            CStr(snapshot("Views")(CStr(key))("Carrier"))) Then needsCopy = True
    Next key
    If needsCopy Then
        ReleaseSource model, True, opened, initialSheet
        direct = False
        opened = False
        Set model = CloneInput(path, runFolder & "\SOURCE_" & runToken & "_" & Format$(fileIndex, "0000") & ".slddrw")
        ' CloneInput verified the byte copy against the original snapshot.
        LogLine "SOURCE MODE private copy: cross-sheet dependent views; " & path
    Else
        LogLine "SOURCE MODE direct read-only transfer: " & path
    End If
    Set PrepareSource = model
    Exit Function
Failed:
    number = Err.Number
    message = Err.Description
    On Error Resume Next
    ReleaseSource model, direct, opened, initialSheet
    On Error GoTo 0
    Err.Raise number, "PrepareSource", message
End Function

Private Sub ReleaseSource(ByRef model As SldWorks.ModelDoc2, ByVal direct As Boolean, _
    ByVal opened As Boolean, ByVal initialSheet As String)
    If model Is Nothing Then Exit Sub
    If direct Then
        If opened Then
            swApp.CloseDoc model.GetTitle
        Else
            If Len(initialSheet) > 0 Then EnsureSheetActive model, initialSheet, "ReleaseSource restore"
            model.ClearSelection2 True
        End If
        Set model = Nothing
    Else
        CloseOwnedDrawing model
    End If
End Sub

Private Function ContainsSheetValue(ByVal mapping As Object, ByVal name As String) As Boolean
    Dim key As Variant
    For Each key In mapping.Keys
        If SameText(CStr(mapping(key)), name) Then
            ContainsSheetValue = True
            Exit Function
        End If
    Next key
End Function

Private Function LocalizeTargetPropertyView(ByVal snapshot As Object, ByVal sourceSheet As String, _
    ByVal targetSheet As String, ByVal requested As String, ByVal helpers As Object) As String
    Dim view As SldWorks.View, sourceView As Object, sheetRec As Object
    Dim name As String, modelPath As String, configuration As String
    RequireOwnedPath targetModel.GetPathName
    Set sheetRec = snapshot("Sheets")(sourceSheet)
    If Len(CStr(sheetRec("PropertyModel"))) = 0 Then
        LocalizeTargetPropertyView = requested
        Exit Function
    End If
    Set view = FindView(targetModel, requested)
    Require Not view Is Nothing, "Не найден источник свойств: " & requested
    If SameText(ViewSheetName(view), targetSheet) Then
        LocalizeTargetPropertyView = requested
        Exit Function
    End If
    Set sheetRec = snapshot("Sheets")(sourceSheet)
    Set sourceView = snapshot("Views")(CStr(sheetRec("EffectivePropertyView")))
    modelPath = CStr(sheetRec("PropertyModel"))
    configuration = CStr(sheetRec("PropertyConfiguration"))
    Require Len(modelPath) > 0 And Len(CStr(sourceView("Orientation"))) > 0, _
        "Недостаточно данных для локального источника свойств: " & targetSheet
    EnsureSheetActive targetModel, targetSheet, "Local property carrier"
    targetModel.ClearSelection2 True
    ' Create geometry from the required model, not from the detail already on this sheet.
    Set view = targetDraw.CreateDrawViewFromModelView3(modelPath, CStr(sourceView("Orientation")), _
        CDbl(sheetRec("Width")) / 2#, CDbl(sheetRec("Height")) / 2#, 0#)
    Require Not view Is Nothing, "Не создан служебный вид свойств: " & targetSheet
    name = view.GetName2
    view.ReferencedConfiguration = configuration
    view.UseSheetScale = 0
    view.UseParentScale = False
    view.ScaleDecimal = 0.01
    view.SetVisible False, False
    targetDraw.HiddenViewsVisible = False
    Require FlagValue(targetDraw.HiddenViewsVisible) = False, "Не отключён показ скрытых служебных видов."
    Require FlagValue(view.GetVisible) = False, "Не скрыт служебный вид: " & name
    Require SameText(ViewSheetName(view), targetSheet), "Служебный вид создан на другом листе."
    Require SameText(NormalizePath(view.GetReferencedModelName), modelPath), "Неверная модель служебного вида."
    Require view.ReferencedConfiguration = configuration, "Неверная конфигурация служебного вида."
    helpers.Add name, sheetRec
    LogLine "PROPERTY CARRIER created: sheet=" & targetSheet & "; view=" & name & "; hidden=True; model=" & modelPath
    LocalizeTargetPropertyView = name
End Function

Private Sub SaveFailedCheckpoint(ByVal model As SldWorks.ModelDoc2, ByVal failure As String)
    Dim saved As Boolean, errors As Long, warnings As Long
    If model Is Nothing Then Exit Sub
    ' Preserve the newly allocated result or private source; never label it SUCCESS.
    ' RequireOwnedPath admits only the exact registered output outside runs.
    RequireOwnedPath model.GetPathName
    model.ClearSelection2 True
    LogLine "FAILED CHECKPOINT begin: " & model.GetPathName
    saved = model.Save3(swSaveAsOptions_Silent, errors, warnings)
    LogLine "FAILED CHECKPOINT saved=" & CStr(saved) & "; errors=" & errors & "; warnings=" & warnings
    WriteStatus "FAILED", failure & vbCrLf & "Непроверенный рабочий файл: " & model.GetPathName & _
        "; сохранён=" & CStr(saved) & "; код=" & errors
End Sub

' Keep live component links. Never substitute evaluated text for a property expression.
Private Function CaptureLinkedNotes(ByVal view As SldWorks.View) As Collection
    Dim result As Collection, notes As Variant, i As Long
    Dim note As SldWorks.Note, annotation As SldWorks.Annotation, rec As Object
    Dim raw As String, p As Variant
    Set result = New Collection
    notes = view.GetNotes
    If IsArray(notes) Then
        For i = LBound(notes) To UBound(notes)
            Set note = notes(i)
            raw = note.PropertyLinkedText
            If IsComponentPropertyNote(raw) Then
                Set annotation = note.GetAnnotation
                Require Not annotation Is Nothing, "Не прочитана аннотация связанной заметки."
                Set rec = NewDictionary()
                rec.Add "Name", note.GetName
                rec.Add "Link", raw
                rec.Add "Text", note.GetText
                rec.Add "Layer", annotation.Layer
                rec.Add "Visible", CLng(annotation.Visible)
                rec.Add "Attachments", NoteAttachmentSignature(annotation, view)
                p = annotation.GetPosition
                Require IsArray(p), "Нет координат связанной заметки: " & note.GetName
                rec.Add "X", CDbl(p(0))
                rec.Add "Y", CDbl(p(1))
                result.Add rec
            End If
        Next i
    Else
        Require view.GetNoteCount = 0, "Не прочитан массив заметок: " & view.GetName2
    End If
    Set CaptureLinkedNotes = result
End Function

Private Function IsComponentPropertyNote(ByVal raw As String) As Boolean
    IsComponentPropertyNote = (InStr(1, raw, "$PRPSMODEL:", vbTextCompare) > 0 And _
        InStr(1, raw, "$COMP:" & Chr$(34), vbTextCompare) > 0)
End Function

Private Function NoteAttachmentSignature(ByVal annotation As SldWorks.Annotation, ByVal view As SldWorks.View) As String
    Dim attached As Variant, i As Long, entity As SldWorks.Entity
    Dim obj As Object, component As SldWorks.Component2, item As String
    Dim drawingComponent As SldWorks.DrawingComponent, parts() As String, j As Long, tmp As String
    Dim count As Long
    count = annotation.GetAttachedEntityCount3
    NoteAttachmentSignature = CStr(count)
    If count = 0 Then Exit Function
    attached = annotation.GetAttachedEntities3
    Require IsArray(attached), "Не прочитаны привязки связанной заметки."
    Require UBound(attached) - LBound(attached) + 1 = count, "Неполный список привязок заметки."
    ReDim parts(LBound(attached) To UBound(attached))
    For i = LBound(attached) To UBound(attached)
        Set obj = attached(i)
        Require Not obj Is Nothing, "Пустая привязка связанной заметки."
        Set component = Nothing
        If TypeOf obj Is SldWorks.Component2 Then
            Set component = obj
        ElseIf TypeOf obj Is SldWorks.DrawingComponent Then
            Set drawingComponent = obj
            Set component = drawingComponent.Component
        Else
            ' Unsupported attachment objects must stop repair, never select another part.
            Set entity = obj
            ' A section entity needs the component instance in this drawing view.
            Set drawingComponent = entity.GetDrawingComponent(view)
            If Not drawingComponent Is Nothing Then Set component = drawingComponent.Component
            If component Is Nothing Then Set component = entity.GetComponent
        End If
        Require Not component Is Nothing, "Не определена деталь привязки связанной заметки."
        item = NormalizePath(component.GetPathName) & "|" & component.ReferencedConfiguration & "|" & component.Name2
        parts(i) = item
    Next i
    ' Attachment ordering is not identity. Preserve multiplicity, compare sorted values.
    For i = LBound(parts) To UBound(parts)
        For j = i + 1 To UBound(parts)
            If StrComp(parts(i), parts(j), vbBinaryCompare) > 0 Then
                tmp = parts(i): parts(i) = parts(j): parts(j) = tmp
            End If
        Next j
    Next i
    NoteAttachmentSignature = NoteAttachmentSignature & ":" & Join(parts, vbTab)
End Function

Private Function RemapComponentViews(ByVal raw As String, ByVal mapping As Object) As String
    Dim cursor As Long, marker As Long, closing As Long, start As Long, slash As Long
    Dim expression As String, at As Long, name As String, output As String
    cursor = 1
    Do
        marker = InStr(cursor, raw, "$COMP:" & Chr$(34), vbTextCompare)
        If marker = 0 Then Exit Do
        start = marker + Len("$COMP:" & Chr$(34))
        closing = InStr(start, raw, Chr$(34))
        Require closing > 0, "Не закрыта строка $COMP в заметке."
        output = output & Mid$(raw, cursor, start - cursor)
        expression = Mid$(raw, start, closing - start)
        ' Only @view/ qualifiers inside this component expression may change.
        ' Work from the original expression so mappings never cascade (1->2->3).
        Do
            at = InStr(1, expression, "@", vbBinaryCompare)
            If at = 0 Then Exit Do
            slash = InStr(at + 1, expression, "/", vbBinaryCompare)
            If slash = 0 Then Exit Do
            name = Mid$(expression, at + 1, slash - at - 1)
            output = output & Left$(expression, at)
            If mapping.Exists(name) Then
                output = output & CStr(mapping(name)) & "/"
            Else
                output = output & name & "/"
            End If
            expression = Mid$(expression, slash + 1)
        Loop
        output = output & expression & Chr$(34)
        cursor = closing + 1
    Loop
    RemapComponentViews = output & Mid$(raw, cursor)
End Function

Private Function MappedLinkedNotes(ByVal original As Collection, ByVal mapping As Object) As Collection
    Dim result As Collection, item As Variant, rec As Object
    Set result = New Collection
    For Each item In original
        Set rec = CloneRecord(item)
        If rec.Exists("ResolvedLink") Then
            rec("Link") = CStr(rec("ResolvedLink"))
        Else
            rec("Link") = RemapComponentViews(CStr(rec("Link")), mapping)
        End If
        result.Add rec
    Next item
    Set MappedLinkedNotes = result
End Function

Private Function SameNotePosition(ByVal a As Object, ByVal b As Object) As Boolean
    SameNotePosition = NearValue(CDbl(a("X")), CDbl(b("X")), GEOM_TOL) And _
        NearValue(CDbl(a("Y")), CDbl(b("Y")), GEOM_TOL)
End Function

Private Sub RestoreImportedLinkedNotes(ByVal original As Object, ByVal sheetMap As Object, ByVal viewMap As Object)
    Dim key As Variant, view As SldWorks.View, drawing As SldWorks.DrawingDoc
    Set drawing = targetModel
    For Each key In original("Order")
        EnsureSheetActive targetModel, CStr(sheetMap(CStr(key))), "Restore linked sheet notes"
        Set view = drawing.GetFirstView
        RestoreViewLinkedNotes view, original("Sheets")(CStr(key))("LinkedNotes"), viewMap
    Next key
    For Each key In original("Views").Keys
        Set view = FindView(targetModel, CStr(viewMap(CStr(key))))
        Require Not view Is Nothing, "Не найден вид для связанных заметок."
        EnsureSheetActive targetModel, ViewSheetName(view), "Restore linked view notes"
        RestoreViewLinkedNotes view, original("Views")(CStr(key))("LinkedNotes"), viewMap
    Next key
End Sub

Private Sub RestoreViewLinkedNotes(ByVal view As SldWorks.View, ByVal expected As Collection, ByVal mapping As Object)
    Dim notes As Variant, item As Variant, note As SldWorks.Note, candidate As SldWorks.Note
    Dim annotation As SldWorks.Annotation
    Dim desired As String, raw As String, i As Long, matches As Long, p As Variant
    Dim beforeAttachments As String, afterAttachments As String, originalLink As String
    If expected.Count = 0 Then Exit Sub
    RequireOwnedPath targetModel.GetPathName
    notes = view.GetNotes
    Require IsArray(notes), "Потеряны связанные заметки вида " & view.GetName2
    For Each item In expected
        desired = RemapComponentViews(CStr(item("Link")), mapping)
        matches = 0
        Set candidate = Nothing
        For i = LBound(notes) To UBound(notes)
            Set note = notes(i)
            raw = note.PropertyLinkedText
            If raw = CStr(item("Link")) Or raw = desired Then
                Set annotation = note.GetAnnotation
                p = annotation.GetPosition
                If NearValue(CDbl(p(0)), CDbl(item("X")), GEOM_TOL) And _
                    NearValue(CDbl(p(1)), CDbl(item("Y")), GEOM_TOL) Then
                    matches = matches + 1
                    Set candidate = note
                End If
            End If
        Next i
        Require matches = 1, "Неоднозначная/потерянная связанная заметка: " & CStr(item("Name")) & "; вид=" & view.GetName2
        Set annotation = candidate.GetAnnotation
        originalLink = candidate.PropertyLinkedText
        beforeAttachments = NoteAttachmentSignature(annotation, view)
        LogLine "LINKED NOTE binding before: view=" & view.GetName2 & "; note=" & CStr(item("Name")) & _
            "; expected=[" & CStr(item("Attachments")) & "]; actual=[" & beforeAttachments & "]"
        LogLine "LINKED NOTE expression: original=[" & CStr(item("Link")) & _
            "]; current=[" & originalLink & "]; requested=[" & desired & "]"
        ' Resolve the live reference in the destination view before checking attachments.
        ' No leaders or attachment entities are reassigned here.
        candidate.PropertyLinkedText = desired
        Require candidate.PropertyLinkedText = desired, "Не восстановлена ссылка заметки: " & CStr(item("Name"))
        Set annotation = candidate.GetAnnotation
        afterAttachments = NoteAttachmentSignature(annotation, view)
        If afterAttachments = CStr(item("Attachments")) Then
            If InStr(1, originalLink, "SectionAssembly", vbTextCompare) > 0 And _
                Len(Trim$(CStr(item("Text")))) > 0 And Len(Trim$(candidate.GetText)) = 0 Then
                desired = SectionComponentLink(view, annotation, desired, afterAttachments)
                candidate.PropertyLinkedText = desired
                Require candidate.PropertyLinkedText = desired, "Не записана ссылка компонента разреза."
                item("ResolvedLink") = desired
                Set annotation = candidate.GetAnnotation
                afterAttachments = NoteAttachmentSignature(annotation, view)
                LogLine "SECTION NOTE rebuilt: view=" & view.GetName2 & "; expression=[" & desired & _
                    "]; text=[" & candidate.GetText & "]"
            End If
        End If
        LogLine "LINKED NOTE binding after: view=" & view.GetName2 & "; note=" & CStr(item("Name")) & _
            "; actual=[" & afterAttachments & "]"
        If afterAttachments <> CStr(item("Attachments")) Then
            candidate.PropertyLinkedText = originalLink
            Require False, "Изменилась деталь привязки заметки: " & CStr(item("Name")) & _
                "; вид=" & view.GetName2 & "; подробности LINKED NOTE binding в merge.log."
        End If
        ' Reset the original layer as paste can inherit the destination's current layer.
        If annotation.Layer <> CStr(item("Layer")) Then annotation.Layer = CStr(item("Layer"))
        If CLng(annotation.Visible) <> CLng(item("Visible")) Then annotation.Visible = CLng(item("Visible"))
        LogLine "LINKED NOTE restored: view=" & view.GetName2 & "; note=" & candidate.GetName & _
            "; sourceText=[" & CStr(item("Text")) & "]; currentText=[" & candidate.GetText & "]"
    Next item
End Sub

Private Sub VerifyLinkedNotes(ByVal expected As Collection, ByVal actual As Collection, ByVal context As String)
    Dim item As Variant, candidate As Variant, count As Long, used As Object, index As Long, found As Long
    Require expected.Count = actual.Count, context & ": изменилось число заметок со свойствами деталей."
    Set used = NewDictionary()
    For Each item In expected
        count = 0
        index = 0
        For Each candidate In actual
            index = index + 1
            If Not used.Exists(CStr(index)) Then
                If CStr(item("Link")) = CStr(candidate("Link")) And SameNotePosition(item, candidate) Then
                    If CStr(item("Attachments")) = CStr(candidate("Attachments")) Then
                        count = count + 1
                        found = index
                    End If
                End If
            End If
        Next candidate
        Require count = 1, context & ": потеряна/неоднозначна ссылка заметки " & CStr(item("Name"))
        used.Add CStr(found), True
        Set candidate = actual(found)
        Require CStr(item("Layer")) = CStr(candidate("Layer")), context & ": изменился слой заметки " & CStr(item("Name"))
        Require CLng(item("Visible")) = CLng(candidate("Visible")), context & ": изменилась видимость заметки " & CStr(item("Name"))
        If Len(Trim$(CStr(item("Text")))) > 0 Then
            LogLine "LINKED NOTE check [" & context & "]: note=" & CStr(item("Name")) & _
                "; expected=[" & CStr(item("Text")) & "]; actual=[" & CStr(candidate("Text")) & "]"
            Require CStr(item("Text")) = CStr(candidate("Text")), context & _
                ": пропал/изменился текст связанной заметки " & CStr(item("Name"))
        End If
    Next item
End Sub


' Use the native drawing instance chain, including the generated SectionAssembly.
Private Function SectionComponentLink(ByVal view As SldWorks.View, ByVal annotation As SldWorks.Annotation, _
    ByVal raw As String, ByVal signature As String) As String
    Dim attached As Variant, obj As Object, entity As SldWorks.Entity
    Dim node As SldWorks.DrawingComponent, component As SldWorks.Component2
    Dim names As Variant, expression As String, unique As String, i As Long
    Dim marker As Long, start As Long, closing As Long, identity As String
    Require Left$(signature, 2) = "1:", "Служебная ссылка разреза имеет несколько привязок."
    marker = InStr(1, raw, "$COMP:" & Chr$(34), vbTextCompare)
    Require marker > 0, "Нет ссылки компонента разреза."
    start = marker + Len("$COMP:" & Chr$(34))
    closing = InStr(start, raw, Chr$(34))
    Require closing > 0, "Не закрыта ссылка компонента разреза."
    Require InStr(closing + 1, raw, "$COMP:", vbTextCompare) = 0, "Несколько ссылок компонентов в заметке разреза."
    attached = annotation.GetAttachedEntities3
    Require IsArray(attached), "Не прочитана привязка заметки разреза."
    Require UBound(attached) = LBound(attached), "Ожидалась одна привязка заметки разреза."
    Set obj = attached(LBound(attached))
    If TypeOf obj Is SldWorks.DrawingComponent Then
        Set node = obj
    ElseIf TypeOf obj Is SldWorks.Component2 Then
        Set component = obj
        Set node = component.GetDrawingComponent(view)
    Else
        Set entity = obj
        Set node = entity.GetDrawingComponent(view)
    End If
    Require Not node Is Nothing, "Нет контекста DrawingComponent привязки разреза."
    Set component = node.Component
    Require Not component Is Nothing, "Нет компонента привязки разреза."
    identity = NormalizePath(component.GetPathName) & "|" & component.ReferencedConfiguration & "|" & component.Name2
    Require identity = Mid$(signature, 3), "Контекст разреза изменил компонент заметки."
    unique = view.GetUniqueName
    Require Len(unique) > 0, "Нет внутреннего имени вида разреза."
    names = Split(node.Name, "/")
    Require UBound(names) >= 1, "Не прочитана полная цепочка экземпляров разреза."
    Require InStr(1, CStr(names(0)), "SectionAssembly", vbTextCompare) > 0, "Нет служебной сборки в цепочке разреза."
    expression = CStr(names(0)) & "@" & unique
    For i = 1 To UBound(names)
        expression = expression & "/" & CStr(names(i)) & "@" & InstanceModelName(CStr(names(i - 1)))
    Next i
    Require InStr(expression, Chr$(34)) = 0, "Неподдерживаемая кавычка в адресе компонента."
    LogLine "SECTION NOTE native address: unique=" & unique & "; chain=[" & node.Name & "]; componentID=[" & component.GetSelectByIDString & "]"
    SectionComponentLink = Left$(raw, start - 1) & expression & Mid$(raw, closing)
End Function

Private Function InstanceModelName(ByVal instance As String) As String
    Dim dash As Long, suffix As String, i As Long
    dash = InStrRev(instance, "-")
    Require dash > 1 And dash < Len(instance), "Нет номера экземпляра: " & instance
    suffix = Mid$(instance, dash + 1)
    For i = 1 To Len(suffix)
        Require Mid$(suffix, i, 1) >= "0" And Mid$(suffix, i, 1) <= "9", "Неподдерживаемый номер экземпляра: " & instance
    Next i
    InstanceModelName = Left$(instance, dash - 1)
End Function

Private Function GetOrderedSourceFiles(ByVal selectedFiles As Collection, _
    ByVal mainPath As String) As Collection
    Dim dialog As frmDrawingOrder, item As Variant, result As Collection
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed
    Set dialog = New frmDrawingOrder
    dialog.LoadCandidates selectedFiles, mainPath
    dialog.Show vbModal
    If dialog.Accepted Then
        Set result = dialog.SelectedFiles
        Set GetOrderedSourceFiles = result
        LogLine "Порядок исходных чертежей (форма VBA):"
        For Each item In result
            LogLine CStr(item)
        Next item
    End If
    Unload dialog
    Exit Function
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    If Not dialog Is Nothing Then Unload dialog
    On Error GoTo 0
    Err.Raise errorNumber, "GetOrderedSourceFiles", errorText
End Function

Public Function OrderPickFiles(ByVal mainPath As String) As Collection
    Set OrderPickFiles = PickMultipleDrawings("Добавить исходные чертежи", GetFolder(mainPath))
End Function

Public Function OrderFolderFiles(ByVal folder As String, ByVal mainPath As String) As Collection
    Set OrderFolderFiles = GetAllDrawingsInFolder(folder, mainPath)
End Function

Private Function PickOneDrawing( _
    ByVal dialogTitle As String, _
    ByVal initialFolder As String) As String

    Dim ofn As OPENFILENAMEW
    Dim filter As String
    Dim buffer As String
    Dim titleBuffer As String
    Dim p As Long

    filter = "Чертежи SOLIDWORKS (*.slddrw)" & vbNullChar & _
             "*.slddrw" & vbNullChar & _
             "Все файлы (*.*)" & vbNullChar & _
             "*.*" & vbNullChar & vbNullChar

    buffer = String$(32767, vbNullChar)
    titleBuffer = String$(1024, vbNullChar)

    With ofn
        .lStructSize = LenB(ofn)
        .lpstrFilter = StrPtr(filter)
        .nFilterIndex = 1
        .lpstrFile = StrPtr(buffer)
        .nMaxFile = Len(buffer)
        .lpstrFileTitle = StrPtr(titleBuffer)
        .nMaxFileTitle = Len(titleBuffer)
        .lpstrTitle = StrPtr(dialogTitle)
        .Flags = OFN_EXPLORER Or OFN_FILEMUSTEXIST Or _
                 OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY

        If initialFolder <> "" Then .lpstrInitialDir = StrPtr(initialFolder)
    End With

    If GetOpenFileNameW(ofn) = 0 Then
        PickOneDrawing = ""
        Exit Function
    End If

    p = InStr(1, buffer, vbNullChar)

    If p > 0 Then
        PickOneDrawing = Left$(buffer, p - 1)
    Else
        PickOneDrawing = buffer
    End If

End Function

Private Function GetAllDrawingsInFolder( _
    ByVal folderPath As String, _
    ByVal mainPath As String) As Collection

    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim files As New Collection
    Dim seen As Object
    Dim fullPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    If Len(folderPath) = 0 Or Not fso.FolderExists(folderPath) Then
        Set GetAllDrawingsInFolder = files
        Exit Function
    End If

    Set folder = fso.GetFolder(folderPath)
    For Each file In folder.files
        If LCase$(fso.GetExtensionName(file.name)) = "slddrw" Then
            If Left$(file.name, 2) = "~$" Then
                Debug.Print "ВРЕМЕННЫЙ ФАЙЛ ПРОПУЩЕН: " & file.Path
            ElseIf LCase$(file.name) Like "merged_########_######*.slddrw" Then
                Debug.Print "ПРЕЖНИЙ РЕЗУЛЬТАТ ПРОПУЩЕН: " & file.Path
            Else
            fullPath = file.Path
            If StrComp(Replace(fullPath, "/", "\"), Replace(mainPath, "/", "\"), vbTextCompare) <> 0 Then
                If Not seen.Exists(fullPath) Then
                    seen.Add fullPath, True
                    files.Add fullPath
                End If
            Else
                Debug.Print "Основной чертёж автоматически исключён: " & fullPath
            End If
            End If
        End If
    Next file

    Set GetAllDrawingsInFolder = files
End Function

' =========================================================
' ДИАЛОГ ВЫБОРА НЕСКОЛЬКИХ ЧЕРТЕЖЕЙ
' =========================================================

Private Function PickMultipleDrawings( _
    ByVal dialogTitle As String, _
    ByVal initialFolder As String) As Collection

    Dim ofn As OPENFILENAMEW
    Dim filter As String
    Dim buffer As String
    Dim titleBuffer As String
    Dim raw As String
    Dim parts As Variant
    Dim files As New Collection
    Dim folder As String
    Dim i As Long
    Dim p As Long

    filter = "Чертежи SOLIDWORKS (*.slddrw)" & vbNullChar & _
             "*.slddrw" & vbNullChar & _
             "Все файлы (*.*)" & vbNullChar & _
             "*.*" & vbNullChar & vbNullChar

    buffer = String$(32767, vbNullChar)
    titleBuffer = String$(1024, vbNullChar)

    With ofn
        .lStructSize = LenB(ofn)
        .lpstrFilter = StrPtr(filter)
        .nFilterIndex = 1
        .lpstrFile = StrPtr(buffer)
        .nMaxFile = Len(buffer)
        .lpstrFileTitle = StrPtr(titleBuffer)
        .nMaxFileTitle = Len(titleBuffer)
        .lpstrTitle = StrPtr(dialogTitle)
        .Flags = OFN_EXPLORER Or OFN_ALLOWMULTISELECT Or _
                 OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST Or _
                 OFN_HIDEREADONLY

        If initialFolder <> "" Then .lpstrInitialDir = StrPtr(initialFolder)
    End With

    If GetOpenFileNameW(ofn) = 0 Then
        Set PickMultipleDrawings = Nothing
        Exit Function
    End If

    p = InStr(1, buffer, vbNullChar & vbNullChar)

    If p > 0 Then
        raw = Left$(buffer, p - 1)
    Else
        raw = buffer
    End If

    parts = Split(raw, vbNullChar)

    If UBound(parts) = 0 Then

        If Len(CStr(parts(0))) > 0 Then
            files.Add CStr(parts(0))
        End If

    Else

        folder = CStr(parts(0))

        For i = 1 To UBound(parts)
            If Len(CStr(parts(i))) > 0 Then
                files.Add folder & "\" & CStr(parts(i))
            End If
        Next i

    End If

    Set PickMultipleDrawings = files

End Function

' =========================================================
' ДИАЛОГ СОХРАНЕНИЯ PDF
' =========================================================

Private Function GetFolder(ByVal filePath As String) As String

    Dim p As Long

    p = InStrRev(filePath, "\")

    If p > 0 Then
        GetFolder = Left$(filePath, p - 1)
    Else
        GetFolder = ""
    End If

End Function

Private Function GetBaseName(ByVal filePath As String) As String

    Dim name As String
    Dim p As Long

    p = InStrRev(filePath, "\")

    If p > 0 Then
        name = Mid$(filePath, p + 1)
    Else
        name = filePath
    End If

    p = InStrRev(name, ".")

    If p > 0 Then name = Left$(name, p - 1)

    GetBaseName = name

End Function






' =========================================================
' ЭКСПОРТ ЧЕРТЕЖА В PDF
' =========================================================

Private Function ExportDrawingToPDF( _
    ByVal swModel As Object, _
    ByVal pdfPath As String, _
    ByRef exportErrors As Long, _
    ByRef exportWarnings As Long) As Boolean

    On Error GoTo ExportError

    Dim swAppLocal As Object
    Dim pdfDataLocal As Object
    Dim modelPathLocal As String
    Dim sheetNamesLocal As Variant
    Dim iLocal As Long
    Dim saveResult As Boolean

    ExportDrawingToPDF = False
    exportErrors = 0
    exportWarnings = 0

    Debug.Print String(60, "=")
    Debug.Print "Начало экспорта PDF"

    If swModel Is Nothing Then
        Debug.Print "ОШИБКА: документ не передан в функцию."
        Exit Function
    End If

    If swModel.GetType <> 3 Then
        Debug.Print "ОШИБКА: переданный документ не является чертежом."
        Exit Function
    End If

    modelPathLocal = swModel.GetPathName

    If Len(modelPathLocal) = 0 Then
        Debug.Print "ОШИБКА: чертёж не сохранён."
        Exit Function
    End If

    If Len(pdfPath) = 0 Then
        Debug.Print "ОШИБКА: путь PDF не задан."
        Exit Function
    End If

    Set swAppLocal = Application.SldWorks
    sheetNamesLocal = swModel.GetSheetNames

    Debug.Print "Чертёж: " & modelPathLocal
    Debug.Print "Количество листов: " & _
        (UBound(sheetNamesLocal) - LBound(sheetNamesLocal) + 1)

    For iLocal = LBound(sheetNamesLocal) To UBound(sheetNamesLocal)
        Debug.Print "Лист: " & CStr(sheetNamesLocal(iLocal))
    Next iLocal

    Debug.Print "Путь PDF: " & pdfPath

    ' Используем проверенный способ экспорта без вызова SetSheets.
    Set pdfDataLocal = swAppLocal.GetExportFileData(1)

    If pdfDataLocal Is Nothing Then
        Debug.Print "ОШИБКА: не удалось создать настройки экспорта PDF."
        Exit Function
    End If

    pdfDataLocal.ExportAs3D = False
    pdfDataLocal.ViewPdfAfterSaving = False

    Debug.Print "Запуск SaveAs..."

    saveResult = swModel.extension.SaveAs( _
        pdfPath, _
        0, _
        1, _
        pdfDataLocal, _
        exportErrors, _
        exportWarnings)

    Debug.Print "SaveAs: " & CStr(saveResult)
    Debug.Print "Ошибки: " & CStr(exportErrors)
    Debug.Print "Предупреждения: " & CStr(exportWarnings)

    If saveResult And exportErrors = 0 Then
        ExportDrawingToPDF = True
        Debug.Print "УСПЕХ: PDF сохранён."
        Debug.Print "Ожидаемое количество листов: " & _
            (UBound(sheetNamesLocal) - LBound(sheetNamesLocal) + 1)
    Else
        Debug.Print "ОШИБКА: не удалось сохранить PDF."
    End If

    Debug.Print "Конец экспорта PDF"
    Exit Function

ExportError:
    Debug.Print "ОШИБКА VBA при экспорте: " & _
                Err.Number & " — " & Err.Description
    ExportDrawingToPDF = False

End Function
