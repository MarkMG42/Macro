Attribute VB_Name = "MergeDrawingsV16_9"
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

' MERGE_DRAWINGS_V16_9_NATIVE_DIALOGS_20261027
' SOLIDWORKS 2025 x64 / VBA7. Entry point: main.
' Source files are read-only inputs. Only private working drawings are edited.
' Native section/detail features are transferred with their complete parent chain.
' A successful PasteSheet is never accepted without verifying the view graph.

Private Const BUILD_ID As String = "MERGE_DRAWINGS_V16_9_NATIVE_DIALOGS_20261027"
Private Const GEOM_TOL As Double = 0.000001 ' metres; 0.001 mm
Private Const VALUE_TOL As Double = 0.0000001
Private Const MAX_ITEMS As Long = 100000
Private runFolder As String
Private runToken As String
Private logStream As Object
Private fsoRun As Object
Private expectedResult As Object
Private runBusy As Boolean

Public Sub main()
    Dim mainPath As String, resultPath As String, pdfPath As String
    Dim sourceFiles As Collection, sourcePath As Variant
    Dim choice As VbMsgBoxResult, added As Long, fileIndex As Long
    Dim errors As Long, warnings As Long, saved As Boolean
    Dim actual As Object, failure As String, failureNumber As Long
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
    Set fsoRun = CreateObject("Scripting.FileSystemObject")

    mainPath = PickOneDrawing("Выберите ОСНОВНОЙ чертеж", "")
    If Len(mainPath) = 0 Then GoTo Finished
    CheckSavedInput mainPath
    InitRunFolder
    LogLine "BUILD_ID: " & BUILD_ID
    LogLine "SOLIDWORKS: " & swApp.RevisionNumber
    LogLine "MAIN: " & mainPath

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

    targetPath = runFolder & "\WORK_TARGET_" & runToken & ".slddrw"
    Set targetModel = CloneInput(mainPath, targetPath)
    Set targetDraw = targetModel
    Set expectedResult = CaptureDrawing(targetModel)
    LogLine "MAIN sheets=" & expectedResult("Sheets").Count & _
        "; views=" & expectedResult("Views").Count

    For Each sourcePath In sourceFiles
        fileIndex = fileIndex + 1
        added = added + AppendDrawing(CStr(sourcePath), fileIndex)
    Next sourcePath

    RequireActive targetModel
    RebuildDrawing targetModel
    Set actual = CaptureDrawing(targetModel)
    VerifyDrawing expectedResult, actual, "FINAL before save"
    WriteManifest actual, runFolder & "\verified_views.tsv"
    targetModel.ClearSelection2 True
    saved = targetModel.Save3(swSaveAsOptions_Silent, errors, warnings)
    LogLine "SAVE work: " & saved & "; errors=" & errors & "; warnings=" & warnings
    Require saved And errors = 0, "Не удалось сохранить рабочий чертеж: " & errors

    ' Verify persistence, not just the in-memory result. Close only our own copy.
    CloseOwnedDrawing targetModel
    Set targetDraw = Nothing
    Set targetModel = OpenWorkingDrawing(targetPath)
    Set targetDraw = targetModel
    RebuildDrawing targetModel
    Set actual = CaptureDrawing(targetModel)
    VerifyDrawing expectedResult, actual, "FINAL after reopen"
    ' Save the rebuilt, verified file before publishing a byte-for-byte copy.
    saved = targetModel.Save3(swSaveAsOptions_Silent, errors, warnings)
    Require saved And errors = 0, "Не удалось сохранить проверенный чертеж: " & errors
    CloseOwnedDrawing targetModel
    Set targetDraw = Nothing

    resultPath = runFolder & "\MERGED_" & runToken & ".slddrw"
    Require Not fsoRun.FileExists(resultPath), "Итоговый файл уже существует."
    fsoRun.CopyFile targetPath, resultPath, False
    targetPath = resultPath
    Set targetModel = OpenWorkingDrawing(resultPath)
    Set targetDraw = targetModel
    Set actual = CaptureDrawing(targetModel)
    VerifyDrawing expectedResult, actual, "PUBLISHED drawing"
    pdfPath = runFolder & "\WORK_PDF.pdf"
    RequireActive targetModel
    targetModel.ClearSelection2 True
    saved = ExportDrawingToPDF(targetModel, pdfPath, errors, warnings)
    LogLine "PDF: " & saved & "; errors=" & errors & "; warnings=" & warnings
    Require saved And errors = 0, "Чертеж проверен, но экспорт PDF не выполнен. Код: " & errors
    Require fsoRun.FileExists(pdfPath), "API сообщил об экспорте, но PDF отсутствует."
    Require fsoRun.GetFile(pdfPath).Size > 0, "Получен пустой файл PDF."
    fsoRun.CopyFile pdfPath, runFolder & "\MERGED_" & runToken & ".pdf", False
    LogLine "SUCCESS sheets=" & targetDraw.GetSheetCount & "; added=" & added
    WriteStatus "SUCCESS", "Проверены виды, родительские связи и повторное открытие."
    MsgBox "Объединение завершено." & vbCrLf & _
        "Добавлено листов: " & added & vbCrLf & _
        "Чертеж, PDF и журнал:" & vbCrLf & runFolder, vbInformation, "V16"
    GoTo Finished

Failed:
    failureNumber = Err.Number
    failure = Err.Description
    ' Capture the actual error before calling anything that can reset Err.
    On Error Resume Next
    LogLine "FAILED " & failureNumber & ": " & failure
    WriteStatus "FAILED", failure
    CloseOwnedDrawing targetModel
    Set targetDraw = Nothing
    If Not initialActive Is Nothing Then RequireActive initialActive
    MsgBox "Объединение остановлено: " & failure & vbCrLf & _
        "Исходные чертежи не перезаписаны." & vbCrLf & _
        "Журнал: " & runFolder & "\merge.log", vbCritical, "V16: проверка не пройдена"
    On Error GoTo 0
Finished:
    On Error Resume Next
    If Not logStream Is Nothing Then logStream.Close
    Set logStream = Nothing
    runBusy = False
    On Error GoTo 0
End Sub

Private Function AppendDrawing(ByVal inputPath As String, ByVal fileIndex As Long) As Long
    Dim source As SldWorks.ModelDoc2, drawing As SldWorks.DrawingDoc
    Dim original As Object, staged As Object, current As Object
    Dim viewMap As Object, sheetMap As Object, rec As Object, copied As Object
    Dim order As Collection, key As Variant, sheetKey As Variant, s As Object
    Dim newSheet As String, destSheet As String, name As String
    Dim view As SldWorks.View, dest As SldWorks.Sheet, item As Object
    Dim i As Long, n As Long, e As Long, message As String
    On Error GoTo Failed
    LogLine "SOURCE " & fileIndex & ": " & inputPath
    Set source = CloneInput(inputPath, runFolder & "\SOURCE_" & runToken & "_" & Format$(fileIndex, "0000") & ".slddrw")
    Set drawing = source
    Set original = CaptureDrawing(source)
    Set order = DependencyOrder(original("Views"))
    Set viewMap = NewDictionary()
    Set sheetMap = NewDictionary()

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
    RebuildDrawing source
    Set staged = CaptureDrawing(source)
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
        Set current = CaptureDrawing(targetModel)
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
        dest.CustomPropertyView = name
        LogLine "PROPERTY VIEW readback: sheet=" & destSheet & "; requested=[" & _
            name & "]; actual=[" & dest.CustomPropertyView & "]"
    Next sheetKey
    RebuildDrawing targetModel

    ' Build the expected final drawing from the ORIGINAL snapshots, not from
    ' whatever happened to survive the paste. Only generated names are mapped.
    Set current = CaptureDrawing(targetModel)
    For Each sheetKey In original("Order")
        Set rec = CloneRecord(original("Sheets")(CStr(sheetKey)))
        name = CStr(sheetMap(CStr(sheetKey)))
        rec("Name") = name
        rec("PropertyView") = MappedPropertyView(original, CStr(sheetKey), viewMap)
        expectedResult("Sheets").Add name, rec
        expectedResult("Order").Add name
    Next sheetKey
    For Each key In original("Views").Keys
        Set rec = CloneRecord(original("Views")(CStr(key)))
        name = CStr(viewMap(CStr(key)))
        rec("Name") = name
        rec("Sheet") = CStr(sheetMap(CStr(rec("Sheet"))))
        If Len(CStr(rec("Parent"))) > 0 Then rec("Parent") = CStr(viewMap(CStr(rec("Parent"))))
        ' PasteSheet(rename=yes) legitimately changes conflicting section labels.
        ' Keep the native new label, but never accept a changed parent identity.
        Set copied = current("Views")(name)
        rec("Label") = CStr(copied("Label"))
        expectedResult("Views").Add name, rec
    Next key
    VerifyDrawing expectedResult, current, "AFTER source " & fileIndex
    n = original("Sheets").Count
    CloseOwnedDrawing source
    AppendDrawing = n
    Exit Function
Failed:
    e = Err.Number
    message = Err.Description
    On Error Resume Next
    CloseOwnedDrawing source
    On Error GoTo 0
    Err.Raise e, "AppendDrawing", inputPath & vbCrLf & message
End Function

Private Function MappedPropertyView(ByVal snapshot As Object, ByVal sheetName As String, _
    ByVal viewMap As Object) As String
    Dim rec As Object, name As String, originalMode As String
    Set rec = snapshot("Sheets")(sheetName)
    name = CStr(rec("PropertyView"))
    originalMode = name
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

Private Function CaptureDrawing(ByVal model As SldWorks.ModelDoc2) As Object
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
    Set sheets = NewDictionary()
    Set views = NewDictionary()
    Set order = New Collection
    For i = LBound(names) To UBound(names)
        EnsureSheetActive model, CStr(names(i)), "CaptureDrawing"
        drawing.EditSheet
        Set sheet = drawing.GetCurrentSheet
        Require SameText(sheet.GetName, CStr(names(i))), "Активирован неверный лист."
        props = sheet.GetProperties2
        Require IsArray(props), "Не прочитаны параметры листа " & sheet.GetName
        Set sheetRec = NewDictionary()
        sheetRec.Add "Name", sheet.GetName
        sheetRec.Add "Width", CDbl(props(5))
        sheetRec.Add "Height", CDbl(props(6))
        sheetRec.Add "Scale1", CDbl(props(2))
        sheetRec.Add "Scale2", CDbl(props(3))
        sheetRec.Add "FirstAngle", FlagValue(CBool(props(4)))
        sheetRec.Add "PropertyView", sheet.CustomPropertyView
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
        If Not SameText(CStr(a("PropertyView")), CStr(b("PropertyView"))) Then
            LogLine "PROPERTY VIEW mismatch [" & stage & "]: sheet=" & CStr(key) & _
                "; expected=[" & CStr(a("PropertyView")) & "]; actual=[" & CStr(b("PropertyView")) & "]"
            Require False, stage & ": вид свойств листа " & CStr(key) & _
                "; ожидалось=[" & CStr(a("PropertyView")) & "]; фактически=[" & CStr(b("PropertyView")) & "]"
        End If
        Require CBool(a("FormatVisible")) = CBool(b("FormatVisible")), stage & ": видимость формата " & CStr(key)
        VerifyAnnotations a, b, stage & ": лист " & CStr(key)
    Next key
    For Each key In expected("Views").Keys
        Set a = expected("Views")(CStr(key))
        Set b = actual("Views")(CStr(key))
        Require SameText(CStr(a("Sheet")), CStr(b("Sheet"))), stage & ": вид на неверном листе " & CStr(key)
        Require SameViewFingerprint(a, b), stage & ": геометрия/настройки вида не совпадают: " & CStr(key)
        For Each field In Array("SheetScale", "ParentScale", "Locked", "Alignment", "Suppressed", "Hatches")
            Require CLng(a(field)) = CLng(b(field)), stage & ": " & CStr(field) & " / " & CStr(key)
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

Private Sub RebuildDrawing(ByVal model As SldWorks.ModelDoc2)
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

Private Function CloneInput(ByVal path As String, ByVal copyPath As String) As SldWorks.ModelDoc2
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
    fsoRun.CopyFile path, copyPath, False
    SetAttr copyPath, vbNormal
    oldFolder = swApp.GetCurrentWorkingDirectory
    EnsureWorkingDirectory fsoRun.GetParentFolderName(path), "CloneInput source"
    Set copy = OpenWorkingDrawing(copyPath)
    EnsureWorkingDirectory oldFolder, "CloneInput restore"
    oldFolder = ""
    LogLine "CLONE copy: " & copyPath
    Set actual = CaptureDrawing(copy)
    VerifyDrawing baseline, actual, "CLONE " & path
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
    Require Left$(full, Len(root)) = root, "Запрещено изменять файл вне папки запуска: " & path
End Sub

Private Sub InitRunFolder()
    Dim base As String, parent As String, token As String, suffix As Long
    base = GetFolder(swApp.GetCurrentMacroPathName)
    Require Len(base) > 0, "Сохраните макрос .swp в папку macro перед запуском."
    parent = fsoRun.BuildPath(base, "runs")
    If Not fsoRun.FolderExists(parent) Then fsoRun.CreateFolder parent
    token = Format$(Now, "yyyymmdd_hhnnss")
    runFolder = fsoRun.BuildPath(parent, token)
    Do While fsoRun.FolderExists(runFolder)
        suffix = suffix + 1
        runFolder = fsoRun.BuildPath(parent, token & "_" & CStr(suffix))
    Loop
    fsoRun.CreateFolder runFolder
    runToken = fsoRun.GetFileName(runFolder)
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

Private Function GetOrderedSourceFiles(ByVal selectedFiles As Collection, _
    ByVal mainPath As String) As Collection
    Dim ordered As New Collection, additions As Collection
    Dim answer As String, parts As Variant, verb As String
    Dim page As Long, pages As Long, first As Long, last As Long, i As Long
    Dim fromIndex As Long, toIndex As Long, moving As String, prompt As String
    Dim folder As String
    AddOrderCandidates ordered, selectedFiles, mainPath
    page = 1
    Do
        pages = (ordered.Count + 4) \ 5
        If pages = 0 Then pages = 1
        If page > pages Then page = pages
        If page < 1 Then page = 1
        first = (page - 1) * 5 + 1
        last = first + 4
        If last > ordered.Count Then last = ordered.Count
        prompt = "Основной чертёж остаётся первым." & vbCrLf & _
            "Источников: " & ordered.Count & "; страница " & page & "/" & pages & vbCrLf
        For i = first To last
            prompt = prompt & CStr(i) & ". " & OrderDisplayPath(CStr(ordered(i))) & vbCrLf
        Next i
        prompt = prompt & vbCrLf & "Введите команду (латиницей):" & vbCrLf & _
            "OK — принять; M 3 1 — переместить №3 на место №1" & vbCrLf & _
            "R 2 — убрать №2; V 2 — полный путь №2" & vbCrLf & _
            "A — добавить файлы; F — добавить папку" & vbCrLf & _
            "N / P — следующая / предыдущая страница" & vbCrLf & _
            "Отмена или пустой ввод — отменить объединение."
        answer = Trim$(VBA.InputBox(prompt, "Порядок чертежей V16.9", "OK"))
        If Len(answer) = 0 Then Exit Function
        ' Normalize whitespace only in commands, never in file paths.
        Do While InStr(answer, "  ") > 0
            answer = Replace(answer, "  ", " ")
        Loop
        parts = Split(UCase$(answer), " ")
        verb = CStr(parts(0))
        If verb = "OK" And UBound(parts) = 0 Then
            For i = 1 To ordered.Count
                Require fsoRun.FileExists(CStr(ordered(i))), "Файл исчез из списка: " & CStr(ordered(i))
            Next i
            LogLine "Порядок исходных чертежей (диалог VBA):"
            For i = 1 To ordered.Count
                LogLine CStr(i) & ". " & CStr(ordered(i))
            Next i
            Set GetOrderedSourceFiles = ordered
            Exit Function
        ElseIf verb = "M" And UBound(parts) = 2 Then
            fromIndex = OrderIndex(CStr(parts(1)), ordered.Count)
            toIndex = OrderIndex(CStr(parts(2)), ordered.Count)
            If fromIndex > 0 And toIndex > 0 Then
                If fromIndex <> toIndex Then
                    moving = CStr(ordered(fromIndex))
                    ordered.Remove fromIndex
                    If toIndex > ordered.Count Then
                        ordered.Add moving
                    Else
                        ordered.Add moving, , toIndex
                    End If
                End If
            Else
                MsgBox "Нужны два существующих номера. Пример: M 3 1", vbExclamation
            End If
        ElseIf (verb = "R" Or verb = "V") And UBound(parts) = 1 Then
            fromIndex = OrderIndex(CStr(parts(1)), ordered.Count)
            If fromIndex = 0 Then
                MsgBox "Нужен существующий номер чертежа.", vbExclamation
            ElseIf verb = "R" Then
                ordered.Remove fromIndex
            Else
                MsgBox CStr(ordered(fromIndex)), vbInformation, "Полный путь"
            End If
        ElseIf verb = "N" And UBound(parts) = 0 Then
            If page < pages Then page = page + 1
        ElseIf verb = "P" And UBound(parts) = 0 Then
            If page > 1 Then page = page - 1
        ElseIf verb = "A" And UBound(parts) = 0 Then
            Set additions = PickMultipleDrawings("Добавить исходные чертежи", GetFolder(mainPath))
            AddOrderCandidates ordered, additions, mainPath
        ElseIf verb = "F" And UBound(parts) = 0 Then
            folder = Trim$(VBA.InputBox("Вставьте полный путь к папке с чертежами:", "Добавить папку"))
            If Len(folder) > 0 Then
                If fsoRun.FolderExists(folder) Then
                    Set additions = GetAllDrawingsInFolder(folder, mainPath)
                    AddOrderCandidates ordered, additions, mainPath
                Else
                    MsgBox "Папка не найдена: " & folder, vbExclamation
                End If
            End If
        Else
            MsgBox "Команда не распознана. Примеры: OK, M 3 1, R 2, V 2, A, F, N, P.", vbExclamation
        End If
    Loop
End Function

Private Sub AddOrderCandidates(ByVal ordered As Collection, ByVal additions As Collection, _
    ByVal mainPath As String)
    Dim seen As Object, item As Variant, path As String, key As String
    If additions Is Nothing Then Exit Sub
    Set seen = NewDictionary()
    For Each item In ordered
        seen(NormalizePath(CStr(item))) = True
    Next item
    seen(NormalizePath(mainPath)) = True
    For Each item In additions
        path = Replace(Trim$(CStr(item)), "/", "\")
        key = NormalizePath(path)
        If Len(path) > 0 Then
            If LCase$(fsoRun.GetExtensionName(path)) = "slddrw" And _
                Left$(fsoRun.GetFileName(path), 2) <> "~$" Then
                If seen.Exists(key) = False Then
                    Require fsoRun.FileExists(path), "Файл не найден: " & path
                    seen.Add key, True
                    ordered.Add path
                End If
            End If
        End If
    Next item
End Sub

Private Function OrderIndex(ByVal text As String, ByVal maximum As Long) As Long
    Dim i As Long, digit As Long, value As Long
    If Len(text) = 0 Then Exit Function
    For i = 1 To Len(text)
        digit = AscW(Mid$(text, i, 1)) - AscW("0")
        If digit < 0 Or digit > 9 Then Exit Function
        ' Bound the multiplication before it can overflow a VBA Long.
        If value > maximum \ 10 Then Exit Function
        If value = maximum \ 10 And digit > maximum Mod 10 Then Exit Function
        value = value * 10 + digit
    Next i
    If value >= 1 And value <= maximum Then OrderIndex = value
End Function

Private Function OrderDisplayPath(ByVal path As String) As String
    If Len(path) > 90 Then
        OrderDisplayPath = "..." & Right$(path, 87)
    Else
        OrderDisplayPath = path
    End If
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
