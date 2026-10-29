Attribute VB_Name = "InsertWeldWithHoles"
Option Explicit

' Replace ONLY the InsertWeldWithHoles module. Entry point: start.
' API lengths are metres. No Save/SaveAs is performed.
Private Const DETAIL_SCALE As Double = 0.2
Private Const GAP As Double = 0.015
Private Const SIDE_MARGIN As Double = 0.02
Private Const BOTTOM_MARGIN As Double = 0.04
Private Const EPS As Double = 0.0000001
Private Const VIEW_PREFIX As String = "DRILL_"
' Optional: supply a local/network cut-list template to create a table and balloons.
' Empty = views only. Existing tables are never deleted or replaced.
Private Const TABLE_TEMPLATE As String = ""

Private app As SldWorks.SldWorks
Private drawingModel As SldWorks.ModelDoc2
Private drawing As SldWorks.DrawingDoc
Private partModel As SldWorks.ModelDoc2
Private oldConfig As String
Private oldOrientation As SldWorks.MathTransform
Private tempName As String
Private madeViews As Collection
Private madeTable As SldWorks.WeldmentCutListAnnotation
Private visited As Object

Public Sub start()
    Dim source As SldWorks.View, sheet As SldWorks.Sheet
    Dim config As String, modelPath As String, message As String
    Dim bodies As Collection, item As Variant, body As SldWorks.Body2
    Dim face As SldWorks.Face2, newView As SldWorks.View
    Dim cursorY As Double, width As Double, height As Double
    Dim props As Variant, tableEnabled As Boolean
    Dim i As Long, errText As String, cleanupText As String
    On Error GoTo Failed
    Set app = Application.SldWorks
    Set drawingModel = app.ActiveDoc
    Set partModel = Nothing
    Set oldOrientation = Nothing
    Set madeTable = Nothing
    Set madeViews = New Collection
    oldConfig = ""
    tempName = ""
    If drawingModel Is Nothing Then Fail "Откройте чертёж рамы."
    If drawingModel.GetType <> swDocDRAWING Then Fail "Активный документ должен быть чертежом."
    Set drawing = drawingModel
    Set source = SourceView()
    Set sheet = drawing.GetCurrentSheet
    If source.Type <> swDrawingNamedView Then Fail "Выберите обычный вид модели, а не разрез или проекционный вид."
    Set partModel = source.ReferencedDocument
    If partModel Is Nothing Then Fail "Загрузите модель исходного вида (не отсоединённый/облегчённый чертёж)."
    If partModel.GetType <> swDocPART Then Fail "Исходный вид должен ссылаться на многотельную деталь."
    modelPath = partModel.GetPathName
    If Len(modelPath) = 0 Then Fail "Сначала сохраните исходную деталь."
    config = source.ReferencedConfiguration
    oldConfig = partModel.ConfigurationManager.ActiveConfiguration.Name
    Set oldOrientation = partModel.ActiveView.Orientation3
    If StrComp(oldConfig, config, vbBinaryCompare) <> 0 Then
        If partModel.ShowConfiguration2(config) = False Then Fail "Не удалось включить конфигурацию вида: " & config
    End If
    Set bodies = New Collection
    Set visited = CreateObject("Scripting.Dictionary")
    UpdateCutList
    CollectBodies partModel.FirstFeature, bodies, False
    If bodies.Count = 0 Then
        message = "В списке вырезов нет тел с круглыми внутренними контурами на плоской грани."
        GoTo Finished
    End If
    ' No silent duplicates on a second run. Delete/rename previous generated views explicitly.
    CheckPreviousRun sheet
    props = sheet.GetProperties
    width = CDbl(props(5))
    height = CDbl(props(6))
    If width <= 2# * SIDE_MARGIN Or height <= BOTTOM_MARGIN Then Fail "Неверные размеры листа."
    cursorY = LowestViewBottom(sheet, height) - GAP
    tableEnabled = (Len(TABLE_TEMPLATE) > 0)
    If tableEnabled Then
        If Len(Dir$(TABLE_TEMPLATE)) = 0 Then Fail "Не найден шаблон таблицы: " & TABLE_TEMPLATE
    End If
    If MsgBox("Создать видов: " & CStr(bodies.Count) & ". Масштаб 1:" & CStr(1# / DETAIL_SCALE) & _
              ". Исходная модель и чертёж автоматически не сохраняются. Продолжить?", vbQuestion + vbYesNo) <> vbYes Then
        message = "Отменено."
        GoTo Finished
    End If
    If tableEnabled Then
        Set madeTable = source.InsertWeldmentTable(False, SIDE_MARGIN, height - SIDE_MARGIN, _
                                                 1, config, TABLE_TEMPLATE)
        If madeTable Is Nothing Then Fail "Не удалось вставить таблицу списка вырезов."
    End If
    For i = 1 To bodies.Count
        Set body = bodies(i)
        Set face = BestHoleFace(body)
        If face Is Nothing Then Fail "Геометрия тела изменилась: " & body.Name
        Set newView = CreateDetail(body, face, modelPath, config, i, VIEW_PREFIX & sheet.GetName & "_" & Format$(i, "000"))
        PlaceDetail newView, cursorY, width
        If tableEnabled Then AddBalloon newView, face
    Next i
    message = "Создано видов: " & CStr(madeViews.Count) & ". Проверьте отверстия и расположение перед сохранением."
    If tableEnabled = False Then message = message & vbCrLf & "Таблица и позиции отключены: TABLE_TEMPLATE не задан."
    GoTo Finished
Failed:
    errText = Err.Description
    cleanupText = RollbackDrawing()
    message = "Операция остановлена: " & errText & vbCrLf & cleanupText
Finished:
    cleanupText = RestoreModel()
    If Len(cleanupText) > 0 Then message = message & vbCrLf & cleanupText
    On Error Resume Next
    If Not drawingModel Is Nothing Then
        ActivateModel drawingModel
        drawingModel.ClearSelection2 True
        drawingModel.GraphicsRedraw2
    End If
    On Error GoTo 0
    MsgBox message, vbInformation, "Сверловка"
End Sub

Private Sub Fail(ByVal text As String)
    Err.Raise vbObjectError + 2101, "InsertWeldWithHoles", text
End Sub

Private Function HasItems(ByVal values As Variant) As Boolean
    On Error GoTo EmptyArray
    If IsArray(values) Then HasItems = (UBound(values) >= LBound(values))
EmptyArray:
End Function

Private Function SourceView() As SldWorks.View
    Dim manager As SldWorks.SelectionMgr, candidate As SldWorks.View
    Dim first As SldWorks.View, count As Long
    Set manager = drawingModel.SelectionManager
    count = manager.GetSelectedObjectCount2(-1)
    If count > 0 Then
        If count <> 1 Then Fail "Выберите один исходный вид рамы."
        If manager.GetSelectedObjectType3(1, -1) <> swSelDRAWINGVIEWS Then Fail "Выберите именно вид рамы."
        Set SourceView = manager.GetSelectedObject6(1, -1)
    Else
        Set first = drawing.GetFirstView
        Set candidate = first.GetNextView
        If candidate Is Nothing Then Fail "На текущем листе нет видов модели."
        Set first = candidate.GetNextView
        If Not first Is Nothing Then Fail "На листе несколько видов. Выберите исходный вид рамы."
        Set SourceView = candidate
    End If
End Function

Private Sub UpdateCutList()
    Dim feature As SldWorks.Feature, folder As SldWorks.BodyFolder
    Set feature = partModel.FirstFeature
    Do While Not feature Is Nothing
        If feature.GetTypeName2 = "SolidBodyFolder" Then
            Set folder = feature.GetSpecificFeature2
            If Not folder Is Nothing Then
                If folder.UpdateCutList() = False Then Fail "Не удалось обновить список вырезов."
            End If
        End If
        Set feature = feature.GetNextFeature
    Loop
End Sub

Private Sub CollectBodies(ByVal feature As SldWorks.Feature, ByVal result As Collection, ByVal subLevel As Boolean)
    Dim folder As SldWorks.BodyFolder, values As Variant
    Dim body As SldWorks.Body2, face As SldWorks.Face2, i As Long
    Do While Not feature Is Nothing
        If feature.GetTypeName2 = "CutListFolder" Then
            If visited.Exists(feature.Name) = False Then
                visited.Add feature.Name, True
                Set folder = feature.GetSpecificFeature2
                If Not folder Is Nothing Then
                    values = folder.GetBodies
                    If HasItems(values) Then
                        For i = LBound(values) To UBound(values)
                            Set body = values(i)
                            If Not body Is Nothing Then
                                Set face = BestHoleFace(body)
                                If Not face Is Nothing Then
                                    result.Add body
                                    Exit For ' One representative per cut-list item, not per instance.
                                End If
                            End If
                        Next i
                    End If
                End If
            End If
        Else
            CollectBodies feature.GetFirstSubFeature, result, True
        End If
        If subLevel Then
            Set feature = feature.GetNextSubFeature
        Else
            Set feature = feature.GetNextFeature
        End If
    Loop
End Sub

Private Function BestHoleFace(ByVal body As SldWorks.Body2) As SldWorks.Face2
    Dim face As SldWorks.Face2, surface As SldWorks.Surface
    Dim holeLoop As SldWorks.Loop2, count As Long, bestCount As Long, area As Double
    Set face = body.GetFirstFace
    Do While Not face Is Nothing
        Set surface = face.GetSurface
        If surface.IsPlane Then
            count = 0
            Set holeLoop = face.GetFirstLoop
            Do While Not holeLoop Is Nothing
                If IsRoundOpening(holeLoop) Then count = count + 1
                Set holeLoop = holeLoop.GetNext
            Loop
            If count > 0 Then
                If count > bestCount Or (count = bestCount And face.GetArea > area) Then
                    bestCount = count
                    area = face.GetArea
                    Set BestHoleFace = face
                End If
            End If
        End If
        Set face = face.GetNextFace
    Loop
End Function

Private Function IsRoundOpening(ByVal holeLoop As SldWorks.Loop2) As Boolean
    Dim edges As Variant, edge As SldWorks.Edge, curve As SldWorks.Curve
    Dim circleData As Variant, first As Variant, i As Long, k As Long, dot As Double
    If holeLoop.IsOuter Then Exit Function
    edges = holeLoop.GetEdges
    If HasItems(edges) = False Then Exit Function
    For i = LBound(edges) To UBound(edges)
        Set edge = edges(i)
        Set curve = edge.GetCurve
        If curve.IsCircle = False Then Exit Function
        circleData = curve.CircleParams
        If circleData(6) <= EPS Then Exit Function
        If i = LBound(edges) Then
            first = circleData
        Else
            For k = 0 To 2
                If Abs(circleData(k) - first(k)) > EPS Then Exit Function
            Next k
            If Abs(circleData(6) - first(6)) > EPS Then Exit Function
            dot = circleData(3) * first(3) + circleData(4) * first(4) + circleData(5) * first(5)
            If Abs(Abs(dot) - 1#) > EPS Then Exit Function
        End If
    Next i
    ' A closed inner BREP holeLoop made of arcs of the same circleData is a circular opening.
    ' Includes tube bores and countersinks; it is not proof of a machining operation.
    IsRoundOpening = True
End Function

Private Sub ActivateModel(ByVal model As SldWorks.ModelDoc2)
    Dim active As SldWorks.ModelDoc2, errors As Long
    Set active = app.ActivateDoc3(model.GetTitle, False, swDontRebuildActiveDoc, errors)
    If active Is Nothing Then Fail "Не удалось активировать документ: " & model.GetTitle
    If Not active Is model Then Fail "Активирован другой документ с таким же именем."
End Sub

Private Function FaceTransform(ByVal face As SldWorks.Face2) As SldWorks.MathTransform
    Dim utility As SldWorks.MathUtility, data As Variant
    Set utility = app.GetMathUtility
    data = BasisData(face.Normal, LongestDirection(face))
    Set FaceTransform = utility.CreateTransform(data)
End Function

Private Function BasisData(ByVal normal As Variant, ByVal axis As Variant) As Variant
    Dim x As Variant, y As Variant, data(15) As Double, k As Long
    normal = UnitVector(normal)
    ' Pick a stable in-plane X axis from the longest straight edge of this face.
    x = ProjectDirection(axis, normal)
    If VectorLength(x) < EPS Then
        axis = Array(1#, 0#, 0#)
        If Abs(normal(0)) > 0.9 Then axis = Array(0#, 1#, 0#)
        x = ProjectDirection(axis, normal)
    End If
    x = UnitVector(x)
    y = CrossVector(normal, x)
    For k = 0 To 2
        data(k * 3) = x(k)
        data(k * 3 + 1) = y(k)
        data(k * 3 + 2) = normal(k)
    Next k
    data(12) = 1#
    BasisData = data
End Function

Private Function LongestDirection(ByVal face As SldWorks.Face2) As Variant
    Dim edges As Variant, edge As SldWorks.Edge, curve As SldWorks.Curve
    Dim first As SldWorks.Vertex, last As SldWorks.Vertex, a As Variant, b As Variant
    Dim candidate As Variant, best As Variant, length As Double, i As Long, k As Long
    best = Array(1#, 0#, 0#)
    edges = face.GetEdges
    If HasItems(edges) Then
        For i = LBound(edges) To UBound(edges)
            Set edge = edges(i)
            Set curve = edge.GetCurve
            If curve.IsLine Then
                Set first = edge.GetStartVertex
                Set last = edge.GetEndVertex
                If Not first Is Nothing And Not last Is Nothing Then
                    a = first.GetPoint
                    b = last.GetPoint
                    candidate = Array(b(0) - a(0), b(1) - a(1), b(2) - a(2))
                    If VectorLength(candidate) > length Then
                        length = VectorLength(candidate)
                        best = candidate
                    End If
                End If
            End If
        Next i
    End If
    For k = 0 To 2
        If Abs(best(k)) > EPS Then
            If best(k) < 0# Then best = Array(-best(0), -best(1), -best(2))
            Exit For
        End If
    Next k
    LongestDirection = best
End Function

Private Function ProjectDirection(ByVal direction As Variant, ByVal normal As Variant) As Variant
    Dim dot As Double
    dot = direction(0) * normal(0) + direction(1) * normal(1) + direction(2) * normal(2)
    ProjectDirection = Array(direction(0) - dot * normal(0), direction(1) - dot * normal(1), direction(2) - dot * normal(2))
End Function

Private Function CrossVector(ByVal a As Variant, ByVal b As Variant) As Variant
    CrossVector = Array(a(1) * b(2) - a(2) * b(1), a(2) * b(0) - a(0) * b(2), a(0) * b(1) - a(1) * b(0))
End Function

Private Function VectorLength(ByVal value As Variant) As Double
    VectorLength = Sqr(value(0) ^ 2 + value(1) ^ 2 + value(2) ^ 2)
End Function

Private Function UnitVector(ByVal value As Variant) As Variant
    Dim length As Double
    length = VectorLength(value)
    If length < EPS Then Fail "Вырожденное направление грани."
    UnitVector = Array(value(0) / length, value(1) / length, value(2) / length)
End Function

Private Function MathVectorOf(ByVal value As Variant) As SldWorks.MathVector
    Dim xyz(2) As Double, utility As SldWorks.MathUtility, k As Long
    For k = 0 To 2
        xyz(k) = CDbl(value(k))
    Next k
    Set utility = app.GetMathUtility
    Set MathVectorOf = utility.CreateVector(xyz)
End Function

Private Function CreateDetail(ByVal body As SldWorks.Body2, ByVal face As SldWorks.Face2, _
                              ByVal modelPath As String, ByVal config As String, ByVal index As Long, ByVal viewName As String) As SldWorks.View
    Dim view As SldWorks.View, singleBody(0) As SldWorks.Body2
    Dim utility As SldWorks.MathUtility, normal As SldWorks.MathVector, v As Variant
    ActivateModel partModel
    Dim proposedName As String
    proposedName = "__DRILL_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(index)
    EnsureUnusedName proposedName
    tempName = proposedName
    partModel.ActiveView.Orientation3 = FaceTransform(face)
    partModel.NameView tempName
    ActivateModel drawingModel
    Set view = drawing.CreateDrawViewFromModelView3(modelPath, tempName, SIDE_MARGIN, SIDE_MARGIN, 0#)
    If view Is Nothing Then Fail "Не удалось создать вид тела: " & body.Name
    madeViews.Add view
    view.ReferencedConfiguration = config
    Set singleBody(0) = body
    view.Bodies = singleBody
    view.UseSheetScale = False
    view.ScaleDecimal = DETAIL_SCALE
    If view.SetName2(viewName) = False Then Fail "Не удалось присвоить имя виду."
    ' Test visibility with hidden lines REMOVED, not with hidden edges displayed.
    If view.SetDisplayMode4(False, swHIDDEN, False, False, True) = False Then Fail "Не удалось установить отображение вида."
    view.SetDisplayTangentEdges2 swTangentEdgesHidden
    drawingModel.EditRebuild3
    VerifySingleBody view, body
    Set utility = app.GetMathUtility
    Set normal = MathVectorOf(face.Normal)
    Set normal = normal.MultiplyTransform(view.ModelToViewTransform)
    v = UnitVector(normal.ArrayData)
    If Abs(v(0)) > 0.0001 Or Abs(v(1)) > 0.0001 Or v(2) < 0.9999 Then
        Fail "Направление отверстия не совпало с направлением вида: " & body.Name
    End If
    If HasVisibleOpening(view, face) = False Then Fail "Не подтверждена видимость круглого отверстия: " & body.Name
    If view.SetDisplayMode4(False, swHIDDEN_GREYED, False, False, True) = False Then Fail "Не удалось включить невидимые линии."
    ' Restore the model orientation and remove the temporary named orientation.
    ActivateModel partModel
    partModel.ActiveView.Orientation3 = oldOrientation
    If partModel.DeleteNamedView(tempName) = False Then Fail "Не удалось удалить временный вид модели."
    tempName = ""
    ActivateModel drawingModel
    Set CreateDetail = view
End Function

Private Function HasVisibleOpening(ByVal view As SldWorks.View, ByVal face As SldWorks.Face2) As Boolean
    Dim components As Variant, visible As Collection, entities As Variant
    Dim component As SldWorks.Component2, holeLoop As SldWorks.Loop2
    Dim edges As Variant, edge As SldWorks.Edge, mapped As Object, entity As Object
    Dim i As Long, j As Long, found As Boolean, complete As Boolean
    Set visible = New Collection
    components = view.GetVisibleComponents
    If HasItems(components) = False Then Exit Function
    For i = LBound(components) To UBound(components)
        Set component = components(i)
        entities = view.GetVisibleEntities2(component, swViewEntityType_Edge)
        If HasItems(entities) Then
            For j = LBound(entities) To UBound(entities)
                visible.Add entities(j)
            Next j
        End If
    Next i
    Set holeLoop = face.GetFirstLoop
    Do While Not holeLoop Is Nothing
        If IsRoundOpening(holeLoop) Then
            edges = holeLoop.GetEdges
            complete = True
            For i = LBound(edges) To UBound(edges)
                Set edge = edges(i)
                Set mapped = view.GetCorrespondingEntity(edge)
                found = False
                For Each entity In visible
                    If app.IsSame(entity, edge) = swObjectSame Then found = True
                    If Not mapped Is Nothing Then
                        If app.IsSame(entity, mapped) = swObjectSame Then found = True
                    End If
                    If found Then Exit For
                Next entity
                If found = False Then
                    complete = False
                    Exit For
                End If
            Next i
            If complete Then
                HasVisibleOpening = True
                Exit Function
            End If
        End If
        Set holeLoop = holeLoop.GetNext
    Loop
End Function

Private Sub EnsureUnusedName(ByVal name As String)
    Dim names As Variant, i As Long
    names = partModel.GetModelViewNames
    If HasItems(names) Then
        For i = LBound(names) To UBound(names)
            If StrComp(CStr(names(i)), name, vbTextCompare) = 0 Then
                Fail "Имя временного вида уже занято. Повторите запуск позже."
            End If
        Next i
    End If
End Sub

Private Sub VerifySingleBody(ByVal view As SldWorks.View, ByVal expected As SldWorks.Body2)
    Dim bodies As Variant, actual As SldWorks.Body2
    bodies = view.Bodies
    If HasItems(bodies) = False Then Fail "Вид не вернул изолированное тело."
    If UBound(bodies) - LBound(bodies) <> 0 Then Fail "В вид попало несколько тел."
    Set actual = bodies(LBound(bodies))
    If app.IsSame(actual, expected) <> swObjectSame Then Fail "В вид попало другое тело."
End Sub

Private Sub CheckPreviousRun(ByVal sheet As SldWorks.Sheet)
    Dim views As Variant, view As SldWorks.View, i As Long
    views = sheet.GetViews
    If HasItems(views) Then
        For i = LBound(views) To UBound(views)
            Set view = views(i)
            If Left$(view.GetName2, Len(VIEW_PREFIX)) = VIEW_PREFIX Then
                Fail "На листе уже есть виды DRILL_. Удалите их перед повторным запуском."
            End If
        Next i
    End If
End Sub

Private Function LowestViewBottom(ByVal sheet As SldWorks.Sheet, ByVal height As Double) As Double
    Dim views As Variant, view As SldWorks.View, outline As Variant, i As Long
    LowestViewBottom = height - SIDE_MARGIN
    views = sheet.GetViews
    If HasItems(views) Then
        For i = LBound(views) To UBound(views)
            Set view = views(i)
            If view.Type <> swDrawingSheet Then
                outline = view.GetOutline
                If outline(1) < LowestViewBottom Then LowestViewBottom = outline(1)
            End If
        Next i
    End If
End Function

Private Sub PlaceDetail(ByVal view As SldWorks.View, ByRef cursorY As Double, ByVal sheetWidth As Double)
    Dim outline As Variant, position As Variant, width As Double, height As Double
    Dim angle As Double, result As Boolean
    outline = view.GetOutline
    width = outline(2) - outline(0)
    height = outline(3) - outline(1)
    If height > width Then
        drawingModel.ClearSelection2 True
        If drawingModel.Extension.SelectByID2(view.GetName2, "DRAWINGVIEW", 0#, 0#, 0#, False, 0, Nothing, 0) = False Then Fail "Не удалось выбрать вид для поворота."
        result = drawing.DrawingViewRotate(view.Angle + 1.5707963267949)
        If result = False Then Fail "Не удалось повернуть вид горизонтально."
        drawingModel.EditRebuild3
        outline = view.GetOutline
        width = outline(2) - outline(0)
        height = outline(3) - outline(1)
    End If
    If width <= EPS Or height <= EPS Then Fail "Пустые габариты вида."
    If width > sheetWidth - 2# * SIDE_MARGIN Then Fail "Вид шире листа. Уменьшите DETAIL_SCALE."
    If cursorY - height < BOTTOM_MARGIN Then Fail "Не хватает места под видами. Освободите лист или уменьшите DETAIL_SCALE."
    position = view.Position
    position(0) = position(0) + SIDE_MARGIN - outline(0)
    position(1) = position(1) + cursorY - outline(3)
    view.Position = position
    drawingModel.EditRebuild3
    outline = view.GetOutline
    If Abs(outline(0) - SIDE_MARGIN) > 0.0001 Or Abs(outline(3) - cursorY) > 0.0001 Then Fail "SOLIDWORKS не применил положение вида."
    cursorY = outline(1) - GAP
End Sub

Private Sub AddBalloon(ByVal view As SldWorks.View, ByVal face As SldWorks.Face2)
    Dim entity As SldWorks.Entity, data As SldWorks.SelectData
    Dim options As SldWorks.BalloonOptions, note As SldWorks.Note
    Dim annotation As SldWorks.Annotation, outline As Variant, position As Variant
    drawingModel.ClearSelection2 True
    Set entity = face
    Set data = drawingModel.SelectionManager.CreateSelectData
    Set data.View = view
    If entity.Select4(False, data) = False Then Fail "Не удалось выбрать грань для позиции."
    Set options = drawingModel.Extension.CreateBalloonOptions
    options.Style = swBS_None
    options.UpperTextContent = swBalloonTextItemNumber
    options.ShowQuantity = True
    options.QuantityPlacement = 3
    options.QuantityDenotationText = "шт"
    options.QuantityOverride = False
    Set note = drawingModel.Extension.InsertBOMBalloon2(options)
    If note Is Nothing Then Fail "Не удалось создать позицию списка вырезов."
    Set annotation = note.GetAnnotation
    outline = view.GetOutline
    annotation.SetPosition2 outline(0), outline(3) + GAP / 3#, 0#
    drawingModel.ClearSelection2 True
End Sub

Private Function RollbackDrawing() As String
    Dim i As Long, view As SldWorks.View, annotation As SldWorks.Annotation
    Dim table As SldWorks.TableAnnotation
    On Error GoTo Incomplete
    If madeViews Is Nothing Then Exit Function
    If drawingModel Is Nothing Then Exit Function
    ActivateModel drawingModel
    For i = madeViews.Count To 1 Step -1
        Set view = madeViews(i)
        drawingModel.ClearSelection2 True
        If drawingModel.Extension.SelectByID2(view.GetName2, "DRAWINGVIEW", 0#, 0#, 0#, False, 0, Nothing, 0) = False Then GoTo Incomplete
        If drawingModel.Extension.DeleteSelection2(swDelete_Absorbed) = False Then GoTo Incomplete
    Next i
    If Not madeTable Is Nothing Then
        Set table = madeTable
        Set annotation = table.GetAnnotation
        drawingModel.ClearSelection2 True
        If annotation.Select3(False, Nothing) = False Then GoTo Incomplete
        If drawingModel.Extension.DeleteSelection2(swDelete_Absorbed) = False Then GoTo Incomplete
    End If
    RollbackDrawing = "Созданные этим запуском виды и таблица удалены."
    Exit Function
Incomplete:
    RollbackDrawing = "Не удалось полностью удалить результат. Проверьте лист вручную; документ не сохранён."
End Function

Private Function RestoreModel() As String
    Dim warning As String, currentConfig As String, result As Boolean
    If partModel Is Nothing Then Exit Function
    ' Independent best-effort steps: failed deletion must not prevent configuration restoration.
    On Error Resume Next
    Err.Clear
    ActivateModel partModel
    If Err.Number <> 0 Then warning = warning & " Не удалось активировать модель."
    Err.Clear
    If Len(tempName) > 0 Then
        result = False
        result = partModel.DeleteNamedView(tempName)
        If Err.Number <> 0 Or result = False Then
            warning = warning & " Остался временный именованный вид: " & tempName & "."
        Else
            tempName = ""
        End If
    End If
    Err.Clear
    If Len(oldConfig) > 0 Then
        currentConfig = partModel.ConfigurationManager.ActiveConfiguration.Name
        If Err.Number <> 0 Then
            warning = warning & " Не прочитана конфигурация модели."
        ElseIf currentConfig <> oldConfig Then
            result = False
            result = partModel.ShowConfiguration2(oldConfig)
            If Err.Number <> 0 Or result = False Then warning = warning & " Не восстановлена конфигурация: " & oldConfig & "."
        End If
    End If
    Err.Clear
    If Not oldOrientation Is Nothing Then partModel.ActiveView.Orientation3 = oldOrientation
    If Err.Number <> 0 Then warning = warning & " Не восстановлена ориентация модели."
    On Error GoTo 0
    If Len(warning) > 0 Then RestoreModel = "Проверьте состояние модели; модель не сохранена." & warning
End Function

' Optional native arithmetic smoke test. Runs without modifying any document.
Public Sub TestGeometry()
    Dim utility As SldWorks.MathUtility, transform As SldWorks.MathTransform
    Dim vector As SldWorks.MathVector, normals As Variant, normal As Variant
    Dim data As Variant, result As Variant, i As Long
    On Error GoTo Failed
    Set app = Application.SldWorks
    Set utility = app.GetMathUtility
    normals = Array(Array(1#, 0#, 0#), Array(-1#, 0#, 0#), Array(0#, 1#, 0#), _
                    Array(0#, -1#, 0#), Array(0#, 0#, 1#), Array(0#, 0#, -1#), Array(1#, 2#, 3#))
    For i = LBound(normals) To UBound(normals)
        normal = UnitVector(normals(i))
        data = BasisData(normal, normal) ' Also exercises parallel-axis fallback.
        Set transform = utility.CreateTransform(data)
        Set vector = MathVectorOf(normal)
        Set vector = vector.MultiplyTransform(transform)
        result = vector.ArrayData
        If Abs(result(0)) > EPS Or Abs(result(1)) > EPS Or Abs(result(2) - 1#) > EPS Then
            Fail "TestGeometry: неверная матрица, случай " & CStr(i)
        End If
    Next i
    MsgBox "TestGeometry: 7 направлений проверены. Это не проверка чертежа и геометрии отверстий.", vbInformation
    Exit Sub
Failed:
    MsgBox Err.Description, vbCritical, "TestGeometry FAILED"
End Sub
