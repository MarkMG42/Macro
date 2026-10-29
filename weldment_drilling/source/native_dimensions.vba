' Everything executes inside the existing SOLIDWORKS VBA host. No external runtime.
Private Function NativeDimensionView(ByVal view As SldWorks.View, ByVal targetRight As Double, ByVal targetTop As Double) As Variant
    Dim created As Collection, expected As Collection, lines() As NDLine, holes() As NDHole
    Dim lineCount As Long, holeCount As Long, bounds As NDBounds, viewScale As Double
    Dim selectedHoles As Collection, xPositions As Collection, yPositions As Collection, groups As Collection
    Dim i As Long, k As Long, index As Long, group As Collection, display As SldWorks.DisplayDimension
    Dim projectedBefore As Boolean, modeKnown As Boolean, originalPosition As Variant, position As Variant
    Dim box As Variant, actual As Variant, dx As Double, dy As Double, number As Long, description As String, origin As String
    Dim breakOutline As Variant
    Dim cleanup As String, prefix As String, breakApplied As Boolean, unbrokenWidth As Double
    On Error GoTo Failed
    Set created = New Collection
    Set expected = New Collection
    If view Is Nothing Then Fail "Нет вида для размеров."
    MarkStage "VBA-размеры: проверка вида"
    SelectDrawingView view
    If view.ReferencedDocument Is Nothing Then Fail "Не загружена модель вида."
    If view.ReferencedDocument.GetType <> swDocPART Then Fail "Размеры поддерживают вид одной детали."
    viewScale = view.ScaleDecimal
    If NdFinite(viewScale) = False Then Fail "Неконечный масштаб вида."
    If viewScale <= 0# Then Fail "Неположительный масштаб вида."
    If NdFinite(targetRight) = False Or NdFinite(targetTop) = False Then Fail "Неконечные координаты размещения."
    originalPosition = ReadPosition(view)
    projectedBefore = view.ProjectedDimensions
    modeKnown = True
    view.ProjectedDimensions = True
    If view.ProjectedDimensions = False Then Fail "Не включены проекционные размеры вида."
    drawingModel.EditRebuild3
    If view.ProjectedDimensions = False Then Fail "После перестроения сброшен проекционный режим."
    MarkStage "VBA-размеры: чтение геометрии"
    NdReadGeometry view, viewScale, lines, lineCount, holes, holeCount, bounds
    NdFindBounds lines, lineCount, bounds, viewScale
    Set selectedHoles = NdPrepareHoles(holes, holeCount, bounds, viewScale)
    If selectedHoles.Count = 0 Then Fail "Для размеров не найдено полного круглого ребра внутренней цилиндрической стенки. Разделённые дуги требуют отдельной проверки."
    Set xPositions = NdPositions(holes, selectedHoles, True, viewScale)
    Set yPositions = NdPositions(holes, selectedHoles, False, viewScale)
    Set groups = NdDiameterGroups(holes, selectedHoles, viewScale)
    WriteLog "VBA-размеры: отверстий=" & CStr(selectedHoles.Count) & "; X=" & CStr(xPositions.Count) & "; Y=" & CStr(yPositions.Count)
    MarkStage "VBA-размеры: создание"
    For i = 1 To xPositions.Count
        index = CLng(xPositions(i))
        NdAddLinear view, lines(bounds.LeftIndex).Entity, holes(index).Entity, True, _
                    (bounds.MinX + holes(index).X) / 2#, bounds.MinY - DIM_BASE_OFFSET - (i - 1) * DIM_LINE_STEP, _
                    holes(index).X - bounds.MinX, True, viewScale, created, expected
    Next i
    NdAddLinear view, lines(bounds.LeftIndex).Entity, lines(bounds.RightIndex).Entity, True, _
                (bounds.MinX + bounds.MaxX) / 2#, bounds.MinY - DIM_BASE_OFFSET - xPositions.Count * DIM_LINE_STEP, _
                bounds.MaxX - bounds.MinX, False, viewScale, created, expected
    For i = 1 To yPositions.Count
        index = CLng(yPositions(i))
        NdAddLinear view, lines(bounds.BottomIndex).Entity, holes(index).Entity, False, _
                    bounds.MaxX + DIM_BASE_OFFSET + (i - 1) * DIM_LINE_STEP, (bounds.MinY + holes(index).Y) / 2#, _
                    holes(index).Y - bounds.MinY, True, viewScale, created, expected
    Next i
    NdAddLinear view, lines(bounds.BottomIndex).Entity, lines(bounds.TopIndex).Entity, False, _
                bounds.MaxX + DIM_BASE_OFFSET + yPositions.Count * DIM_LINE_STEP, (bounds.MinY + bounds.MaxY) / 2#, _
                bounds.MaxY - bounds.MinY, False, viewScale, created, expected
    For i = 1 To groups.Count
        Set group = groups(i)
        index = CLng(group(1))
        drawingModel.ClearSelection2 True
        NdSelectEdge view, holes(index).Entity, False
        Set display = drawingModel.AddDiameterDimension2(holes(index).X, bounds.MaxY + DIM_BASE_OFFSET + (i - 1) * DIM_LINE_STEP, 0#)
        NdRegister display, 2# * holes(index).Radius, created, expected
        If display.Type2 <> swDiameterDimension Then Fail "Создан не диаметральный размер."
        NdValidate display, 2# * holes(index).Radius, viewScale
        If group.Count > 1 Then
            prefix = CStr(group.Count) & " X "
            display.SetText swDimensionTextPrefix, prefix
            If StrComp(display.GetText(swDimensionTextPrefix), prefix, vbBinaryCompare) <> 0 Then Fail "Не установлено количество отверстий."
        End If
    Next i
    drawingModel.EditRebuild3
    NdValidateAll created, expected, viewScale
    If (bounds.MaxX - bounds.MinX) / viewScale > BREAK_MODEL_THRESHOLD Then
        unbrokenWidth = bounds.MaxX - bounds.MinX
        If NdTryBreakLongView(view, holes, holeCount, bounds, viewScale, created, expected) Then
            breakApplied = True
            breakOutline = ReadOutline(view)
            bounds.MinX = CDbl(breakOutline(0)): bounds.MaxX = CDbl(breakOutline(2))
            bounds.MinY = CDbl(breakOutline(1)): bounds.MaxY = CDbl(breakOutline(3))
        End If
    End If
    box = NdDimensionEnvelope(created, bounds)
    dx = targetRight - box(3)
    dy = targetTop - box(4)
    position = ReadPosition(view)
    position(0) = position(0) + dx
    position(1) = position(1) + dy
    MarkStage "VBA-размеры: окончательное размещение"
    view.Position = position
    drawingModel.EditRebuild3
    actual = ReadPosition(view)
    If Abs(actual(0) - position(0)) > 0.0001 Or Abs(actual(1) - position(1)) > 0.0001 Then Fail "Не применено положение вида с размерами."
    bounds.MinX = bounds.MinX + dx: bounds.MaxX = bounds.MaxX + dx
    bounds.MinY = bounds.MinY + dy: bounds.MaxY = bounds.MaxY + dy
    If view.ProjectedDimensions = False Then Fail "После размещения сброшен проекционный режим."
    NdValidateAll created, expected, viewScale
    box = NdDimensionEnvelope(created, bounds)
    If Abs(box(3) - targetRight) > 0.0001 Or Abs(box(4) - targetTop) > 0.0001 Then Fail "Не подтверждены границы размещения вида с размерами."
    drawingModel.ClearSelection2 True
    NativeDimensionView = box
    Exit Function
Failed:
    number = Err.Number
    description = Err.Description
    origin = Err.Source
    On Error GoTo 0
    If breakApplied Then
        cleanup = NdUndoBreak(view, unbrokenWidth)
        If breaksCompleted > 0 Then breaksCompleted = breaksCompleted - 1
        breakWarnings.Add "Вид [" & currentView & "]: разрыв отменён после отказа размеров/компоновки."
    End If
    cleanup = cleanup & NdRollback(created)
    cleanup = cleanup & NdRestore(view, originalPosition, modeKnown, projectedBefore)
    WriteLog "VBA-размеры: отказ=" & CStr(number) & "; " & description & cleanup
    If number = 0 Then number = vbObjectError + 1900
    Err.Raise number, origin, description & cleanup
End Function

Private Function NdFinite(ByVal value As Double) As Boolean
    NdFinite = (value = value And Abs(value) <= 1.79769313486231E+308)
End Function

Private Function NdFullCircle(ByVal uMin As Double, ByVal uMax As Double) As Boolean
    NdFullCircle = Abs(Abs(uMax - uMin) - 2# * ND_PI) <= 0.000001
End Function

Private Function NdMatchesValue(ByVal actual As Double, ByVal expectedSheet As Double, ByVal viewScale As Double) As Boolean
    NdMatchesValue = Abs(actual - expectedSheet / viewScale) <= ND_DIAMETER_TOL
End Function

Private Function NdPositionNear(ByVal first As Double, ByVal second As Double, ByVal viewScale As Double) As Boolean
    NdPositionNear = Abs(first - second) <= ND_MODEL_TOL * viewScale
End Function

Private Function NdDiameterNear(ByVal firstRadius As Double, ByVal secondRadius As Double, ByVal viewScale As Double) As Boolean
    NdDiameterNear = Abs(2# * firstRadius - 2# * secondRadius) <= ND_DIAMETER_TOL * viewScale
End Function

Private Function NdFaceOn(ByVal center As Variant, ByVal normalPoint As Variant) As Boolean
    Dim dx As Double, dy As Double, dz As Double, length As Double
    dx = normalPoint(0) - center(0)
    dy = normalPoint(1) - center(1)
    dz = normalPoint(2) - center(2)
    length = Sqr(dx * dx + dy * dy + dz * dz)
    If length <= 0# Then Exit Function
    NdFaceOn = Sqr(dx * dx + dy * dy) / length <= 0.000001
End Function

Private Function NdProject(ByVal utility As SldWorks.MathUtility, ByVal transform As SldWorks.MathTransform, ByVal values As Variant) As Variant
    Dim point As SldWorks.MathPoint, projected As SldWorks.MathPoint
    Dim xyz(2) As Double, result As Variant, i As Long
    If HasItems(values) = False Then Fail "Нет координат точки геометрии."
    If LBound(values) <> 0 Or UBound(values) < 2 Then Fail "Неполные координаты геометрии."
    For i = 0 To 2
        xyz(i) = CDbl(values(i))
        If NdFinite(xyz(i)) = False Then Fail "Неконечные координаты модели."
    Next i
    Set point = utility.CreatePoint(xyz)
    If point Is Nothing Then Fail "Не создана геометрическая точка."
    Set projected = point.MultiplyTransform(transform)
    If projected Is Nothing Then Fail "Не спроецирована геометрическая точка."
    result = projected.ArrayData
    If HasItems(result) = False Then Fail "Нет спроецированных координат."
    If LBound(result) <> 0 Or UBound(result) < 2 Then Fail "Неполная проекция точки."
    For i = 0 To 2
        If NdFinite(CDbl(result(i))) = False Then Fail "Неконечная проекция точки."
    Next i
    NdProject = result
End Function

Private Sub NdIncludePoint(ByRef bounds As NDBounds, ByRef initialized As Boolean, ByVal x As Double, ByVal y As Double)
    If NdFinite(x) = False Or NdFinite(y) = False Then Fail "Неконечный экстремум контура."
    If initialized = False Then
        bounds.MinX = x: bounds.MaxX = x: bounds.MinY = y: bounds.MaxY = y
        initialized = True
    Else
        If x < bounds.MinX Then bounds.MinX = x
        If x > bounds.MaxX Then bounds.MaxX = x
        If y < bounds.MinY Then bounds.MinY = y
        If y > bounds.MaxY Then bounds.MaxY = y
    End If
End Sub

Private Function NdAtan2(ByVal y As Double, ByVal x As Double) As Double
    If x > 0# Then
        NdAtan2 = Atn(y / x)
    ElseIf x < 0# Then
        NdAtan2 = Atn(y / x)
        If y >= 0# Then NdAtan2 = NdAtan2 + ND_PI Else NdAtan2 = NdAtan2 - ND_PI
    ElseIf y > 0# Then
        NdAtan2 = ND_PI / 2#
    ElseIf y < 0# Then
        NdAtan2 = -ND_PI / 2#
    Else
        NdAtan2 = 0#
    End If
End Function

Private Sub NdArcExtrema(ByVal center As Variant, ByVal atZero As Variant, ByVal atQuarter As Variant, _
                         ByVal uMin As Double, ByVal uMax As Double, ByRef bounds As NDBounds, ByRef initialized As Boolean)
    Dim cx As Double, cy As Double, sx As Double, sy As Double, root As Double, first As Double, u As Double
    Dim axis As Long, k As Long
    If NdFinite(uMin) = False Or NdFinite(uMax) = False Then Fail "Неконечный интервал кругового ребра."
    If uMax < uMin Or uMax - uMin > 2# * ND_PI + 0.000001 Then Fail "Неподдерживаемый интервал дуги."
    cx = atZero(0) - center(0): cy = atZero(1) - center(1)
    sx = atQuarter(0) - center(0): sy = atQuarter(1) - center(1)
    NdArcPoint center, cx, cy, sx, sy, uMin, bounds, initialized
    NdArcPoint center, cx, cy, sx, sy, uMax, bounds, initialized
    For axis = 0 To 1
        If axis = 0 Then root = NdAtan2(sx, cx) Else root = NdAtan2(sy, cy)
        first = -Int(-((uMin - root) / ND_PI))
        For k = 0 To 3
            u = root + (first + k) * ND_PI
            If u <= uMax Then NdArcPoint center, cx, cy, sx, sy, u, bounds, initialized
        Next k
    Next axis
End Sub

Private Sub NdArcPoint(ByVal center As Variant, ByVal cx As Double, ByVal cy As Double, ByVal sx As Double, ByVal sy As Double, _
                       ByVal u As Double, ByRef bounds As NDBounds, ByRef initialized As Boolean)
    Dim x As Double, y As Double
    x = center(0) + cx * Cos(u) + sx * Sin(u)
    y = center(1) + cy * Cos(u) + sy * Sin(u)
    NdIncludePoint bounds, initialized, x, y
End Sub

Private Function NdBreakPlan(ByRef holes() As NDHole, ByVal holeCount As Long, ByRef bounds As NDBounds) As Variant
    Dim starts() As Double, finishes() As Double, leftEdge As Double, rightEdge As Double
    Dim width As Double, required As Double, middleLimit As Double, startX As Double, finishX As Double
    Dim bestStart As Double, bestFinish As Double, cursor As Double, swapStart As Double, swapFinish As Double
    Dim i As Long, j As Long, occupied As Long, hasMiddleHole As Boolean
    width = bounds.MaxX - bounds.MinX
    required = width - BREAK_MAX_WIDTH + BREAK_GAP + 0.002
    leftEdge = bounds.MinX + BREAK_END_CLEARANCE
    rightEdge = bounds.MaxX - BREAK_END_CLEARANCE
    For i = 1 To holeCount
        If holes(i).X > bounds.MinX + width / 3# And holes(i).X < bounds.MaxX - width / 3# Then
            middleLimit = holes(i).X - holes(i).Radius - BREAK_HOLE_CLEARANCE
            If hasMiddleHole = False Or middleLimit < rightEdge Then rightEdge = middleLimit
            hasMiddleHole = True
        End If
    Next i
    If rightEdge <= leftEdge Then Exit Function
    If rightEdge - leftEdge < required Then Exit Function
    ReDim starts(1 To holeCount)
    ReDim finishes(1 To holeCount)
    For i = 1 To holeCount
        startX = holes(i).X - holes(i).Radius - BREAK_HOLE_CLEARANCE
        finishX = holes(i).X + holes(i).Radius + BREAK_HOLE_CLEARANCE
        If finishX > leftEdge And startX < rightEdge Then
            occupied = occupied + 1
            If startX < leftEdge Then startX = leftEdge
            If finishX > rightEdge Then finishX = rightEdge
            starts(occupied) = startX
            finishes(occupied) = finishX
        End If
    Next i
    For i = 2 To occupied
        swapStart = starts(i): swapFinish = finishes(i)
        j = i - 1
        Do While j >= 1
            If starts(j) <= swapStart Then Exit Do
            starts(j + 1) = starts(j): finishes(j + 1) = finishes(j)
            j = j - 1
        Loop
        starts(j + 1) = swapStart: finishes(j + 1) = swapFinish
    Next i
    cursor = leftEdge
    For i = 1 To occupied
        If starts(i) - cursor > bestFinish - bestStart Then
            bestStart = cursor: bestFinish = starts(i)
        End If
        If finishes(i) > cursor Then cursor = finishes(i)
    Next i
    If rightEdge - cursor > bestFinish - bestStart Then
        bestStart = cursor: bestFinish = rightEdge
    End If
    If bestFinish - bestStart < required Then Exit Function
    startX = bestStart + (bestFinish - bestStart - required) / 2#
    NdBreakPlan = Array(startX, startX + required)
End Function

Private Function NdBreakSheetX(ByVal projectedX As Double, ByRef bounds As NDBounds, ByVal outline As Variant) As Double
    Dim projectedWidth As Double, outlineWidth As Double
    projectedWidth = bounds.MaxX - bounds.MinX
    outlineWidth = CDbl(outline(2)) - CDbl(outline(0))
    If projectedWidth <= 0# Or outlineWidth <= 0# Then Fail "Не определена ширина вида для линий разрыва."
    NdBreakSheetX = CDbl(outline(0)) + (projectedX - bounds.MinX) * outlineWidth / projectedWidth
End Function

Public Sub TestBreakPlan()
    Dim holes(1 To 3) As NDHole, bounds As NDBounds, plan As Variant
    bounds.MinX = 0#: bounds.MaxX = 0.34
    holes(1).X = 0.006: holes(1).Radius = 0.0014
    holes(2).X = 0.024: holes(2).Radius = 0.0014
    plan = NdBreakPlan(holes, 2, bounds)
    If IsEmpty(plan) Then Fail "Тест разрыва: не найден пустой участок после левых отверстий."
    If CDbl(plan(0)) < holes(2).X + holes(2).Radius + BREAK_HOLE_CLEARANCE Then Fail "Тест разрыва: скрыто левое отверстие."
    holes(3).X = 0.22: holes(3).Radius = 0.002
    plan = NdBreakPlan(holes, 3, bounds)
    If IsEmpty(plan) Then Fail "Тест разрыва: не найден участок слева от среднего отверстия."
    If CDbl(plan(1)) > holes(3).X - holes(3).Radius - BREAK_HOLE_CLEARANCE Then Fail "Тест разрыва: разрыв пересёк среднее отверстие."
    holes(3).X = 0.17
    plan = NdBreakPlan(holes, 3, bounds)
    If IsEmpty(plan) = False Then Fail "Тест разрыва: выбран слишком короткий участок."
    MsgBox BUILD_ID & ": TestBreakPlan passed; это проверка расчёта, без создания разрыва в SOLIDWORKS.", vbInformation
End Sub

Private Function NdUndoBreak(ByVal view As SldWorks.View, ByVal originalWidth As Double) As String
    Dim outline As Variant, remaining As Variant, i As Long, line As SldWorks.BreakLine
    On Error GoTo Failed
    SelectDrawingView view
    drawing.UnBreakView
    remaining = view.GetBreakLines
    If HasItems(remaining) Then
        For i = UBound(remaining) To LBound(remaining) Step -1
            Set line = remaining(i)
            If Not line Is Nothing Then
                drawingModel.ClearSelection2 True
                If line.Select(False, Nothing) = False Then Fail "Не выбрана оставшаяся линия разрыва."
                If drawingModel.Extension.DeleteSelection2(swDelete_Absorbed) = False Then _
                    Fail "Не удалена оставшаяся линия разрыва."
            End If
        Next i
    End If
    drawingModel.EditRebuild3
    If HasItems(view.GetBreakLines) Then Fail "После отмены остались линии разрыва."
    outline = ReadOutline(view)
    If CDbl(outline(2)) - CDbl(outline(0)) < originalWidth - 0.0001 Then _
        Fail "После отмены разрыва исходная длина вида не восстановилась."
    NdUndoBreak = " Разрыв отменён."
    Exit Function
Failed:
    NdUndoBreak = " ОШИБКА ОТМЕНЫ РАЗРЫВА: " & Err.Description
End Function

Private Function NdTryBreakLongView(ByVal view As SldWorks.View, ByRef holes() As NDHole, ByVal holeCount As Long, _
                                    ByRef bounds As NDBounds, ByVal viewScale As Double, _
                                    ByVal created As Collection, ByVal expected As Collection) As Boolean
    Dim plan As Variant, breakLine As SldWorks.BreakLine, secondLine As SldWorks.BreakLine, outline As Variant
    Dim postLines() As NDLine, postHoles() As NDHole, postBounds As NDBounds
    Dim postLineCount As Long, postHoleCount As Long, oldGap As Double, inserted As Boolean, gapChanged As Boolean
    Dim failure As String, rollback As String, firstX As Double, secondX As Double, middleY As Double
    Dim breakLines As Variant
    plan = NdBreakPlan(holes, holeCount, bounds)
    If IsEmpty(plan) Then
        breakWarnings.Add "Вид [" & currentView & "]: нет свободного участка для длины до 180 мм без закрытия отверстий."
        Exit Function
    End If
    On Error GoTo Failed
    MarkStage "Разрыв длинного вида после нанесения размеров"
    If HasItems(view.GetBreakLines) Then Fail "В скопированном виде уже есть линии разрыва."
    outline = ReadOutline(view)
    firstX = NdBreakSheetX(CDbl(plan(0)), bounds, outline)
    secondX = NdBreakSheetX(CDbl(plan(1)), bounds, outline)
    middleY = (CDbl(outline(1)) + CDbl(outline(3))) / 2#
    If firstX <= CDbl(outline(0)) Or secondX >= CDbl(outline(2)) Or firstX >= secondX Then _
        Fail "Линии разрыва оказались вне контура вида."
    WriteLog "Разрыв: проекция X=" & CStr(CDbl(plan(0))) & "..." & CStr(CDbl(plan(1))) & _
             "; контур X=" & CStr(CDbl(outline(0))) & "..." & CStr(CDbl(outline(2))) & _
             "; линии на листе X=" & CStr(firstX) & "..." & CStr(secondX) & "; Y=" & CStr(middleY)
    oldGap = view.BreakLineGap
    view.BreakLineGap = BREAK_GAP
    gapChanged = True
    SelectDrawingView view
    Set breakLine = view.InsertBreak3(swBreakLineVertical, firstX, middleY, swBreakLine_Curve, 0, True)
    If breakLine Is Nothing Then Fail "SOLIDWORKS не создал первую линию разрыва."
    inserted = True
    Set secondLine = view.InsertBreak3(swBreakLineVertical, secondX, middleY, swBreakLine_Curve, 0, True)
    If secondLine Is Nothing Then Fail "SOLIDWORKS не создал вторую линию разрыва."
    breakLines = view.GetBreakLines
    If HasItems(breakLines) = False Then Fail "Линии разрыва не обнаружены в виде."
    If UBound(breakLines) - LBound(breakLines) + 1 < 2 Then Fail "Вид содержит менее двух линий разрыва."
    drawing.BreakView
    drawingModel.EditRebuild3
    outline = ReadOutline(view)
    If CDbl(outline(2)) - CDbl(outline(0)) > BREAK_MAX_WIDTH + 0.0001 Then Fail "Разрыв не сократил вид до 180 мм."
    NdReadGeometry view, viewScale, postLines, postLineCount, postHoles, postHoleCount, postBounds
    If postHoleCount <> holeCount Then Fail "После разрыва изменилось число видимых отверстий."
    NdValidateAll created, expected, viewScale
    breaksCompleted = breaksCompleted + 1
    WriteLog "Разрыв подтверждён: до=" & CStr((bounds.MaxX - bounds.MinX) * 1000#) & _
             " мм; после=" & CStr((CDbl(outline(2)) - CDbl(outline(0))) * 1000#) & " мм."
    NdTryBreakLongView = True
    Exit Function
Failed:
    failure = Err.Description
    If inserted = False Then
        On Error Resume Next
        breakLines = view.GetBreakLines
        inserted = HasItems(breakLines)
        On Error GoTo 0
    End If
    If inserted Then
        On Error GoTo 0
        rollback = NdUndoBreak(view, bounds.MaxX - bounds.MinX)
        If InStr(rollback, "ОШИБКА") > 0 Then Fail "Не удалось отменить ошибочный разрыв вида [" & currentView & "]: " & failure & rollback
        NdValidateAll created, expected, viewScale
    End If
    If gapChanged Then view.BreakLineGap = oldGap
    WriteLog "Разрыв пропущен: " & failure
    breakWarnings.Add "Вид [" & currentView & "]: разрыв пропущен; " & failure
End Function

Private Sub NdReadGeometry(ByVal view As SldWorks.View, ByVal viewScale As Double, ByRef lines() As NDLine, ByRef lineCount As Long, _
                           ByRef holes() As NDHole, ByRef holeCount As Long, ByRef bounds As NDBounds)
    Dim components As Variant, edges As Variant, component As SldWorks.Component2, edge As SldWorks.Edge
    Dim curve As SldWorks.Curve, vertexA As SldWorks.Vertex, vertexB As SldWorks.Vertex, parameters As Object
    Dim utility As SldWorks.MathUtility, transform As SldWorks.MathTransform
    Dim p As Variant, q As Variant, circleData As Variant, center As Variant, atZero As Variant, atQuarter As Variant
    Dim normalPoint As Variant, uMin As Double, uMax As Double, initialized As Boolean, i As Long
    lineCount = 0: holeCount = 0
    components = view.GetVisibleComponents
    If HasItems(components) = False Then Fail "Нет видимого компонента вида."
    If UBound(components) - LBound(components) <> 0 Then Fail "Для размеров нужен один видимый компонент."
    Set component = components(LBound(components))
    If component Is Nothing Then Fail "Пустой видимый компонент."
    edges = view.GetVisibleEntities2(component, swViewEntityType_Edge)
    If HasItems(edges) = False Then Fail "Нет видимых рёбер для размеров."
    ReDim lines(1 To UBound(edges) - LBound(edges) + 1)
    ReDim holes(1 To UBound(edges) - LBound(edges) + 1)
    Set utility = app.GetMathUtility
    If utility Is Nothing Then Fail "Недоступен MathUtility."
    Set transform = view.ModelToViewTransform
    If transform Is Nothing Then Fail "Нет преобразования вида."
    For i = LBound(edges) To UBound(edges)
        Set edge = edges(i)
        If edge Is Nothing Then Fail "Пустое ребро вида."
        Set curve = edge.GetCurve
        If curve Is Nothing Then Fail "Не прочитана кривая ребра."
        If curve.IsLine Then
            Set vertexA = edge.GetStartVertex
            Set vertexB = edge.GetEndVertex
            If vertexA Is Nothing Then Fail "Нет начала прямого ребра."
            If vertexB Is Nothing Then Fail "Нет конца прямого ребра."
            p = NdProject(utility, transform, vertexA.GetPoint)
            q = NdProject(utility, transform, vertexB.GetPoint)
            lineCount = lineCount + 1
            lines(lineCount).X1 = p(0): lines(lineCount).Y1 = p(1)
            lines(lineCount).X2 = q(0): lines(lineCount).Y2 = q(1)
            Set lines(lineCount).Entity = edge
            NdIncludePoint bounds, initialized, CDbl(p(0)), CDbl(p(1))
            NdIncludePoint bounds, initialized, CDbl(q(0)), CDbl(q(1))
        ElseIf curve.IsCircle Then
            Set parameters = edge.GetCurveParams3
            If parameters Is Nothing Then Fail "Нет параметров кругового ребра."
            circleData = curve.CircleParams
            If HasItems(circleData) = False Then Fail "Нет параметров окружности."
            If LBound(circleData) <> 0 Or UBound(circleData) < 6 Then Fail "Неполные параметры окружности."
            center = NdProject(utility, transform, Array(circleData(0), circleData(1), circleData(2)))
            atZero = NdProject(utility, transform, curve.Evaluate2(0#, 0))
            atQuarter = NdProject(utility, transform, curve.Evaluate2(ND_PI / 2#, 0))
            uMin = parameters.UMinValue: uMax = parameters.UMaxValue
            NdArcExtrema center, atZero, atQuarter, uMin, uMax, bounds, initialized
            If NdFullCircle(uMin, uMax) Then
                normalPoint = NdProject(utility, transform, Array(circleData(0) + circleData(3), circleData(1) + circleData(4), circleData(2) + circleData(5)))
                If NdFaceOn(center, normalPoint) Then
                    If NdInnerCylinder(edge) Then
                        holeCount = holeCount + 1
                        holes(holeCount).X = center(0): holes(holeCount).Y = center(1)
                        holes(holeCount).Radius = CDbl(circleData(6)) * viewScale
                        Set holes(holeCount).Entity = edge
                    End If
                End If
            End If
        Else
            Fail "Сплайн, эллипс или иная неподдерживаемая кривая: габарит для размеров неоднозначен."
        End If
    Next i
    If lineCount = 0 Or initialized = False Then Fail "Нет прямолинейного контура для размеров."
End Sub

Private Function NdInnerCylinder(ByVal edge As SldWorks.Edge) As Boolean
    Dim faces As Variant, face As SldWorks.Face2, surface As SldWorks.Surface, i As Long
    faces = edge.GetTwoAdjacentFaces2
    If HasItems(faces) = False Then Fail "Не прочитаны соседние грани окружности."
    For i = LBound(faces) To UBound(faces)
        Set face = faces(i)
        If Not face Is Nothing Then
            Set surface = face.GetSurface
            If Not surface Is Nothing Then
                If surface.IsCylinder Then
                    If face.FaceInSurfaceSense Then NdInnerCylinder = True
                End If
            End If
        End If
    Next i
End Function

Private Sub NdFindBounds(ByRef lines() As NDLine, ByVal count As Long, ByRef bounds As NDBounds, ByVal viewScale As Double)
    bounds.LeftIndex = NdBoundary(lines, count, True, bounds.MinX, viewScale)
    bounds.RightIndex = NdBoundary(lines, count, True, bounds.MaxX, viewScale)
    bounds.BottomIndex = NdBoundary(lines, count, False, bounds.MinY, viewScale)
    bounds.TopIndex = NdBoundary(lines, count, False, bounds.MaxY, viewScale)
    If bounds.MaxX - bounds.MinX <= ND_MODEL_TOL * viewScale Or bounds.MaxY - bounds.MinY <= ND_MODEL_TOL * viewScale Then Fail "Вырожденный контур вида."
    If bounds.LeftIndex = 0 Or bounds.RightIndex = 0 Or bounds.BottomIndex = 0 Or bounds.TopIndex = 0 Then
        Fail "Нужен ортогональный вид с четырьмя прямыми на действительных крайних границах. Косые торцы автоматически не размеряются."
    End If
End Sub

Private Function NdBetterBoundary(ByVal offset As Double, ByVal length As Double, ByVal bestOffset As Double, ByVal bestLength As Double) As Boolean
    NdBetterBoundary = offset < bestOffset Or (offset = bestOffset And length > bestLength)
End Function

Private Function NdBoundary(ByRef lines() As NDLine, ByVal count As Long, ByVal vertical As Boolean, ByVal extreme As Double, ByVal viewScale As Double) As Long
    Dim i As Long, length As Double, cross As Double, offset As Double, bestOffset As Double, bestLength As Double
    bestOffset = 1E+300
    For i = 1 To count
        length = Sqr((lines(i).X2 - lines(i).X1) ^ 2 + (lines(i).Y2 - lines(i).Y1) ^ 2)
        If vertical Then
            cross = Abs(lines(i).X1 - lines(i).X2)
            offset = Abs((lines(i).X1 + lines(i).X2) / 2# - extreme)
        Else
            cross = Abs(lines(i).Y1 - lines(i).Y2)
            offset = Abs((lines(i).Y1 + lines(i).Y2) / 2# - extreme)
        End If
        If length > ND_MODEL_TOL * viewScale And cross <= length * 0.000001 And offset <= ND_DIAMETER_TOL * viewScale Then
            If NdBetterBoundary(offset, length, bestOffset, bestLength) Then
                NdBoundary = i
                bestOffset = offset: bestLength = length
            End If
        End If
    Next i
End Function

Private Function NdHoleBefore(ByRef holes() As NDHole, ByVal a As Long, ByVal b As Long, ByVal mode As Long) As Boolean
    If mode = 1 Then
        If holes(a).Radius <> holes(b).Radius Then
            NdHoleBefore = holes(a).Radius < holes(b).Radius
            Exit Function
        End If
    ElseIf mode = 2 Then
        If holes(a).Y <> holes(b).Y Then
            NdHoleBefore = holes(a).Y < holes(b).Y
            Exit Function
        End If
    End If
    If holes(a).X <> holes(b).X Then
        NdHoleBefore = holes(a).X < holes(b).X
    ElseIf holes(a).Y <> holes(b).Y Then
        NdHoleBefore = holes(a).Y < holes(b).Y
    Else
        NdHoleBefore = holes(a).Radius < holes(b).Radius
    End If
End Function

Private Sub NdSortIndices(ByRef holes() As NDHole, ByRef order() As Long, ByVal count As Long, ByVal mode As Long)
    Dim i As Long, j As Long, value As Long
    For i = 2 To count
        value = order(i)
        j = i - 1
        Do While j >= 1
            If NdHoleBefore(holes, value, order(j), mode) = False Then Exit Do
            order(j + 1) = order(j)
            j = j - 1
        Loop
        order(j + 1) = value
    Next i
End Sub

Private Function NdPrepareHoles(ByRef holes() As NDHole, ByVal count As Long, ByRef bounds As NDBounds, ByVal viewScale As Double) As Collection
    Dim result As Collection, order() As Long, i As Long, j As Long, index As Long, other As Long, duplicate As Boolean, tolerance As Double
    Set result = New Collection
    tolerance = ND_MODEL_TOL * viewScale
    If count > 0 Then
        ReDim order(1 To count)
        For i = 1 To count: order(i) = i: Next i
        NdSortIndices holes, order, count, 0
        For i = 1 To count
            index = order(i)
            If NdFinite(holes(index).X) = False Or NdFinite(holes(index).Y) = False Or NdFinite(holes(index).Radius) = False Then Fail "Неконечная окружность."
            If holes(index).Radius <= 0# Then Fail "Неположительный радиус отверстия."
            If holes(index).X > bounds.MinX + tolerance And holes(index).X < bounds.MaxX - tolerance And _
               holes(index).Y > bounds.MinY + tolerance And holes(index).Y < bounds.MaxY - tolerance Then
                duplicate = False
                For j = 1 To result.Count
                    other = CLng(result(j))
                    If NdPositionNear(holes(index).X, holes(other).X, viewScale) And NdPositionNear(holes(index).Y, holes(other).Y, viewScale) Then
                        If NdDiameterNear(holes(index).Radius, holes(other).Radius, viewScale) = False Then Fail "Концентрические окружности разных диаметров: выбор отверстия неоднозначен."
                        duplicate = True
                        Exit For
                    End If
                Next j
                If duplicate = False Then result.Add index
            End If
        Next i
    End If
    Set NdPrepareHoles = result
End Function

Private Function NdPositions(ByRef holes() As NDHole, ByVal selected As Collection, ByVal horizontal As Boolean, ByVal viewScale As Double) As Collection
    Dim result As Collection, order() As Long, i As Long, j As Long, index As Long, other As Long, duplicate As Boolean
    Dim first As Double, second As Double
    Set result = New Collection
    If selected.Count > 0 Then
        ReDim order(1 To selected.Count)
        For i = 1 To selected.Count: order(i) = CLng(selected(i)): Next i
        If horizontal Then NdSortIndices holes, order, selected.Count, 0 Else NdSortIndices holes, order, selected.Count, 2
        For i = 1 To selected.Count
            index = order(i)
            If horizontal Then first = holes(index).X Else first = holes(index).Y
            duplicate = False
            For j = 1 To result.Count
                other = CLng(result(j))
                If horizontal Then second = holes(other).X Else second = holes(other).Y
                If NdPositionNear(first, second, viewScale) Then
                    duplicate = True
                    Exit For
                End If
            Next j
            If duplicate = False Then result.Add index
        Next i
    End If
    Set NdPositions = result
End Function

Private Function NdDiameterGroups(ByRef holes() As NDHole, ByVal selected As Collection, ByVal viewScale As Double) As Collection
    Dim result As Collection, group As Collection, order() As Long, i As Long, j As Long, index As Long, representative As Long, found As Boolean
    Set result = New Collection
    If selected.Count > 0 Then
        ReDim order(1 To selected.Count)
        For i = 1 To selected.Count: order(i) = CLng(selected(i)): Next i
        NdSortIndices holes, order, selected.Count, 1
        For i = 1 To selected.Count
            index = order(i)
            found = False
            For j = 1 To result.Count
                Set group = result(j)
                representative = CLng(group(1))
                If NdDiameterNear(holes(index).Radius, holes(representative).Radius, viewScale) Then
                    group.Add index
                    found = True
                    Exit For
                End If
            Next j
            If found = False Then
                Set group = New Collection
                group.Add index
                result.Add group
            End If
        Next i
    End If
    Set NdDiameterGroups = result
End Function

Private Sub NdSelectEdge(ByVal view As SldWorks.View, ByVal edge As SldWorks.Edge, ByVal append As Boolean)
    If edge Is Nothing Then Fail "Нет ребра привязки размера."
    If view.SelectEntity(edge, append) = False Then Fail "Не выбрано ребро в требуемом виде."
End Sub

Private Sub NdRegister(ByVal display As SldWorks.DisplayDimension, ByVal expectedValue As Double, ByVal created As Collection, ByVal expected As Collection)
    If display Is Nothing Then Fail "SolidWorks не создал размер."
    created.Add display
    expected.Add expectedValue
    drawingModel.ClearSelection2 True
End Sub

Private Sub NdAddLinear(ByVal view As SldWorks.View, ByVal datum As SldWorks.Edge, ByVal target As SldWorks.Edge, ByVal horizontal As Boolean, _
                        ByVal x As Double, ByVal y As Double, ByVal expectedValue As Double, ByVal toCircle As Boolean, ByVal viewScale As Double, _
                        ByVal created As Collection, ByVal expected As Collection)
    Dim display As SldWorks.DisplayDimension, dimension As SldWorks.Dimension, i As Long, condition As Long
    drawingModel.ClearSelection2 True
    NdSelectEdge view, datum, False
    NdSelectEdge view, target, True
    If horizontal Then Set display = drawingModel.AddHorizontalDimension2(x, y, 0#) Else Set display = drawingModel.AddVerticalDimension2(x, y, 0#)
    NdRegister display, expectedValue, created, expected
    If horizontal Then
        If display.Type2 <> swHorLinearDimension And display.Type2 <> swLinearDimension Then Fail "Создан размер другого типа вместо горизонтального."
    Else
        If display.Type2 <> swVertLinearDimension And display.Type2 <> swLinearDimension Then Fail "Создан размер другого типа вместо вертикального."
    End If
    Set dimension = display.GetDimension2(0)
    If dimension Is Nothing Then Fail "Не прочитаны данные размера."
    If toCircle Then
        For i = 0 To 1
            condition = dimension.GetArcEndCondition(i)
            If condition >= swArcEndConditionCenter And condition <= swArcEndConditionMax Then
                dimension.SetArcEndCondition i, swArcEndConditionCenter
                If dimension.GetArcEndCondition(i) <> swArcEndConditionCenter Then Fail "Не подтверждена привязка размера к центру окружности."
            End If
        Next i
    End If
    NdValidate display, expectedValue, viewScale
End Sub

Private Sub NdValidate(ByVal display As SldWorks.DisplayDimension, ByVal expectedValue As Double, ByVal viewScale As Double)
    Dim dimension As SldWorks.Dimension, actual As Double
    If display Is Nothing Then Fail "Пустой созданный размер."
    Set dimension = display.GetDimension2(0)
    If dimension Is Nothing Then Fail "Нет значения созданного размера."
    actual = dimension.SystemValue
    If NdFinite(actual) = False Then Fail "Неконечное значение созданного размера."
    WriteLog "VBA-размер: тип=" & CStr(display.Type2) & "; ожидалось мм=" & CStr(expectedValue / viewScale * 1000#) & "; получено мм=" & CStr(actual * 1000#)
    If NdMatchesValue(actual, expectedValue, viewScale) = False Then Fail "Значение размера не совпало с геометрией. Возможна привязка к касательной."
End Sub

Private Sub NdValidateAll(ByVal created As Collection, ByVal expected As Collection, ByVal viewScale As Double)
    Dim i As Long, display As SldWorks.DisplayDimension
    If created.Count <> expected.Count Then Fail "Нарушен учёт созданных размеров."
    For i = 1 To created.Count
        Set display = created(i)
        NdValidate display, CDbl(expected(i)), viewScale
    Next i
End Sub

Private Function NdDimensionEnvelope(ByVal created As Collection, ByRef bounds As NDBounds) As Variant
    Dim minX As Double, minY As Double, maxX As Double, maxY As Double, i As Long
    Dim display As SldWorks.DisplayDimension, annotation As SldWorks.Annotation, point As Variant
    minX = bounds.MinX: minY = bounds.MinY: maxX = bounds.MaxX: maxY = bounds.MaxY
    If created.Count = 0 Then Fail "Нет подтверждённых размеров."
    For i = 1 To created.Count
        Set display = created(i)
        Set annotation = display.GetAnnotation
        If annotation Is Nothing Then Fail "Нет аннотации созданного размера."
        point = annotation.GetPosition
        If HasItems(point) = False Then Fail "Нет позиции размерной надписи."
        If LBound(point) <> 0 Or UBound(point) < 1 Then Fail "Неполная позиция размерной надписи."
        If NdFinite(CDbl(point(0))) = False Or NdFinite(CDbl(point(1))) = False Then Fail "Неконечная позиция надписи."
        If point(0) < minX Then minX = point(0)
        If point(0) > maxX Then maxX = point(0)
        If point(1) < minY Then minY = point(1)
        If point(1) > maxY Then maxY = point(1)
    Next i
    ' Anchor positions plus allowance, not exact glyph bounds.
    NdDimensionEnvelope = Array(created.Count, minX - 0.015, minY - 0.015, maxX + 0.015, maxY + 0.015)
End Function

Private Function NdDeleteDimension(ByVal display As SldWorks.DisplayDimension) As Boolean
    Dim annotation As SldWorks.Annotation, manager As SldWorks.SelectionMgr
    On Error GoTo Failed
    If display Is Nothing Then Exit Function
    drawingModel.ClearSelection2 True
    Set annotation = display.GetAnnotation
    If annotation Is Nothing Then Exit Function
    If annotation.Select3(False, Nothing) = False Then Exit Function
    Set manager = drawingModel.SelectionManager
    If manager Is Nothing Then Exit Function
    If manager.GetSelectedObjectCount2(-1) <> 1 Then Exit Function
    NdDeleteDimension = drawingModel.Extension.DeleteSelection2(0)
    Exit Function
Failed:
    WriteLog "VBA-размеры: ошибка удаления нового размера: " & Err.Description
End Function

Private Function NdRollback(ByVal created As Collection) As String
    Dim i As Long, display As SldWorks.DisplayDimension, failed As Long
    If created Is Nothing Then Exit Function
    For i = created.Count To 1 Step -1
        Set display = created(i)
        If NdDeleteDimension(display) = False Then failed = failed + 1
    Next i
    If failed > 0 Then NdRollback = vbCrLf & "Не все новые размеры удалось удалить; проверьте вид вручную."
End Function

Private Function NdRestorePosition(ByVal view As SldWorks.View, ByVal position As Variant) As String
    Dim actual As Variant
    On Error GoTo Failed
    If view Is Nothing Then Exit Function
    If HasItems(position) = False Then Exit Function
    view.Position = position
    drawingModel.EditRebuild3
    actual = ReadPosition(view)
    If Abs(actual(0) - position(0)) > 0.0001 Or Abs(actual(1) - position(1)) > 0.0001 Then NdRestorePosition = " Положение не восстановлено."
    Exit Function
Failed:
    NdRestorePosition = " Ошибка восстановления положения: " & Err.Description
End Function

Private Function NdRestoreMode(ByVal view As SldWorks.View, ByVal modeKnown As Boolean, ByVal projectedBefore As Boolean) As String
    On Error GoTo Failed
    If view Is Nothing Then Exit Function
    If modeKnown = False Then Exit Function
    view.ProjectedDimensions = projectedBefore
    drawingModel.EditRebuild3
    If view.ProjectedDimensions <> projectedBefore Then NdRestoreMode = " Режим размеров не восстановлен."
    Exit Function
Failed:
    NdRestoreMode = " Ошибка восстановления режима: " & Err.Description
End Function

Private Function NdRestore(ByVal view As SldWorks.View, ByVal position As Variant, ByVal modeKnown As Boolean, ByVal projectedBefore As Boolean) As String
    ' Restoration attempts are independent: a position failure must not skip the mode.
    NdRestore = NdRestorePosition(view, position) & NdRestoreMode(view, modeKnown, projectedBefore)
End Function
