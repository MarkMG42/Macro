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
    ' Preserve incomplete work privately; never label it SUCCESS or export a final PDF.
    RequireOwnedPath model.GetPathName
    If Not SameText(NormalizePath(fsoRun.GetParentFolderName(model.GetPathName)), NormalizePath(runFolder)) Then Exit Sub
    model.ClearSelection2 True
    LogLine "FAILED CHECKPOINT begin: " & model.GetPathName
    saved = model.Save3(swSaveAsOptions_Silent, errors, warnings)
    LogLine "FAILED CHECKPOINT saved=" & CStr(saved) & "; errors=" & errors & "; warnings=" & warnings
    WriteStatus "FAILED", failure & vbCrLf & "Непроверенный рабочий файл: " & model.GetPathName & _
        "; сохранён=" & CStr(saved) & "; код=" & errors
End Sub
