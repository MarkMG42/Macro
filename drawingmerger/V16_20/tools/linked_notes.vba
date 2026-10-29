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
                desired = SectionComponentLink(view, desired, afterAttachments)
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

' Rebuild only a single-component SectionAssembly reference from the native tree.
' Never infer an instance from evaluated property text or from the first leader.
Private Function SectionComponentLink(ByVal view As SldWorks.View, ByVal raw As String, ByVal signature As String) As String
    Dim root As SldWorks.DrawingComponent, found As String, matches As Long
    Dim marker As Long, start As Long, closing As Long
    Require Left$(signature, 2) = "1:", "Служебная ссылка разреза имеет несколько привязок."
    marker = InStr(1, raw, "$COMP:" & Chr$(34), vbTextCompare)
    Require marker > 0, "Нет ссылки компонента разреза."
    start = marker + Len("$COMP:" & Chr$(34))
    closing = InStr(start, raw, Chr$(34))
    Require closing > 0, "Не закрыта ссылка компонента разреза."
    Require InStr(closing + 1, raw, "$COMP:", vbTextCompare) = 0, "Несколько ссылок компонентов в заметке разреза."
    Set root = view.RootDrawingComponent
    Require Not root Is Nothing, "Не прочитано дерево компонентов разреза."
    FindSectionComponent root, "", view.GetName2, Mid$(signature, 3), 0, found, matches
    Require matches = 1, "Неоднозначный/отсутствующий компонент в дереве разреза: " & view.GetName2
    SectionComponentLink = Left$(raw, start - 1) & found & Mid$(raw, closing)
End Function

Private Sub FindSectionComponent(ByVal node As SldWorks.DrawingComponent, ByVal prefix As String, _
    ByVal parentName As String, ByVal wanted As String, ByVal depth As Long, ByRef found As String, ByRef matches As Long)
    Dim component As SldWorks.Component2, children As Variant, child As SldWorks.DrawingComponent
    Dim name As String, path As String, modelName As String, expression As String, i As Long, dot As Long
    Require depth < 64, "Слишком глубокое дерево компонентов разреза."
    Set component = node.Component
    Require Not component Is Nothing, "Не прочитан компонент дерева разреза."
    name = component.Name2
    If InStrRev(name, "/") > 0 Then name = Mid$(name, InStrRev(name, "/") + 1)
    Require Len(name) > 0 And InStr(name, Chr$(34)) = 0 And InStr(name, "@") = 0, "Неподдерживаемое имя компонента разреза."
    Require InStr(parentName, Chr$(34)) = 0 And InStr(parentName, "@") = 0, "Неподдерживаемое имя родителя компонента."
    expression = prefix & name & "@" & parentName
    path = component.GetPathName
    If NormalizePath(path) & "|" & component.ReferencedConfiguration & "|" & component.Name2 = wanted Then
        matches = matches + 1
        found = expression
    End If
    children = node.GetChildren
    If IsArray(children) Then
        modelName = Mid$(path, InStrRev(path, "\") + 1)
        dot = InStrRev(modelName, ".")
        If dot > 0 Then modelName = Left$(modelName, dot - 1)
        Require Len(modelName) > 0, "Нет имени родительской модели разреза."
        For i = LBound(children) To UBound(children)
            Set child = children(i)
            FindSectionComponent child, expression & "/", modelName, wanted, depth + 1, found, matches
        Next i
    End If
End Sub
