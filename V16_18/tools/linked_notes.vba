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
                rec.Add "Attachments", NoteAttachmentSignature(annotation)
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

Private Function NoteAttachmentSignature(ByVal annotation As SldWorks.Annotation) As String
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
            Set component = entity.GetComponent
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
        rec("Link") = RemapComponentViews(CStr(rec("Link")), mapping)
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
        Require NoteAttachmentSignature(annotation) = CStr(item("Attachments")), _
            "Изменилась деталь привязки заметки: " & CStr(item("Name"))
        ' Reset the original layer as paste can inherit the destination's current layer.
        If annotation.Layer <> CStr(item("Layer")) Then annotation.Layer = CStr(item("Layer"))
        If CLng(annotation.Visible) <> CLng(item("Visible")) Then annotation.Visible = CLng(item("Visible"))
        ' Setting the live expression also asks SOLIDWORKS to reevaluate stale text.
        candidate.PropertyLinkedText = desired
        Require candidate.PropertyLinkedText = desired, "Не восстановлена ссылка заметки: " & CStr(item("Name"))
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
