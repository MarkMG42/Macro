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
