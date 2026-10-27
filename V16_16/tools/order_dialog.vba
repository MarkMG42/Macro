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
