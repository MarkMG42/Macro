Attribute VB_Name = "AuditDimensionsModule"
Option Explicit

' Run after the merge stops. Open the source drawing and keep the failed result open.
' No Save, Paste or explicit Rebuild calls. Sheet activation may update SW caches.
Public Sub AuditDrawingDimensions()
    Dim app As Object, doc As Object, nextDoc As Object, initial As Object
    Dim drawing As Object, sheet As Object, view As Object, dd As Object, dimObj As Object
    Dim fso As Object, output As Object, names As Variant, name As Variant
    Dim folder As String, report As String, originalSheet As String
    Dim errors As Long, active As Object, guard As Long, enumerated As Long
    Dim itemName As String, itemValue As String, failure As String
    On Error GoTo Failed
    Set app = Application.SldWorks
    Set initial = app.ActiveDoc
    If initial Is Nothing Then Err.Raise vbObjectError + 1900, , "Откройте исходник и результат объединения."
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = fso.GetParentFolderName(initial.GetPathName)
    If Len(folder) = 0 Then Err.Raise vbObjectError + 1901, , "Активный чертёж должен иметь путь на диске."
    report = fso.BuildPath(folder, "DimensionsAudit_" & Format$(Now, "yyyymmdd_hhnnss") & ".tsv")
    Set output = fso.CreateTextFile(report, False, True)
    output.WriteLine "document" & vbTab & "sheet" & vbTab & "view" & vbTab & "record" & vbTab & "name_or_count" & vbTab & "value"
    Set doc = app.GetFirstDocument
    Do While Not doc Is Nothing
        Set nextDoc = doc.GetNext
        If doc.GetType = 3 Then
            Set active = app.ActivateDoc3(doc.GetTitle, False, swDontRebuildActiveDoc, errors)
            If active Is Nothing Then Err.Raise vbObjectError + 1902, , "Не активирован документ: " & doc.GetTitle
            If StrComp(active.GetPathName, doc.GetPathName, vbTextCompare) <> 0 Then Err.Raise vbObjectError + 1903, , "Активирован другой документ."
            Set drawing = doc
            Set sheet = drawing.GetCurrentSheet
            originalSheet = sheet.GetName
            names = drawing.GetSheetNames
            For Each name In names
                If drawing.ActivateSheet(CStr(name)) = False Then Err.Raise vbObjectError + 1904, , "Не активирован лист."
                Set view = drawing.GetFirstView
                If Not view Is Nothing Then Set view = view.GetNextView
                Do While Not view Is Nothing
                    output.WriteLine Cell(doc.GetPathName) & vbTab & Cell(CStr(name)) & vbTab & Cell(view.GetName2) & vbTab & "API_COUNT" & vbTab & view.GetDimensionCount4 & vbTab
                    enumerated = 0
                    guard = 0
                    Set dd = view.GetFirstDisplayDimension5
                    Do While Not dd Is Nothing
                        guard = guard + 1
                        If guard > 100000 Then Err.Raise vbObjectError + 1905, , "Зацикливание обхода размеров."
                        Set dimObj = dd.GetDimension2(0)
                        itemName = "<dimension unavailable>"
                        itemValue = ""
                        If Not dimObj Is Nothing Then
                            itemName = dimObj.FullName
                            itemValue = CStr(dimObj.SystemValue)
                        End If
                        output.WriteLine Cell(doc.GetPathName) & vbTab & Cell(CStr(name)) & vbTab & Cell(view.GetName2) & vbTab & "DIMENSION" & vbTab & Cell(itemName) & vbTab & Cell(itemValue)
                        enumerated = enumerated + 1
                        Set dd = dd.GetNext5
                    Loop
                    output.WriteLine Cell(doc.GetPathName) & vbTab & Cell(CStr(name)) & vbTab & Cell(view.GetName2) & vbTab & "ENUMERATED" & vbTab & enumerated & vbTab
                    Set view = view.GetNextView
                Loop
            Next name
            drawing.ActivateSheet originalSheet
            originalSheet = ""
        End If
        Set doc = nextDoc
    Loop
    output.Close
    Set output = Nothing
    Set active = app.ActivateDoc3(initial.GetTitle, False, swDontRebuildActiveDoc, errors)
    MsgBox "Отчёт создан:" & vbCrLf & report, vbInformation
    Exit Sub
Failed:
    failure = Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not output Is Nothing Then
        output.WriteLine "ERROR" & vbTab & Cell(failure)
        output.Close
    End If
    If Len(originalSheet) > 0 Then drawing.ActivateSheet originalSheet
    If Not initial Is Nothing Then Set active = app.ActivateDoc3(initial.GetTitle, False, swDontRebuildActiveDoc, errors)
    MsgBox "Диагностика остановлена: " & failure & vbCrLf & report, vbExclamation
End Sub

Private Function Cell(ByVal value As String) As String
    Cell = Replace(Replace(Replace(value, vbTab, " "), vbCr, " "), vbLf, " ")
End Function
