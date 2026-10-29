Attribute VB_Name = "InsertWeldWithHoles"
Option Explicit

Const PATH_WELD_TABLE As String = "\\192.168.1.14\SolidWorks\Библиотека Solid Works НОВАЯ\Шаблоны\Сварные конструкции.sldwldtbt"
'Const PATH_WELD_TABLE As String = "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\lang\english\cut list.sldwldtbt"
Dim swApp As SldWorks.SldWorks
Dim sNeedPaleteNames(3) As String
Dim swModel As SldWorks.ModelDoc2
Dim swDrawDoc As SldWorks.DrawingDoc
Dim lStatus As Long
Dim status As Boolean
Dim fso As New FileSystemObject

Private Function initial()
''Populate need palete names
    sNeedPaleteNames(0) = "*Спереди"
    sNeedPaleteNames(1) = "*Слева"
    sNeedPaleteNames(2) = "*Сверху"
'    sNeedPaleteNames(3) = "*Back"
'    sNeedPaleteNames(4) = "*Right"
'    sNeedPaleteNames(5) = "*Bottom"
End Function

Sub start()
'On Error GoTo catch_
Call initial
Set swApp = Application.SldWorks

If swMessageBoxResult_e.swMbHitYes <> swApp.SendMsgToUser2("Запустить вставку деталей", swMessageBoxIcon_e.swMbQuestion, swMessageBoxBtn_e.swMbYesNo) Then
    Exit Sub
End If

Dim t As Single
t = Timer

Dim swView As SldWorks.View
Set swModel = swApp.ActiveDoc

enableBoost swModel

Set swDrawDoc = swModel
'Dim path$
'path = "C:\Users\RudyMan\Desktop\frame\frame.sldprt"
'status = swDraw.GenerateViewPaletteViews(path)
'Set swView = swDraw.GetFirstView
'Set swView = swDraw.DropDrawingViewFromPalette2("*isometric", 0.05, 0.05, 0)
Dim swSheetView As SldWorks.View
Set swSheetView = swDrawDoc.GetFirstView
Set swView = swSheetView.GetNextView

Dim refDoc As ModelDoc2
Set refDoc = swView.ReferencedDocument
If refDoc.GetType() <> swDocumentTypes_e.swDocPART Then
     Err.Raise vbError, "Sub start()", "Вид должен содержать деталь"
End If


InsertWeldTable swView

'Dim vOutline As Variant
'vOutline = swView.GetOutline
'Debug.Print "Xmin: " + CStr(vOutline(0)) + "Ymin: " + CStr(vOutline(1)) + "Xmax: " + CStr(vOutline(2)) + "Ymax: " + CStr(vOutline(3))
'Dim heightView As Double
'Dim widthView As Double
'heightView = vOutline(3) - vOutline(1)
'widthView = vOutline(2) - vOutline(0)

'Dim heightSheet As Double
'Dim widthSheet As Double
'Dim vOutlineSheet As Variant
'vOutlineSheet = swSheetView.GetOutline
'heightSheet = vOutlineSheet(3) - vOutlineSheet(1)
'widthSheet = vOutlineSheet(2) - vOutlineSheet(0)
''Debug.Print CStr(heightSheet) + " " + CStr(widthSheet)

'Dim vPos As Variant
'vPos = swView.Position
'Debug.Print CStr(vPos(0)) + " " + CStr(vPos(1))
status = setDisplayViewMain(swView)

Dim vNeedBodies As Variant
vNeedBodies = getUniqBodiesWithCircleHoles(swView)

'Dim count As Integer
'count = UBound(vNeedBodies)
copyAndPastView2 swView, swModel, vNeedBodies

GoTo finaly_
Exit Sub
catch_:
    lStatus = swApp.SendMsgToUser2(Err.Description + vbCrLf + Err.Source + vbCrLf, swMessageBoxIcon_e.swMbStop, swMessageBoxBtn_e.swMbOk)
finaly_:
2
disableBoost swModel

Dim min As Integer
Dim sec As Integer
sec = (Timer - t) Mod 60 * 60
min = Timer - t - sec

lStatus = swApp.SendMsgToUser2("Время выполнения операции: " + CStr(min) + "мин. " + CStr(sec) + "сек.", swMessageBoxIcon_e.swMbInformation, swMessageBoxBtn_e.swMbOk)

End Sub

Private Function InsertWeldTable(swView As SldWorks.View)
    If fso.FileExists(PATH_WELD_TABLE) = False Then
        Err.Raise vbError, "Function InsertWeldTable(swView As SldWorks.View)", "Неверный путь к файлу шаблона таблицы списка вырезов или файл не существует"
    End If
    Dim weldTable As SldWorks.WeldmentCutListAnnotation
    
    Dim model As ModelDoc2
    Dim swConf As Configuration
    Set model = swView.ReferencedDocument
    Set swConf = model.GetActiveConfiguration
'    Set weldTable = swView.InsertWeldmentTable(False, 0.5, 0.5, 1, swConf.Name, PATH_WELD_TABLE)
'    swApp.SendMsgToUser swConf.Name
    
    Set weldTable = swView.InsertWeldmentTable(False, 0.5, 0.5, 1, "По умолчанию<Как сварной>", PATH_WELD_TABLE)
    If weldTable Is Nothing Then
        Set weldTable = swView.InsertWeldmentTable(False, 0.5, 0.5, 1, "По умолчанию", PATH_WELD_TABLE)
        If weldTable Is Nothing Then
            Err.Raise vbError, "Function InsertWeldTable(swView As SldWorks.View)", "Не удалось создать таблицу списка вырезов"
        End If
    End If
End Function

'TODO
Private Function insertBreakLine(swView As SldWorks.View, swModel As SldWorks.ModelDoc2)
    Dim swDraw As SldWorks.DrawingDoc
    Set swDraw = swModel
    Dim vPos As Variant
    vPos = swView.Position
    Dim start As Double
    Dim finish As Double
    start = vPos(0)
    finish = vPos(1)
    
    Dim swBreakLIne As SldWorks.BreakLine
    Set swBreakLIne = swView.InsertBreak3(swBreakLineOrientation_e.swBreakLineVertical, 0, 0, swBreakLineStyle_e.swBreakLine_Curve, 0, True)

End Function

Private Function setDisplayViewMain(swView As SldWorks.View) As Boolean
    setDisplayViewMain = swView.SetDisplayMode4(False, swDisplayMode_e.swSHADED, False, True, True)
End Function

Private Function setDisplayViewPart(swView As SldWorks.View) As Boolean
    setDisplayViewPart = swView.SetDisplayMode4(False, swDisplayMode_e.swHIDDEN_GREYED, False, False, True)
    swView.SetDisplayTangentEdges2 (swDisplayTangentEdges_e.swTangentEdgesHidden)
End Function

Private Function getBodiesWithExtrudes(swView As SldWorks.View) As Variant
    Dim resBodies() As Body2
    Dim resBodiesCount As Integer
    Dim refModel As SldWorks.ModelDoc2
    Set refModel = swView.ReferencedDocument
    Dim swFeat As SldWorks.Feature
    Set swFeat = refModel.FirstFeature
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2() = "ICE" Then
            Dim swBody As SldWorks.Body2
            Dim swFace As SldWorks.face2
            Set swFace = swFeat.GetFaces()(0)
            Set swBody = swFace.GetBody()
            resBodiesCount = resBodiesCount + 1
            ReDim Preserve resBodies(resBodiesCount)
            Set resBodies(resBodiesCount - 1) = swBody
        End If
        Set swFeat = swFeat.GetNextFeature()
    Loop
    getBodiesWithExtrudes = resBodies
End Function

Private Function isIncludeCircleHole(swBody As SldWorks.Body) As Boolean
    Dim swFace As SldWorks.face2
    Set swFace = swBody.GetFirstFace
    Do While Not swFace Is Nothing
        Dim swLoop As SldWorks.Loop2
        Set swLoop = swFace.GetFirstLoop
        Do While Not swLoop Is Nothing
            If swLoop.GetEdgeCount() = 1 Then
                Dim swEdge As SldWorks.Edge
                Set swEdge = swLoop.GetEdges()(0)
                Dim swCurveParamData As SldWorks.CurveParamData
                Set swCurveParamData = swEdge.GetCurveParams3()
                If swCurveParamData.CurveType = swCurveTypes_e.CIRCLE_TYPE Then
                    isIncludeCircleHole = True
                    Exit Function
                End If
            End If
            Set swLoop = swLoop.GetNext()
        Loop
        Set swFace = swFace.GetNextFace()
    Loop
    isIncludeCircleHole = False
End Function

Private Function getVisibleFaceWithHoles(swView As SldWorks.View) As SldWorks.face2
    Dim swBody As SldWorks.Body2
    Set swBody = swView.Bodies(0)
    Dim swFace As SldWorks.face2
    Set swFace = swBody.GetFirstFace()
    Do While Not swFace Is Nothing
        Dim swLoop As SldWorks.Loop2
        Set swLoop = swFace.GetFirstLoop()
        Do While Not swLoop Is Nothing
            If swLoop.GetEdgeCount = 1 Then
                Dim swEdge As SldWorks.Edge
                Set swEdge = swLoop.GetEdges()(0)
                    Dim swCurveParamData As SldWorks.CurveParamData
                    Set swCurveParamData = swEdge.GetCurveParams3
                    If swCurveParamData.CurveType = swCurveTypes_e.CIRCLE_TYPE Then
                        Set getVisibleFaceWithHoles = swFace
                        Exit Function
                    End If
            End If
            Set swLoop = swLoop.GetNext()
        Loop
        Set swFace = swFace.GetNextFace()
    Loop
    Set getVisibleFaceWithHoles = Nothing
End Function

Private Function getUniqBodiesWithCircleHoles(swView As SldWorks.View) As Variant
    Dim refModel As SldWorks.ModelDoc2
    Set refModel = swView.ReferencedDocument
    Dim resBodies() As SldWorks.Body2
    Dim sizeBodies As Integer
    sizeBodies = 0
    Dim swFeat As SldWorks.Feature
    Set swFeat = refModel.FirstFeature()
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2() = "SolidBodyFolder" Then
            Dim swSubFeat As SldWorks.Feature
            Set swSubFeat = swFeat.GetFirstSubFeature()
            Do While Not swSubFeat Is Nothing
                Dim folder As SldWorks.BodyFolder
                Set folder = swSubFeat.GetSpecificFeature2()
                Dim swBody As SldWorks.Body2
'                Debug.Print swSubFeat.Name + " " + swSubFeat.GetTypeName2()
'                Debug.Print folder.GetBodyCount()
                If folder.GetBodyCount() > 0 Then
                    Set swBody = folder.GetBodies()(0)
                    If isIncludeCircleHole(swBody) Then
                        sizeBodies = sizeBodies + 1
                        ReDim Preserve resBodies(sizeBodies)
                        Set resBodies(sizeBodies - 1) = swBody
                    End If
                End If
                Set swSubFeat = swSubFeat.GetNextSubFeature()
            Loop
        End If
        Set swFeat = swFeat.GetNextFeature()
    Loop
    getUniqBodiesWithCircleHoles = resBodies
End Function

Private Function copyAndPastView2(swView As SldWorks.View, model As SldWorks.ModelDoc2, vBodies As Variant)
    model.ClearSelection2 True
    Dim centerPoint As Variant
    centerPoint = swView.Position
    Dim vSizes As Variant
    vSizes = getSizeView(swView)
    Dim Xmin As Double
    Dim Ymin As Double
    Xmin = vSizes(0)
    Ymin = vSizes(1)
    Dim status As Boolean
    status = swModel.Extension.SelectByID2(swView.GetName2, "DRAWINGVIEW", Xmin, Ymin, 0, True, 0, Nothing, swSelectOption_e.swSelectOptionDefault)
    swModel.EditCopy
    Dim i As Integer
    Dim prevView As SldWorks.View
    Dim pastedView As SldWorks.View
    Dim swDrawDoc As SldWorks.DrawingDoc
    Set swDrawDoc = swModel
    Dim x As Double
    
    For i = 0 To UBound(vBodies) - 1
        Dim visibleBodies(1) As SldWorks.Body2
        Set visibleBodies(0) = vBodies(i)
        Dim j As Integer
        For j = 0 To UBound(sNeedPaleteNames) - 1
            Ymin = Ymin - 0.1
            Set pastedView = pastView(Xmin, Ymin)
            nextShowNameView swModel, pastedView, sNeedPaleteNames(j)
            pastedView.Bodies = visibleBodies
            
            If isExistHoleVisibleFace2(pastedView) = False Then
                swModel.ClearSelection2 True
                status = swModel.Extension.SelectByID2(pastedView.GetName2, "DRAWINGVIEW", Xmin, Ymin, 0, True, 0, Nothing, swSelectOption_e.swSelectOptionDefault)
                swModel.EditDelete
                Ymin = Ymin + 0.1
            Else

                'Set scale 1/5
                pastedView.ScaleDecimal = 0.2
                'Set display
                setDisplayViewPart pastedView
                
                Dim swFace As SldWorks.face2
                Set swFace = getVisibleFaceWithHoles(pastedView)
                'Insert Ballon with position cutlistitem
                InsertBaloonNote swFace
                
'                SetDisplayDims pastedView
                
                'Set orientation drawing view to horizontal
                setOrientationToHorizontal pastedView
                
'                status = swModel.Extension.SelectByID2(pastedView.GetName2, "DRAWINGVIEW", Xmin, Ymin, 0, True, 0, Nothing, swSelectOption_e.swSelectOptionDefault)
'                swModel.EditCopy
'                Dim posPastedView As Variant
'                posPastedView = pastedView.Position
'                Set pastedView = pastView(x + 0.7, Ymin)
'
'                status = swModel.Extension.SelectByID2(pastedView.GetName2, "DRAWINGVIEW", Xmin, Ymin, 0, True, 0, Nothing, swSelectOption_e.swSelectOptionDefault)
'                status = pastedView.ReplaceViewWithSketch()
                
                
'                Dim swComp As Component2
'                Set swComp = pastedView.GetVisibleComponents()(0)
'                Dim swEnts As Variant
'                swEnts = pastedView.GetVisibleEntities2(swComp, swViewEntityType_e.swViewEntityType_Vertex)
'                Dim vVerts As Collection
'                Set vVerts = getNeedVertexes(swEnts, pastedView)
'
'                Debug.Print vVerts.count
                
            End If
        Next j
        
        Dim newSizes As Variant
        newSizes = getSizeView(pastedView)
        
        If i = 0 Then
            Dim sizes As Variant
            sizes = pastedView.Position
            x = sizes(0)
            swModel.ClearSelection2 True

            Dim p(2) As Double
            p(0) = x
            p(1) = -0.2

            pastedView.Position = p
        End If
    Next i
    swModel.ClearSelection2 True
    
End Function

Private Function setOrientationToHorizontal(swView As SldWorks.View)
'    Debug.Print swView.GetName2
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    Dim vPos As Variant
    vPos = swView.Position
    Dim status As Boolean
    status = swModel.Extension.SelectByID2(swView.GetName2, "DRAWINGVIEW", vPos(0), vPos(1), 0, False, 0, Nothing, 0)
    Dim swDraw As SldWorks.DrawingDoc
    Set swDraw = swModel
    Dim vSizes As Variant
    vSizes = getSizeView(swView)
    If vSizes(0) > vSizes(1) Then
        status = swDraw.DrawingViewRotate(1.5707963267949)
    End If

End Function

Private Function nextShowNameView(swModel As SldWorks.ModelDoc2, swView As SldWorks.View, nameShowView As String)
    Dim vPos As Variant
    Dim status As Boolean
    vPos = swView.Position
    status = swModel.Extension.SelectByID2(swView.GetName2, "DRAWINGVIEW", vPos(0), vPos(1), 0, False, 0, Nothing, 0)
    swModel.ShowNamedView2 nameShowView, -1
    swModel.EditRebuild3
    swModel.ClearSelection2 True
End Function

Private Function pastView(x As Double, y As Double) As SldWorks.View
    Dim status As Boolean
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    swModel.ClearSelection2 True
    Dim swSheet As SldWorks.sheet
    Dim swDrawDoc As SldWorks.DrawingDoc
    Set swDrawDoc = swModel
    Set swSheet = swDrawDoc.GetCurrentSheet
    status = swModel.Extension.SelectByID2(swSheet.GetName, "SHEET", x, y, 0, False, 0, Nothing, 0)
    swModel.Paste
    swModel.EditRebuild3
    swModel.ClearSelection2 True
    Set pastView = getLastViewOnPage(swDrawDoc)
End Function

Private Function getLastViewOnPage(swDrawDoc As SldWorks.DrawingDoc) As SldWorks.View
    Dim swSheet As SldWorks.sheet
    Dim swViewPage As SldWorks.View
    Set swViewPage = swDrawDoc.GetFirstView
    Dim status As Boolean
    status = swDrawDoc.ActivateSheet(swViewPage.GetName2())
    Set swSheet = swDrawDoc.GetCurrentSheet
    Dim vViews As Variant
    vViews = swSheet.GetViews
    Dim swCurView As SldWorks.View
    Dim swLastView As SldWorks.View
    Set swLastView = Nothing
    Dim i As Integer
    For i = 0 To UBound(vViews)
        Set swCurView = vViews(i)
        If swCurView.Type = swDrawingViewTypes_e.swDrawingNamedView Then
            If include(swCurView.GetName2, "Чертежный вид") Then
                Set swLastView = swCurView
            End If
        End If
    Next i
    Set getLastViewOnPage = swLastView
End Function

Private Function getNamedViewsOnPage(swDrawDoc As SldWorks.DrawingDoc) As Collection
    Dim swSheet As SldWorks.sheet
    Dim swViewPage As SldWorks.View
    Set swViewPage = swDrawDoc.GetFirstView
    Dim status As Boolean
    status = swDrawDoc.ActivateSheet(swViewPage.GetName2())
    Set swSheet = swDrawDoc.GetCurrentSheet
    Dim vViews As Variant
    vViews = swSheet.GetViews
    Dim swCurView As SldWorks.View
    Dim swLastView As SldWorks.View
    Set swLastView = Nothing
    Dim i As Integer
    Dim nameViewCol As New Collection
    For i = 0 To UBound(vViews)
        Set swCurView = vViews(i)
        If swCurView.Type = swDrawingViewTypes_e.swDrawingNamedView Then
            If include(swCurView.GetName2, "Чертежный вид") Then
                Set swLastView = swCurView
                nameViewCol.Add swCurView
            End If
        End If
    Next i
    Set getNamedViewsOnPage = nameViewCol
End Function

Private Function getPrevView(swDrawDoc As SldWorks.DrawingDoc) As SldWorks.View
    Dim col As New Collection
    Set col = getNamedViewsOnPage(swDrawDoc)
    Set getPrevView = col.Item(col.count - 1)
End Function

Private Function getSizeView(swView As SldWorks.View) As Variant
    Dim res(2) As Double
    Dim vSizes As Variant
    vSizes = swView.GetOutline
    'Height, Axis Y
    res(0) = vSizes(3) - vSizes(1)
    'Width, Axis X
    res(1) = vSizes(2) - vSizes(0)
    getSizeView = res
End Function

Private Function isExistHoleVisibleFace2(swView As SldWorks.View) As Boolean
'    Debug.Print swView.GetName2()
    Dim vEnts As Variant
    Dim swComp As SldWorks.Component2
    Dim swEnt As SldWorks.Entity
    Dim i As Integer
    Set swComp = swView.GetVisibleComponents()(0)
    vEnts = swView.GetVisibleEntities2(swComp, swViewEntityType_e.swViewEntityType_Face)
    For i = 0 To UBound(vEnts)
        Dim swFace As SldWorks.face2
        Set swFace = vEnts(i)
'        Debug.Print swFace.GetLoopCount()
'        If swFace.GetLoopCount > 1 Then
            Dim swLoop As SldWorks.Loop2
            Set swLoop = swFace.GetFirstLoop()
'            Debug.Print swLoop.GetEdgeCount()
            Do While Not swLoop Is Nothing
                If swLoop.GetEdgeCount = 1 Then
                    Dim swEdge As SldWorks.Edge
                    Set swEdge = swLoop.GetEdges()(0)
                        Dim swCurveParamData As SldWorks.CurveParamData
                        Set swCurveParamData = swEdge.GetCurveParams3
                        If swCurveParamData.CurveType = swCurveTypes_e.CIRCLE_TYPE Then
                            isExistHoleVisibleFace2 = True
                            Exit Function
                        End If
                End If
                Set swLoop = swLoop.GetNext()
            Loop
'        End If
    Next i
    isExistHoleVisibleFace2 = False
End Function

Private Function InsertBaloonNote(swFace As SldWorks.face2)
    Dim status As Boolean
    Dim lStatus As Long
    Dim swEnt As SldWorks.Entity
    Set swEnt = swFace
    Dim swSelData As SldWorks.SelectData
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = swModel.SelectionManager
    Set swSelData = swSelMgr.CreateSelectData()
    status = swEnt.Select4(False, swSelData)
    
    Dim myNote As SldWorks.Note
    Dim BomBalloonParams As SldWorks.BalloonOptions
    Set BomBalloonParams = swModel.Extension.CreateBalloonOptions()
    BomBalloonParams.Style = swBalloonStyle_e.swBS_None
    BomBalloonParams.Size = 2
    BomBalloonParams.UpperTextContent = 1
    BomBalloonParams.UpperText = """"
    BomBalloonParams.ShowQuantity = True
    BomBalloonParams.QuantityPlacement = 3
    BomBalloonParams.QuantityDenotationText = "шт"
    BomBalloonParams.QuantityOverride = False
    Set myNote = swModel.Extension.InsertBOMBalloon2(BomBalloonParams)
    
    Dim swAnnot As SldWorks.Annotation
    Set swAnnot = myNote.GetAnnotation
    lStatus = swAnnot.SetLeader3(swLeaderStyle_e.swNO_LEADER, swLeaderSide_e.swLS_LEFT, False, True, False, False)
    'boolstatus = Part.ActivateSheet("Лист1")
    'boolstatus = Part.ActivateView("Drawing View1")
    'boolstatus = Part.Extension.SelectByID2("DetailItem10@Drawing View1", "NOTE", 0.377200225054452, 0.145680006421856, 0, False, 0, Nothing, 0)
    'Set myNote = swModel.Extension.EditBalloonProperties2(swBalloonStyle_e.swBS_None, swBalloonFit_e.swBF_2Chars, swBalloonTextContent_e.swBalloonTextItemNumber, "1", 0, "", 0.01016, True, 3, "шт", 0.001)

End Function

Private Function SetDisplayDims(swView As SldWorks.View)
    Dim swSelMgr As SelectionMgr
    Set swSelMgr = swModel.SelectionManager
    Dim swComp As SldWorks.Component2
    Set swComp = swView.GetVisibleComponents()(0)
    Dim vEnts As Variant
    vEnts = swView.GetVisibleEntities2(swComp, swViewEntityType_e.swViewEntityType_Vertex)
    
    Dim swVert As SldWorks.Vertex
    Dim i As Integer
    Dim swSelData As SldWorks.SelectData
    
    Dim swSelDataOrigin As SelectData
    Set swSelDataOrigin = swSelMgr.CreateSelectData()
    swSelDataOrigin.Mark = swAutodimMark_e.swAutodimMarkOriginDatum
    
    Set swSelData = swSelMgr.CreateSelectData()
    Dim status As Boolean
    
    swModel.ClearSelection2 True
    
    'Debug.Print swView.GetOrientationName()
    'Debug.Print "Xmax: " + CStr(Xmax) + " Ymax: " + CStr(Ymax) + " Xmin: " + CStr(Xmin) + " Ymin: " + CStr(Ymin)
    Dim vFaces As Variant
    Dim swEnt As Entity
    vFaces = swView.GetVisibleEntities2(swComp, swViewEntityType_e.swViewEntityType_Face)
    Dim swFace As SldWorks.face2
    For i = 0 To UBound(vFaces)
        Set swFace = vFaces(i)
        If swFace.GetLoopCount() > 1 Then
            Dim swLoop As SldWorks.Loop2
            Set swLoop = swFace.GetFirstLoop()
            Do While Not swLoop Is Nothing
                If swLoop.IsOuter = False Then
                    If swLoop.GetEdgeCount() = 1 Then
                        Set swEnt = swLoop.GetEdges()(0)
                        status = swEnt.Select4(True, swSelData)
                    End If
                End If
                Set swLoop = swLoop.GetNext()
            Loop
        End If
    Next i
    
    Dim vVerts As Collection
    Set vVerts = getNeedVertexes(vEnts, swView)
'    Dim swEnt As Entity
    Set swEnt = vVerts.Item(1)
    status = swEnt.Select4(True, swSelData)
    Set swEnt = vVerts.Item(2)
    status = swEnt.Select4(True, swSelDataOrigin)
    
    Debug.Print swSelMgr.GetSelectedObjectCount2(-1)
    Dim lStatus As Long
    lStatus = swDrawDoc.AutoDimension(swAutodimEntities_e.swAutodimEntitiesSelected, 1, -1, 1, swAutodimVerticalPlacement_e.swAutodimVerticalPlacementRight)
    
    swModel.ClearSelection2 True
    
    Dim swDispDim As SldWorks.DisplayDimension
    Set swDispDim = swView.GetFirstDisplayDimension5()
    Do While Not swDispDim Is Nothing
        status = swModel.Extension.SelectByID2(swDispDim.GetNameForSelection(), "DIMENSION", 0, 0, 0, True, -1, Nothing, 0)
        Set swDispDim = swDispDim.GetNext5()
    Loop
    
    status = swModel.Extension.AlignDimensions(swAlignDimensionType_e.swAlignDimensionType_AutoArrange, 0.04)
End Function

Function getNeedVertexes(vEnts As Variant, swView) As Collection
    Dim resVertEnts As New Collection
    Dim i As Integer
    Dim swVert As Vertex
    Dim point As Variant
    Dim Xmax, Xmin, Ymax, Ymin, Zmax, Zmin As Double
    
    For i = 0 To UBound(vEnts)
        Set swVert = vEnts(i)
        Dim swEnt As SldWorks.Entity
        Set swEnt = swVert
        point = swVert.GetPoint()
    '    Debug.Print CStr(i + 1) + " x: " + CStr(point(0)) + " y: " + CStr(point(1)) + " z: " + CStr(point(2))
        Xmax = Helpers.MaxDouble(CDbl(Xmax), CDbl(point(0)))
        Ymax = Helpers.MaxDouble(CDbl(Ymax), CDbl(point(1)))
        Zmax = Helpers.MaxDouble(CDbl(Zmax), CDbl(point(2)))
        If i = 0 Then
            Xmin = point(0)
            Ymin = point(1)
            Zmin = point(2)
        Else
            Xmin = Helpers.MinDouble(CDbl(Xmin), CDbl(point(0)))
            Ymin = Helpers.MinDouble(CDbl(Ymin), CDbl(point(1)))
            Zmin = Helpers.MinDouble(CDbl(Zmin), CDbl(point(2)))
        End If
    '    status = swEnt.Select4(True, swSelData)
    Next i
    
    For i = 0 To UBound(vEnts)
        Set swVert = vEnts(i)
        point = swVert.GetPoint()
       
        Select Case swView.GetOrientationName()
        'YZ
        Case "*Справа"
            If point(1) = Ymin Then
                If point(2) = Zmin Then
                    Set swEnt = vEnts(i)
'                    status = swEnt.Select4(True, swSelData)
                    resVertEnts.Add swEnt
                End If
            End If
            If point(1) = Ymax Then
                If point(2) = Zmax Then
                    Set swEnt = vEnts(i)
'                    status = swEnt.Select4(True, swSelData)
                    resVertEnts.Add swEnt
                End If
            End If
        Case "*Слева"
            If point(1) = Ymin Then
                If point(2) = Zmin Then
                    Set swEnt = vEnts(i)
'                    status = swEnt.Select4(True, swSelData)
                    resVertEnts.Add swEnt
                End If
            End If
            If point(1) = Ymax Then
                If point(2) = Zmax Then
                    Set swEnt = vEnts(i)
'                    status = swEnt.Select4(True, swSelData)
                    resVertEnts.Add swEnt
                End If
            End If
        'XZ
        Case "*Сверху"
            If point(0) = Xmax Then
                If point(2) = Zmin Then
                    Set swEnt = vEnts(i)
'                    status = swEnt.Select4(True, swSelData)
                    resVertEnts.Add swEnt
                End If
            End If
            If point(0) = Xmin Then
                If point(2) = Zmax Then
                    Set swEnt = vEnts(i)
'                    status = swEnt.Select4(True, swSelData)
                    resVertEnts.Add swEnt
                End If
            End If
        Case "*Снизу"
            If point(0) = Xmin Then
                If point(2) = Zmin Then
                    Set swEnt = vEnts(i)
'                    status = swEnt.Select4(True, swSelData)
                    resVertEnts.Add swEnt
                End If
            End If
            If point(0) = Xmax Then
                If point(2) = Zmax Then
                    Set swEnt = vEnts(i)
'                    status = swEnt.Select4(True, swSelData)
                    resVertEnts.Add swEnt
                End If
            End If
        'XY
        Case "*Спереди"
            If point(0) = Xmin Then
                If point(1) = Ymin Then
                    Set swEnt = vEnts(i)
'                    status = swEnt.Select4(True, swSelData)
                    resVertEnts.Add swEnt
                End If
            End If
            If point(0) = Xmax Then
                If point(1) = Ymax Then
                    Set swEnt = vEnts(i)
'                    status = swEnt.Select4(True, swSelData)
                    resVertEnts.Add swEnt
                End If
            End If
        Case "*Сзади"
            If point(0) = Xmax Then
                If point(1) = Ymin Then
                    Set swEnt = vEnts(i)
'                    status = swEnt.Select4(True, swSelData)
                    resVertEnts.Add swEnt
                End If
            End If
            If point(0) = Xmin Then
                If point(1) = Ymax Then
                    Set swEnt = vEnts(i)
'                    status = swEnt.Select4(True, swSelData)
                    resVertEnts.Add swEnt
                End If
            End If
        End Select
        '    Debug.Print CStr(i + 1) + " x: " + CStr(point(0)) + " y: " + CStr(point(1))
    Next i
    Set getNeedVertexes = resVertEnts
    
End Function

Private Function enableBoost(swModel As ModelDoc2)
'    swModel.Lock
    Dim swModelView As ModelView
    Set swModelView = swModel.ActiveView
    swModelView.EnableGraphicsUpdate = False
    
    Dim swConfMgr As ConfigurationManager
    Set swConfMgr = swModel.ConfigurationManager
    swConfMgr.EnableConfigurationTree = False
    
'    Dim swFeatMgr As FeatureManager
'    Set swFeatMgr.EnableFeatureTree = False
End Function

Private Function disableBoost(swModel As ModelDoc2)
'    swModel.UnLock
    Dim swModelView As ModelView
    Set swModelView = swModel.ActiveView
    swModelView.EnableGraphicsUpdate = True
    
    Dim swConfMgr As ConfigurationManager
    Set swConfMgr = swModel.ConfigurationManager
    swConfMgr.EnableConfigurationTree = True
    
'    Dim swFeatMgr As FeatureManager
'    Set swFeatMgr.EnableFeatureTreeWindow = True
End Function
