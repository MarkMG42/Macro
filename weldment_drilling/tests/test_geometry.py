"""Pure VBA-expression checks; NOT a VBA or SOLIDWORKS/COM runtime.
Native copy/paste, null COM returns, geometry visibility and cleanup require
TestCopyRoundTrip and the native scenarios in README_RU.md.
"""
from pathlib import Path
import math
import random
import re
import unittest
ROOT = Path(__file__).resolve().parents[1]
SOURCE = (ROOT/'InsertWeldWithHoles.txt').read_text(encoding='utf-8')

def body(name):
    return re.search(rf'Private (?:Function|Sub) {name}\b.*?\n(.*?)\nEnd (?:Function|Sub)', SOURCE, re.S)[1]

def expression(text, values):
    text = text.strip().replace('#','').replace('^','**')
    for old,new in [('And','and'),('Or','or'),('Not','not')]:
        text = re.sub(rf'\b{old}\b',new,text)
    text = re.sub(r'(?<![<>=])=(?!=)','==',text)
    for name,value in values.items():
        if isinstance(value,(list,tuple)):
            text = re.sub(rf'\b{name}\(([^()]*)\)',rf'{name}[\1]',text)
    if re.search(r'[^\w\s.,+*/()\[\]<>=-]',text):
        raise ValueError(text)
    return eval(text,{'__builtins__':{},'Sqr':math.sqrt,'Abs':abs,'Len':len,'Cos':math.cos,'Sin':math.sin,'Atn':math.atan,'Int':math.floor,
        'StrComp':lambda a,b,mode:(a.casefold()>b.casefold())-(a.casefold()<b.casefold()),'vbTextCompare':1},values)

def assignment(function,target,values):
    rhs = re.search(rf'^\s*{re.escape(target)} = (.+)$',body(function),re.M)[1]
    return expression(rhs,values)

def facing(normal,rotation):
    values = dict(normal=normal,rotation=rotation)
    values['normalLength'] = assignment('FaceFacing','normalLength',values)
    values['axisLength'] = assignment('FaceFacing','axisLength',values)
    if values['normalLength']<1e-7 or values['axisLength']<1e-7:
        return 0.
    return assignment('FaceFacing','FaceFacing',values)

def unit(v):
    length = math.sqrt(sum(x*x for x in v))
    return tuple(x/length for x in v)

def camera(axis):
    data = [0.]*16
    data[2],data[5],data[8] = axis
    data[12] = 1.
    return data

class CoreTests(unittest.TestCase):
    def test_six_standard_axes(self):
        for axis in [(1,0,0),(-1,0,0),(0,1,0),(0,-1,0),(0,0,1),(0,0,-1)]:
            self.assertAlmostEqual(facing(axis,camera(axis)),1.)
            self.assertAlmostEqual(facing(tuple(-x for x in axis),camera(axis)),-1.)

    def test_normal_facing_at_200_angles_and_scales(self):
        rng = random.Random(42)
        for _ in range(200):
            axis = unit(tuple(rng.uniform(-1,1) for _ in range(3)))
            ref = (1,0,0) if abs(axis[0])<.9 else (0,1,0)
            tangent = unit((axis[1]*ref[2]-axis[2]*ref[1],axis[2]*ref[0]-axis[0]*ref[2],axis[0]*ref[1]-axis[1]*ref[0]))
            angle,scale = rng.uniform(0,math.pi),rng.uniform(.001,10.)
            normal = tuple((math.cos(angle)*a+math.sin(angle)*b)*scale for a,b in zip(axis,tangent))
            self.assertAlmostEqual(facing(normal,camera(axis)),math.cos(angle),places=10)

    def test_zero_vectors_not_accepted(self):
        self.assertEqual(facing((0,0,0),camera((0,0,1))),0.)
        self.assertEqual(facing((0,0,1),camera((0,0,0))),0.)

    def test_candidate_prefers_holes_then_area(self):
        for holes,area,best,best_area,expected in [(0,100,0,0,False),(1,1,0,0,True),(2,1,1,100,True),(1,100,2,1,False),(2,10,2,1,True),(2,1,2,10,False),(2,1,2,1,False)]:
            self.assertEqual(assignment('BetterCandidate','BetterCandidate',dict(holes=holes,area=area,bestHoles=best,bestArea=best_area)),expected)

    def test_body_mapping_rejects_wrong_ambiguous_empty_names(self):
        for expected,actual,matches,accepted in [('Pipe-1','Pipe-1',1,True),('Pipe-1','pipe-1',1,True),('Pipe-1','Pipe-2',1,False),('Pipe-1','Pipe-1',0,False),('Pipe-1','Pipe-1',2,False),('','',1,False),('Pipe-1','',1,False)]:
            self.assertEqual(assignment('BodyNameMatches','BodyNameMatches',dict(expectedName=expected,actualName=actual,matches=matches)),accepted)

    def test_paste_requires_one_accounted_view(self):
        for added,unknown,expected in [(0,False,False),(1,False,True),(2,False,False),(1,True,False),(0,True,False)]:
            self.assertEqual(assignment('CanUsePastedView','CanUsePastedView',dict(added=added,unknownPaste=unknown)),expected)

    def test_partial_failure_must_not_leave_orphan_views(self):
        for unknown,remaining,completed,expected in [(False,0,0,True),(False,3,3,True),(False,4,3,False),(True,3,3,False),(True,0,0,False),(False,2,3,False)]:
            self.assertEqual(assignment('CanContinueAfterFailure','CanContinueAfterFailure',dict(unknownPaste=unknown,remaining=remaining,completed=completed)),expected)

    def test_first_left_and_other_views_right_use_outline_not_origin(self):
        outline,position = (-.06,-.025,.14,.075),[.2,.3]
        code=body('PlaceDetail')
        moves=re.findall(r'^\s*position\(0\) = (.+)$',code,re.M)
        self.assertEqual(len(moves),2)
        env = dict(position=position,outline=outline,targetX=.04)
        left=expression(moves[0],env)
        self.assertAlmostEqual(outline[0]+left-position[0],.04)
        env['targetX']=.4
        right=expression(moves[1],env)
        self.assertAlmostEqual(outline[2]+right-position[0],.4)
        self.assertIn('If placeLeft Then',code)
        self.assertIn('targetX = SIDE_MARGIN + BALLOON_LEFT_OFFSET',code)
        self.assertIn('If placeLeft = False Then cursorY',body('DimensionAndArrange'))

    def test_no_model_window_or_cross_context_dependency(self):
        code = '\n'.join(line.split("'",1)[0] for line in SOURCE.splitlines())
        for forbidden in ['partModel.ActiveView','.NameView','CreateDrawViewFromModelView3','MathVector','GetCorrespondingEntity']:
            self.assertNotIn(forbidden,code)

    def test_hole_group_left_right_symmetric_and_tolerance(self):
        # Screenshot: 500 mm bar, holes at 30/120; scale 1:5.
        for centres,expected in [([30,120],False),([380,470],True),([30,470],False),([250],False),([250.2],False),([251],True)]:
            mean=sum(centres)/len(centres)*.0002
            self.assertEqual(assignment('HolesNeedFlip','HolesNeedFlip',dict(meanX=mean,left=0.,right=.1,tolerance=.0001)),expected)

    def test_hole_side_invariant_under_translation_and_scale(self):
        for scale in [.05,.2,1.,2.]:
            for shift in [-2.,0.,3.]:
                for fraction,expected in [(.15,False),(.5,False),(.85,True)]:
                    self.assertEqual(assignment('HolesNeedFlip','HolesNeedFlip',dict(meanX=shift+fraction*.5*scale,left=shift,right=shift+.5*scale,tolerance=.0001)),expected)

    def test_half_turn_reflects_group_and_preserves_spacing(self):
        # Geometry of an in-plane half turn: x'=left+right-x.
        left,right=.02,.12
        original=[left+.076,left+.094]
        reflected=[left+right-x for x in original]
        self.assertAlmostEqual(abs(reflected[1]-reflected[0]),abs(original[1]-original[0]))
        self.assertFalse(assignment('HolesNeedFlip','HolesNeedFlip',dict(meanX=sum(reflected)/2,left=left,right=right,tolerance=.0001)))

    def test_loop_centres_use_point_transform_once_without_extra_scale(self):
        code=body('HoleMeanX')
        self.assertIn('If IsRoundOpening(holeLoop) Then',code)
        self.assertEqual(code.count('count = count + 1'),1)
        self.assertIn('projected = point.MultiplyTransform(transform)',code)
        self.assertIn('transform = view.ModelToViewTransform',code)
        self.assertNotIn('* view.ScaleDecimal',code)
        self.assertIn('If projected Is Nothing Then',code)
        place=body('PlaceDetail')
        self.assertLess(place.index('DrawingViewRotate'),place.index('OrientHolesLeft'))
        self.assertLess(place.index('OrientHolesLeft'),place.index('view.Position = position'))

    def test_snapshot_accepts_repeated_names_without_using_them_as_identity(self):
        code=body('ViewSnapshot')
        self.assertIn('Set result = New Collection',code)
        self.assertIn('If ViewIndex(result, view) = 0 Then',code)
        self.assertIn('result.Add view',code)
        self.assertNotIn('Dictionary',code)
        self.assertNotIn('.Exists',code)
        identity=body('SameDrawingView')
        self.assertIn('If first Is second Then',identity)
        self.assertIn('equality = app.IsSame(first, second)',identity)
        self.assertNotIn('GetName2',identity)
        self.assertNotIn('GetUniqueName',identity)
        self.assertRegex(identity,r'Case swObjectUnsupported[\s\S]*?SameDrawingView = False')

    def test_existing_objects_preserved_before_delta_and_name_free_selection(self):
        code=body('CopySourceView')
        intact=code.index('If OriginalViewsPresent(before, after) = False Then')
        register=code.index('madeViews.Add candidate')
        self.assertLess(intact,register)
        self.assertIn('pendingPaste = True',code[intact:register])
        self.assertIn('Fail ',code[intact:register])
        self.assertIn('If ViewIndex(before, candidate) = 0 Then',code)
        select=body('SelectDrawingView')
        self.assertIn('Set objects(0) = view',select)
        self.assertIn('MultiSelect2(objects, False, Nothing)',select)
        self.assertIn('SameDrawingView(selected, view)',select)
        self.assertNotIn('SelectByID2',select)
        self.assertNotIn('ViewKey',SOURCE)

    def test_identity_preflight_reads_twice_before_paste(self):
        code=body('CheckViewIdentity')
        self.assertEqual(code.count('= ViewSnapshot()'),2)
        self.assertIn('first.Count <> second.Count',code)
        self.assertIn('OriginalViewsPresent(first, second)',code)
        self.assertIn('OriginalViewsPresent(second, first)',code)
        self.assertIn('ViewIndex(first, source)',code)
        self.assertNotIn('Paste',code)
        self.assertNotIn('DeleteSelection2',code)

    def test_delete_proves_survivors_without_dereferencing_deleted_target(self):
        code=body('TryDeleteView')
        self.assertLess(code.index('targetIndex = ViewIndex(before, view)'),code.index('DeleteSelection2'))
        deleted=code[code.index('DeleteSelection2'):]
        self.assertIn('remaining.Count <> before.Count - 1',deleted)
        self.assertIn('If i <> targetIndex Then',deleted)
        self.assertIn('ViewIndex(remaining, survivor)',deleted)
        self.assertNotIn('ViewIndex(remaining, view)',deleted)
        self.assertNotIn('view.Get',deleted)

    def test_dimension_envelope_reflow_has_left_and_right_anchors(self):
        box=[-.04,-.07,.18,.09]
        env=dict(box=box,sheetWidth=.42,SIDE_MARGIN=.02,cursorY=.277)
        code=body('PlaceDimensionBox')
        moves=re.findall(r'^\s*dx = (.+)$',code,re.M)
        self.assertEqual(len(moves),2)
        left=expression(moves[0],dict(env,BALLOON_LEFT_OFFSET=.02))
        self.assertAlmostEqual(box[0]+left,.04)
        dx=expression(moves[1],env)
        dy=assignment('PlaceDimensionBox','dy',env)
        self.assertAlmostEqual(box[2]+dx,.4)
        self.assertAlmostEqual(box[3]+dy,.277)
        moved=[box[0]+dx,box[1]+dy,box[2]+dx,box[3]+dy]
        next_y=expression(re.search(r'If placeLeft = False Then cursorY = (.+)',code)[1],dict(moved=moved,GAP=.015))
        self.assertAlmostEqual((box[1]+dy)-next_y,.015)

    def test_dimension_pass_uses_native_vba_on_each_created_view(self):
        code=body('DimensionAndArrange')
        self.assertIn('For i = 1 To madeViews.Count',code)
        self.assertIn('placeLeft = (i = 1)',code)
        self.assertIn('If placeLeft = False Then cursorY = box(1) - GAP',code)
        self.assertIn('If placeLeft = False Then cursorY = moved(1) - GAP',body('PlaceDimensionBox'))
        attempt=body('TryDimensionView')
        self.assertLess(attempt.index('SelectDrawingView view'),attempt.index('NativeDimensionView(view, targetX, targetTop, placeLeft)'))
        self.assertIn('CDbl(result(0)) <= 0#',attempt)
        native=body('NativeDimensionView')
        self.assertIn('If placeLeft Then dx = targetX - box(1) Else dx = targetX - box(3)',native)
        self.assertIn('Abs(box(1) - targetX)',native)
        self.assertNotIn('CleanupSince',code+attempt)
        self.assertNotIn('TryDeleteView',code+attempt)
        self.assertLess(SOURCE.index('If ProcessBody(source, config, i, cursorY, width)'),SOURCE.index('DimensionAndArrange width, height'))

    def test_positions_follow_final_layout_and_use_linked_quantity(self):
        start=SOURCE.split('Public Sub start()',1)[1].split('\nEnd Sub',1)[0]
        self.assertLess(start.index('DimensionAndArrange width, height'),start.index('If ADD_BALLOONS And completed > 0 Then AddPositions'))
        self.assertIn('Private Const ADD_BALLOONS As Boolean = True',SOURCE)
        self.assertIn('madeBodyNames.Add currentBody',body('ProcessBody'))
        positions=body('AddPositions')
        self.assertIn('If madeTableReady = False Or madeTable Is Nothing Then',positions)
        self.assertIn('madeBodyNames.Count <> madeViews.Count',positions)
        self.assertIn('currentBody = CStr(madeBodyNames(i))',positions)
        balloon=body('TryAddPosition')
        self.assertIn('Set data.View = view',balloon)
        self.assertIn('options.UpperTextContent = swBalloonTextItemNumber',balloon)
        self.assertIn('options.ShowQuantity = True',balloon)
        self.assertIn('options.QuantityOverride = False',balloon)
        self.assertIn('actualOptions.ShowQuantity = False Or actualOptions.QuantityOverride',balloon)
        self.assertIn('x = CDbl(outline(0)) - BALLOON_LEFT_OFFSET',balloon)
        self.assertIn('y = (CDbl(outline(1)) + CDbl(outline(3))) / 2#',balloon)
        self.assertIn('DeleteSelection2(swDelete_Absorbed)',balloon)

    def test_no_external_dimension_runtime_or_process_launch(self):
        code='\n'.join(line.split("'",1)[0] for line in SOURCE.splitlines())
        for external in ['DIMENSIONS_PROGID','ProcessSelectedView','RegAsm','Shell','CreateProcess','GetActiveObject','RunMacro2','Declare ']:
            self.assertNotIn(external,code)
        created=re.findall(r'CreateObject\(([^)]+)\)',code)
        self.assertTrue(all(value in {'"Scripting.Dictionary"','"Scripting.FileSystemObject"'} for value in created))

    def test_native_dimension_failure_is_explicit_and_preserves_body_views(self):
        engine=body('NativeDimensionView')
        self.assertIn('If selectedHoles.Count = 0 Then Fail',engine)
        failed=engine[engine.index('Failed:'):]
        self.assertLess(failed.index('number = Err.Number'),failed.index('NdRollback'))
        self.assertLess(failed.index('NdRollback'),failed.index('NdRestore'))
        self.assertIn('Err.Raise number, origin, description & cleanup',failed)
        self.assertNotIn('TryDeleteView',failed)
        self.assertNotIn('CleanupSince',failed)
        self.assertIn('NdRestorePosition(view, position) & NdRestoreMode',body('NdRestore'))
        self.assertIn('GetSelectedObjectCount2(-1) <> 1',body('NdDeleteDimension'))

    def test_dimensions_registered_before_validation_and_rechecked_after_move(self):
        linear=body('NdAddLinear')
        self.assertLess(linear.index('NdRegister display'),linear.index('display.Type2'))
        self.assertIn('swLinearDimension',linear)
        self.assertIn('dimension.GetArcEndCondition(i) <> swArcEndConditionCenter',linear)
        engine=body('NativeDimensionView')
        self.assertEqual(engine.count('NdValidateAll created, expected, viewScale'),2)
        moved=engine.index('view.Position = position')
        self.assertIn('drawingModel.EditRebuild3',engine[moved:])
        self.assertIn('NdValidateAll created, expected, viewScale',engine[moved:])
        # Geometry and dimension-anchor boxes must be in the same translated coordinates.
        self.assertIn('bounds.MinX = bounds.MinX + dx',engine[moved:])
        self.assertIn('bounds.MinY = bounds.MinY + dy',engine[moved:])

    def test_native_full_circle_rejects_partial_arcs(self):
        for lo,hi,expected in [(0,2*math.pi,True),(1,1+2*math.pi,True),(0,math.pi,False),(0,2*math.pi-1e-5,False),(0,2*math.pi-5e-7,True)]:
            self.assertEqual(assignment('NdFullCircle','NdFullCircle',dict(uMin=lo,uMax=hi,ND_PI=math.pi)),expected)
        self.assertIn('uMax < uMin Or uMax - uMin > 2# * ND_PI',body('NdArcExtrema'))

    def test_native_size_values_are_model_metres_at_all_scales(self):
        for scale in [1.,.1,.01,2.]:
            for delta,expected in [(0,True),(0.5e-6,True),(2e-6,False),(-2e-6,False)]:
                self.assertEqual(assignment('NdMatchesValue','NdMatchesValue',dict(actual=.03+delta,expectedSheet=.03*scale,viewScale=scale,ND_DIAMETER_TOL=1e-6)),expected)

    def test_native_positions_and_diameters_use_distinct_model_tolerances(self):
        for scale in [1.,.1,.01,2.]:
            # 1 mm separate centres must remain distinct even at 1:100.
            self.assertFalse(assignment('NdPositionNear','NdPositionNear',dict(first=.01*scale,second=.011*scale,viewScale=scale,ND_MODEL_TOL=.00005)))
            self.assertTrue(assignment('NdPositionNear','NdPositionNear',dict(first=.01*scale,second=.01002*scale,viewScale=scale,ND_MODEL_TOL=.00005)))
            self.assertFalse(assignment('NdDiameterNear','NdDiameterNear',dict(firstRadius=.007*scale,secondRadius=.007002*scale,viewScale=scale,ND_DIAMETER_TOL=1e-6)))

    def test_native_boundary_prefers_extreme_then_length(self):
        for offset,length,bestOffset,bestLength,expected in [(0,.01,1e-8,1,True),(1e-8,1,0,.01,False),(0,2,0,1,True),(0,1,0,2,False),(0,1,0,1,False)]:
            self.assertEqual(assignment('NdBetterBoundary','NdBetterBoundary',dict(offset=offset,length=length,bestOffset=bestOffset,bestLength=bestLength)),expected)
        condition=re.search(r'If (length > .+?) Then',body('NdBoundary'))[1]
        for scale in [1.,.01,2.]:
            for cross,offset,expected in [(0,0,True),(.01*scale,0,False),(0,.00002*scale,False)]:
                self.assertEqual(expression(condition,dict(length=.1*scale,cross=cross,offset=offset,viewScale=scale,ND_MODEL_TOL=.00005,ND_DIAMETER_TOL=1e-6)),expected)

    def test_native_face_on_uses_projected_axis_not_global_z(self):
        for delta,expected in [((0,0,1),True),((0,0,-1),True),((1,0,0),False),((1e-8,0,1),True),((.01,0,1),False)]:
            env=dict(center=[2,-3,.5],normalPoint=[2+delta[0],-3+delta[1],.5+delta[2]])
            for var in ['dx','dy','dz','length']:
                env[var]=assignment('NdFaceOn',var,env)
            self.assertEqual(assignment('NdFaceOn','NdFaceOn',env),expected)

    def test_native_arc_points_keep_projected_radius_and_translation(self):
        for u,wantx,wanty in [(0,12,-3),(math.pi/2,10,-2),(math.pi/4,10+math.sqrt(2),-3+math.sqrt(.5))]:
            env=dict(center=[10,-3,0],cx=2,cy=0,sx=0,sy=1,u=u)
            self.assertAlmostEqual(assignment('NdArcPoint','x',env),wantx)
            self.assertAlmostEqual(assignment('NdArcPoint','y',env),wanty)

    def test_single_importable_module_contains_exact_native_fragments(self):
        for section,name in [('DECLARATIONS','native_dimensions_declarations.vba'),('ENGINE','native_dimensions.vba')]:
            fragment=(ROOT/'source'/name).read_text().rstrip()
            self.assertIn("' BEGIN NATIVE DIMENSIONS "+section+'\n'+fragment+"\n' END NATIVE DIMENSIONS "+section,SOURCE)
        self.assertLess(SOURCE.index('Private Type NDLine'),SOURCE.index('Public Sub start()'))

    def test_snapshot_logging_is_optional_without_changing_identity_checks(self):
        code=body('ViewSnapshot')
        self.assertIn('If VERBOSE_VIEW_SNAPSHOTS Then WriteLog',code)
        self.assertIn('If ViewIndex(result, view) = 0 Then',code)
        self.assertIn('Private Const VERBOSE_VIEW_SNAPSHOTS As Boolean = False',SOURCE)
        self.assertIn('OriginalViewsPresent(before, after)',body('CopySourceView'))

    def test_windows_import_encoding(self):
        self.assertEqual((ROOT/'InsertWeldWithHoles.bas').read_text(encoding='cp1251'),SOURCE)

    def test_layout_space_classifies_below_sheet_without_changing_scale(self):
        env=dict(SIDE_MARGIN=.02,BOTTOM_MARGIN=.04,sheetWidth=.42)
        for width,height,top,expected in [(.2,.05,.1,True),(.2,.05,.03,False),(.2,.05,-.1,False),(.5,.05,.2,False)]:
            self.assertEqual(assignment('FitsLayout','FitsLayout',dict(env,width=width,height=height,top=top)),expected)

    def test_final_warning_counts_actual_view_outline(self):
        env=dict(SIDE_MARGIN=.02,BOTTOM_MARGIN=.04,sheetWidth=.42,sheetHeight=.297)
        for outline,expected in [((.02,.05,.39,.27),True),((.02,-.1,.2,-.05),False),((.02,.04,.5,.1),False)]:
            self.assertEqual(assignment('OutlineInLayout','OutlineInLayout',dict(env,outline=outline)),expected)

    def test_break_runs_after_dimension_validation_and_before_final_layout(self):
        code=body('NativeDimensionView')
        self.assertLess(code.index('NdValidateAll created, expected, viewScale'),code.index('NdTryBreakLongView'))
        self.assertLess(code.index('NdTryBreakLongView'),code.index('box = NdDimensionEnvelope(created, bounds)'))
        self.assertIn('If postHoleCount <> holeCount Then Fail',body('NdTryBreakLongView'))
        self.assertIn('NdValidateAll created, expected, viewScale',body('NdTryBreakLongView'))
        self.assertIn('NdUndoBreak(view, unbrokenWidth)',code)

    def test_break_uses_two_vertical_lines_at_the_same_sheet_y(self):
        code=body('NdTryBreakLongView')
        self.assertIn('firstX = NdBreakSheetX(CDbl(plan(0)), bounds, outline)',code)
        self.assertIn('secondX = NdBreakSheetX(CDbl(plan(1)), bounds, outline)',code)
        self.assertIn('view.InsertBreak3(swBreakLineVertical, firstX, middleY,',code)
        self.assertIn('view.InsertBreak3(swBreakLineVertical, secondX, middleY,',code)
        self.assertLess(code.index('Set breakLine = view.InsertBreak3'),code.index('Set secondLine = view.InsertBreak3'))
        self.assertLess(code.index('Set secondLine = view.InsertBreak3'),code.index('drawing.BreakView'))
        self.assertIn('If UBound(breakLines) - LBound(breakLines) + 1 < 2 Then Fail',code)
        self.assertIn('If HasItems(view.GetBreakLines) Then Fail',body('NdUndoBreak'))

    def test_dimension_lines_are_six_millimetres_apart(self):
        self.assertIn('Private Const DIM_BASE_OFFSET As Double = 0.007',SOURCE)
        self.assertIn('Private Const DIM_LINE_STEP As Double = 0.006',SOURCE)
        code=body('NativeDimensionView')
        self.assertIn('bounds.MinY - DIM_BASE_OFFSET - (i - 1) * DIM_LINE_STEP',code)
        self.assertIn('bounds.MinY - DIM_BASE_OFFSET - xPositions.Count * DIM_LINE_STEP',code)
        self.assertNotIn('0.008',code)

    def test_cut_list_outside_sheet_and_ten_millimetre_view_gap(self):
        code=body('TryInsertOutsideCutList')
        self.assertIn('x = sheetWidth + SIDE_MARGIN',code)
        self.assertIn('source.InsertWeldmentTable(False, sheetWidth / 2#, sheetHeight / 2#, 1, CStr(candidate), tableTemplate)',code)
        self.assertIn('AddTableConfiguration candidates, seen, "По умолчанию<Как сварной>"',code)
        self.assertIn('If madeTable Is Nothing Then',code)
        self.assertIn('If table.RowCount < 2 Then Fail',code)
        start=SOURCE.split('Public Sub start()',1)[1].split('\nEnd Sub',1)[0]
        self.assertIn('For i = 1 To bodies.Count',start)
        self.assertLess(start.index('If TryInsertOutsideCutList'),start.index('For i = 1 To bodies.Count'))
        self.assertEqual(assignment('PlaceDetail','cursorY',dict(outline=[0,.12,.2,.2],GAP=.01)),.11)
        self.assertIn('Private Const GAP As Double = 0.01',SOURCE)

    def test_cut_list_template_cannot_be_confused_with_bom_template(self):
        configured=re.search(r'^Private Const TABLE_TEMPLATE As String = "(.*)"$',SOURCE,re.M)[1]
        self.assertTrue(configured.startswith('\\\\'))
        self.assertTrue(configured.lower().endswith('.sldwldtbt'))
        self.assertIn('Right$(filename, 10)',body('TableFileExists'))
        self.assertIn('".sldwldtbt"',body('TableFileExists'))
        self.assertIn('NormalizeTablePath(candidate)',body('ResolveTableTemplate'))

if __name__=='__main__':
    unittest.main()
