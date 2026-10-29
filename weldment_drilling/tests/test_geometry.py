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
    return eval(text,{'__builtins__':{},'Sqr':math.sqrt,'Abs':abs,'Len':len,
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

    def test_layout_uses_outline_not_origin(self):
        outline,position = (-.06,-.025,.14,.075),[.2,.3]
        env = dict(position=position,outline=outline,SIDE_MARGIN=.02,cursorY=.4)
        x,y = assignment('PlaceDetail','position(0)',env),assignment('PlaceDetail','position(1)',env)
        self.assertAlmostEqual(outline[0]+x-position[0],.02)
        self.assertAlmostEqual(outline[3]+y-position[1],.4)
        self.assertAlmostEqual(assignment('PlaceDetail','cursorY',dict(outline=outline,GAP=.015)),-.04)

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

if __name__=='__main__':
    unittest.main()
