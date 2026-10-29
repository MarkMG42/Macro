"""Exercise scalar equations extracted from the shipped VBA, without simulating COM.

This is NOT a VBA interpreter or a SOLIDWORKS test. Native checks are in README_RU.
Only the small, whitelisted arithmetic-expression subset below is evaluated.
"""
from pathlib import Path
import math
import random
import re
import unittest

SOURCE = (Path(__file__).resolve().parents[1] / 'InsertWeldWithHoles.txt').read_text(encoding='utf-8')


def body(name):
    return re.search(rf'Private (?:Function|Sub) {name}\b.*?\n(.*?)\nEnd (?:Function|Sub)', SOURCE, re.S)[1]


def expression(text, values):
    text = text.strip().replace('#', '').replace('^', '**')
    for name in values:
        if isinstance(values[name], (list, tuple)):
            text = re.sub(rf'\b{name}\(([^()]*)\)', rf'{name}[\1]', text)
    # Restrict evaluation to equations maintained in this module.
    if re.search(r'[^\w\s.,+*/()\[\]<>=-]', text):
        raise ValueError(text)
    return eval(text, {'__builtins__': {}, 'Array': lambda *v: tuple(v),
                       'Sqr': math.sqrt, 'Abs': abs}, values)


def assignment(function, target, values):
    rhs = re.search(rf'^\s*{re.escape(target)} = (.+)$', body(function), re.M)[1]
    return expression(rhs, values)


def length(value):
    return assignment('VectorLength', 'VectorLength', {'value': value})


def unit(value):
    n = length(value)
    if n < 1e-7:
        raise ValueError('degenerate vector')
    return assignment('UnitVector', 'UnitVector', {'value': value, 'length': n})


def project(axis, normal):
    v = {'direction': axis, 'normal': normal}
    v['dot'] = assignment('ProjectDirection', 'dot', v)
    return assignment('ProjectDirection', 'ProjectDirection', v)


def cross(a, b):
    return assignment('CrossVector', 'CrossVector', {'a': a, 'b': b})


def basis(normal, axis):
    normal = unit(normal)
    x = project(axis, normal)
    if length(x) < 1e-7:
        axis = (0., 1., 0.) if abs(normal[0]) > .9 else (1., 0., 0.)
        x = project(axis, normal)
    x = unit(x)
    y = cross(normal, x)
    data = [0.] * 16
    # Read the matrix slot assignments from VBA, so transpositions regress this test.
    lines = re.findall(r'^\s*data\(([^)]+)\) = (.+)$', body('BasisData'), re.M)
    for k in range(3):
        for slot, rhs in lines:
            if 'k' in slot:
                values = dict(k=k, x=x, y=y, normal=normal)
                data[int(expression(slot, values))] = expression(rhs, values)
    data[12] = assignment('BasisData', 'data(12)', {})
    return data


def transform(v, data):
    # Independent SOLIDWORKS row-vector convention, not taken from the VBA assignments.
    return tuple(sum(v[k] * data[3*k+j] for k in range(3)) for j in range(3))


class GeometryTests(unittest.TestCase):
    def assertVector(self, actual, expected):
        for a, b in zip(actual, expected):
            self.assertAlmostEqual(a, b, places=10)

    def test_six_normals_face_viewer(self):
        for normal in [(1,0,0), (-1,0,0), (0,1,0), (0,-1,0), (0,0,1), (0,0,-1)]:
            self.assertVector(transform(normal, basis(normal, (1,0,0))), (0,0,1))

    def test_arbitrary_rotations_preserve_lengths_and_handedness(self):
        rng = random.Random(42)
        for _ in range(200):
            normal = unit(tuple(rng.uniform(-1,1) for _ in range(3)))
            axis = tuple(rng.uniform(-1,1) for _ in range(3))
            data = basis(normal, axis)
            self.assertVector(transform(normal, data), (0,0,1))
            self.assertAlmostEqual(length(transform(axis, data)), length(axis), places=10)
            self.assertVector(transform(unit(project(axis, normal)), data), (1,0,0))
            x, y, z = (tuple(data[j+3*k] for k in range(3)) for j in range(3))
            self.assertVector(cross(x,y), z)
            self.assertEqual(data[12], 1.)

    def test_parallel_axis_fallback(self):
        for n in [(1,0,0), (0,1,0), (0,0,-1), (1,2,3)]:
            self.assertVector(transform(unit(n), basis(n,n)), (0,0,1))

    def test_zero_direction_rejected(self):
        with self.assertRaises(ValueError):
            basis((0,0,0), (1,0,0))
        self.assertIn('If length < EPS Then Fail', body('UnitVector'))

    def test_layout_uses_outline_offsets_not_view_origin(self):
        outline = (-.06, -.025, .14, .075)
        position = [.2, .3]
        expected_left, expected_top = .02, .4
        env = dict(position=position, outline=outline, SIDE_MARGIN=expected_left, cursorY=expected_top)
        new_x = assignment('PlaceDetail', 'position(0)', env)
        new_y = assignment('PlaceDetail', 'position(1)', env)
        self.assertAlmostEqual(outline[0] + new_x - position[0], expected_left)
        self.assertAlmostEqual(outline[3] + new_y - position[1], expected_top)
        cursor = assignment('PlaceDetail', 'cursorY', dict(outline=outline, GAP=.015))
        self.assertAlmostEqual(cursor, -.04)

    def test_release_encodings(self):
        path = Path(__file__).resolve().parents[1] / 'InsertWeldWithHoles.bas'
        if not path.exists():
            self.skipTest('Run tools/package_release.py first')
        self.assertEqual(path.read_text(encoding='cp1251'), SOURCE)


if __name__ == '__main__':
    unittest.main()
