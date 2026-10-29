"""Read-only metadata audit for the pure VBA port. Developer dependency: dnfile."""
from pathlib import Path
import hashlib
import json
import sys
import dnfile
ROOT = Path(__file__).resolve().parents[1]
MEMBERS = {
    'ISldWorks': ['GetExecutablePath'],
    'IView': ['get_ProjectedDimensions', 'set_ProjectedDimensions', 'SelectEntity',
              'InsertWeldmentTable', 'InsertBreak3', 'GetBreakLines',
              'get_BreakLineGap', 'set_BreakLineGap'],
    'IDrawingDoc': ['BreakView', 'UnBreakView'],
    'ICurve': ['Evaluate2'],
    'IEdge': ['GetCurveParams3', 'GetTwoAdjacentFaces2'],
    'ICurveParamData': ['get_UMinValue', 'get_UMaxValue'],
    'IFace2': ['FaceInSurfaceSense'],
    'IAnnotation': ['GetPosition', 'SetPosition2', 'Select3'],
    'ITableAnnotation': ['GetAnnotation', 'get_RowCount', 'set_Anchored'],
    'IDisplayDimension': ['GetDimension2', 'GetAnnotation', 'get_Type2', 'SetText', 'GetText'],
    'IDimension': ['GetArcEndCondition', 'SetArcEndCondition', 'get_SystemValue'],
    'IModelDoc2': ['AddHorizontalDimension2', 'AddVerticalDimension2', 'AddDiameterDimension2'],
}
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
reports = []
for label, path in [('provided_interop', ROOT.parent/'autosizedrilling/SolidWorks.Interop.sldworks.dll'),
                    ('nuget_32_1', Path(sys.argv[1]))]:
    pe = dnfile.dnPE(str(path))
    types = {str(t.TypeName): t for t in pe.net.mdtables.TypeDef}
    members = []
    for interface, names in MEMBERS.items():
        methods = {str(m.row.Name): m.row for m in types[interface].MethodList}
        for name in names:
            row = methods[name]
            members.append(dict(interface=interface, member=name, signature_blob_hex=row.Signature.value.hex(),
                                parameter_names=[str(p.row.Name) for p in row.ParamList if p.row.Sequence > 0]))
    reports.append(dict(label=label, assembly_sha256=sha(path), members=members))
sources = [ROOT/'InsertWeldWithHoles.txt', ROOT/'source/native_dimensions.vba',
           ROOT/'source/native_dimensions_declarations.vba']
report = dict(build_id='DRILL_COPY_VIEW_R11_TABLE_PATH_20261029', runtime='SOLIDWORKS VBA only',
              source_sha256={p.relative_to(ROOT).as_posix(): sha(p) for p in sources},
              assemblies=reports, metadata_audited=True, native_vba_compile=False, native_solidworks_run=False,
              scope='Member names/signatures only; not VBA typechecking, COM execution or Defender scanning')
(ROOT/'checks/native_dimensions.json').write_text(json.dumps(report, indent=2)+'\n')
print('Verified',len(reports),'interop assemblies;',len(reports[0]['members']),'members each; no native VBA/COM execution')
