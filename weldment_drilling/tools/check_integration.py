"""Audit/copy the compiled x64 bridge; requires dnfile, never executes COM."""
from pathlib import Path
import hashlib
import json
import shutil
import sys
import subprocess
import re
import dnfile

ROOT = Path(__file__).resolve().parents[1]
REPO = ROOT.parent
sdk = sys.argv[1]
artifacts = Path(sys.argv[2])
commands = []
def run(arguments):
    command = [sdk] + arguments
    result = subprocess.run(command, cwd=REPO, text=True, capture_output=True)
    commands.append(dict(command=command, exit_code=result.returncode))
    if result.returncode:
        print(result.stdout + result.stderr)
        raise SystemExit(result.returncode)
    return result.stdout

build_outputs = []
for project in ('weldment_drilling/integration/DrillingDimensionsBridge.csproj',
                'autosizedrilling/src/AutoSizeDrilling.csproj',
                'autosizedrilling/tests/GeometryTests.csproj'):
    build_outputs.append(run(['build', project, '-c', 'Release', '--artifacts-path', str(artifacts), '--nologo']))
test_output = run([str(artifacts / 'bin/GeometryTests/release/GeometryTests.dll')])
counts = re.search(r'(\d+) passed, (\d+) failed', test_output)
assert counts and int(counts[1]) > 0 and int(counts[2]) == 0, test_output
warning_counts = [re.search(r'(\d+) Warning\(s\)', output) for output in build_outputs]
error_counts = [re.search(r'(\d+) Error\(s\)', output) for output in build_outputs]
assert all(warning_counts) and all(error_counts)
warnings = sum(int(match[1]) for match in warning_counts)
errors = sum(int(match[1]) for match in error_counts)
assert errors == 0
build = artifacts / 'bin/DrillingDimensionsBridge/release'
assembly = build / 'Macro.DrillingDimensionsBridge.dll'
pe = dnfile.dnPE(str(assembly))
assert pe.FILE_HEADER.Machine == 0x8664, 'Bridge must target x64'
types = {str(t.TypeName): t for t in pe.net.mdtables.TypeDef}
assert 'SolidWorksMacro' not in types, 'Legacy VSTA CLSID must NOT be registered'
assert {'Bridge', 'IBridge', 'AutoSizeDrill'} <= types.keys()
attrs = {}
for attribute in pe.net.mdtables.CustomAttribute:
    constructor = attribute.Type.row
    owner = getattr(constructor, 'Class', None)
    if owner is None:
        continue
    attr_type = getattr(owner.row, 'TypeName', None)
    if attr_type is None:
        continue
    attrs[(id(attribute.Parent.row), str(attr_type))] = attribute.Value.value
assembly_row = pe.net.mdtables.Assembly.rows[0]
assert attrs[(id(assembly_row), 'ComVisibleAttribute')] == bytes.fromhex('0100000000')
for type_name in ('Bridge', 'IBridge'):
    assert attrs[(id(types[type_name]), 'ComVisibleAttribute')] == bytes.fromhex('0100010000')
assert b'Macro.DrillingDimensions' in attrs[(id(types['Bridge']), 'ProgIdAttribute')]
assert b'93769c05-7e2b-4350-bbbe-36a2a60f52d7' in attrs[(id(types['Bridge']), 'GuidAttribute')]
assert all(b'fcaaccd5-4cd6-4682-a2d6-5f5c3d34c3d8' not in value for value in attrs.values())
method = next(m.row for m in types['IBridge'].MethodList if str(m.row.Name) == 'ProcessSelectedView')
parameters = [str(p.row.Name) for p in method.ParamList if p.row.Sequence > 0]
assert parameters == ['application', 'expectedView', 'targetRight', 'targetTop']
release = ROOT / 'integration/release'
release.mkdir(exist_ok=True)
files = ['Macro.DrillingDimensionsBridge.dll', 'Macro.DrillingDimensionsBridge.pdb',
         'SolidWorks.Interop.sldworks.dll', 'SolidWorks.Interop.swconst.dll']
for name in files:
    shutil.copyfile(build / name, release / name)
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
for name in files[2:]:
    assert sha(release / name) == sha(REPO / 'autosizedrilling' / name)
sources = [ROOT / 'integration/Bridge.cs', ROOT / 'integration/Properties/AssemblyInfo.cs',
           ROOT / 'integration/DrillingDimensionsBridge.csproj']
sources += [REPO / 'autosizedrilling/src' / name for name in ('SolidWorksMacro.cs', 'Geometry.cs', 'DiagnosticLog.cs')]
report = dict(target='net48/x64', sdk=run(['--version']).strip(), build_warnings=warnings, build_errors=errors,
              metadata_audited=True, legacy_vsta_class_excluded=True,
              assembly_com_visible=False, bridge_parameters=parameters,
              bridge_signature_blob_hex=method.Signature.value.hex(),
              source_sha256={p.relative_to(REPO).as_posix(): sha(p) for p in sources},
              binary_sha256={p.relative_to(REPO).as_posix(): sha(p) for p in (release / n for n in files)},
              geometry_tests=dict(passed=int(counts[1]), failed=int(counts[2]), native_com=False),
              executed_commands=commands,
              standalone_build_passed=True,
              native_registration=False, native_com_activation=False, native_solidworks_run=False)
(ROOT / 'checks/integration.json').write_text(json.dumps(report, indent=2) + '\n')
print(json.dumps({k: v for k, v in report.items() if k not in ('source_sha256', 'binary_sha256')}, indent=2))
