"""Export Windows-1251 import module and deterministic review archive (stdlib only)."""
from pathlib import Path
import hashlib
import json
import unittest
import zipfile

ROOT = Path(__file__).resolve().parents[1]


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    source = ROOT / 'InsertWeldWithHoles.txt'
    code = source.read_text(encoding='utf-8')
    module = ROOT / 'InsertWeldWithHoles.bas'
    module.write_bytes(code.replace('\n', '\r\n').encode('cp1251'))
    assert module.read_text(encoding='cp1251') == code
    install_module = ROOT / 'install' / module.name
    install_module.parent.mkdir(exist_ok=True)
    install_module.write_bytes(module.read_bytes())
    assert install_module.read_bytes() == module.read_bytes()
    grammar = json.loads((ROOT / 'checks/grammar.json').read_text())
    assert not grammar['errors'], 'VBA grammar errors'
    assert grammar['source_sha256'] == sha(source), 'Rerun grammar validation for this source'
    suite = unittest.defaultTestLoader.discover(str(ROOT / 'tests'))
    result = unittest.TextTestRunner(verbosity=1).run(suite)
    tests = dict(tests=result.testsRun, failures=len(result.failures), errors=len(result.errors),
                 skipped=len(result.skipped), source_sha256=sha(source),
                 tests_sha256=sha(ROOT / 'tests/test_geometry.py'),
                 native_vba_execution=False, native_solidworks_execution=False,
                 scope='Pure scalar/policy expressions, encoding and architectural regression; no COM execution')
    (ROOT / 'checks/tests.json').write_text(json.dumps(tests, indent=2) + '\n')
    assert result.wasSuccessful() and result.testsRun > 0 and not result.skipped, 'Release tests failed/skipped or empty'
    # Source-only container rebuilding failed in the user's SOLIDWORKS host.
    # Do not accidentally redistribute an unsupported binary in a source release.
    assert not list(ROOT.rglob('*.swp')), 'Do not ship SWP without native load validation'
    report = dict(source_sha256=sha(source), import_module_sha256=sha(module),
                  original_module_sha256=sha(ROOT / 'original/InsertWeldWithHoles.bas'),
                  build_id='DRILL_COPY_VIEW_R16_FIRST_LEFT_20261029', workflow='Copy source drawing view, filter body, inspect six directions',
                  native_vba_compile=False, native_solidworks_run=False,
                  runtime='SOLIDWORKS VBA only', external_dll_required=False, administrator_required=False,
                  binary_swp_rebuilt=False, binary_swp_withdrawn=True,
                  native_solidworks_load=False, original_swp_sha256=
                  'ed7d9dc23db5bead52dc6208b314897fd7c96cf735f4ce82b0f945101c82363e')
    (ROOT / 'checks/release.json').write_text(json.dumps(report, indent=2) + '\n')
    for section, filename in (('DECLARATIONS', 'native_dimensions_declarations.vba'), ('ENGINE', 'native_dimensions.vba')):
        fragment = (ROOT / 'source' / filename).read_text(encoding='utf-8').rstrip()
        assert "' BEGIN NATIVE DIMENSIONS " + section + "\n" + fragment + "\n' END NATIVE DIMENSIONS " + section in code, 'Reassemble native fragments before grammar validation'
    assert 'DIMENSIONS_PROGID' not in code and 'ProcessSelectedView' not in code
    native = json.loads((ROOT / 'checks/native_dimensions.json').read_text())
    assert native['metadata_audited']
    for relative, expected in native['source_sha256'].items():
        assert sha(ROOT / relative) == expected, 'Rerun native API audit: ' + relative
    archive = ROOT / 'Drilling_Reviewed.zip'
    files = sorted(p for p in ROOT.rglob('*') if p.is_file()
                   and '__pycache__' not in p.parts
                   and not {'obj', 'bin', 'integration'}.intersection(p.relative_to(ROOT).parts)
                   and p.relative_to(ROOT).as_posix() not in ('checks/integration.json', 'tools/check_integration.py')
                   and p not in (archive, ROOT / 'SHA256SUMS.txt'))
    assert not any(p.suffix.lower() in ('.dll', '.pdb', '.exe', '.cmd', '.vbs', '.ps1') for p in files), 'Native VBA release must not include external programs/installers'
    sums = ROOT / 'SHA256SUMS.txt'
    sums.write_text(''.join(f'{sha(p)}  {p.relative_to(ROOT).as_posix()}\n' for p in files), encoding='utf-8')
    files.append(sums)
    with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED) as z:
        for path in files:
            info = zipfile.ZipInfo(path.relative_to(ROOT).as_posix(), (2026, 10, 29, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, path.read_bytes())
    with zipfile.ZipFile(archive) as z:
        assert z.testzip() is None
        for path in files:
            assert z.read(path.relative_to(ROOT).as_posix()) == path.read_bytes()
    print(f'Validated {archive.name}: {sha(archive)}')


if __name__ == '__main__':
    main()
