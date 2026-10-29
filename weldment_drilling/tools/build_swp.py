"""Build a SOURCE-ONLY SWP, never reusing old compiled VBA caches.

Requires oletools==0.60.2, olefile==0.47, Java 17+, and Apache POI jars.
Usage: python3 tools/build_swp.py original.swp /path/to/poi-jars
This validates the OLE/source structure; it does NOT execute SOLIDWORKS.
Source-only cache reset technique: Mandiant OfficePurge, commit
0e838faa8e3a887ecf3c74e810ce62acd230f53d. Its use with SOLIDWORKS still needs
native load/Compile/save validation. No claim of native compilation is made.
"""
from pathlib import Path
import hashlib
import json
import re
import struct
import subprocess
import sys
import tempfile

import olefile
from oletools.olevba import VBA_Project, decompress_stream

ROOT = Path(__file__).resolve().parents[1]
VBA_ROOT = 'apc/The VBA Project/_VBA_Project/'


def digest(data):
    return hashlib.sha256(data).hexdigest()


def compress_source(data):
    """MS-OVBA literal compressed chunks, with no copied-byte tokens."""
    result = bytearray(b'\x01')
    for offset in range(0, len(data), 3000):
        chunk = data[offset:offset + 3000]
        payload = b''.join(b'\x00' + chunk[i:i + 8] for i in range(0, len(chunk), 8))
        result.extend(struct.pack('<H', 0xB000 | (len(payload) - 1)))
        result.extend(payload)
    assert decompress_stream(bytes(result)) == data
    return bytes(result)


def project(ole):
    p = VBA_Project(ole, VBA_ROOT, VBA_ROOT + 'PROJECT', VBA_ROOT + 'VBA/dir', relaxed=False)
    p.parse_project_stream()
    list(p.parse_modules())
    return p


def main():
    original = Path(sys.argv[1]).resolve()
    jars = Path(sys.argv[2]).resolve()
    original_bytes = original.read_bytes()
    expected_original = 'ed7d9dc23db5bead52dc6208b314897fd7c96cf735f4ce82b0f945101c82363e'
    assert digest(original_bytes) == expected_original, 'Unexpected input SWP: inspect it before rebuilding'
    output = ROOT / 'Сверловка_исправлено.swp'
    with olefile.OleFileIO(str(original)) as ole:
        before = {'/'.join(path): ole.openstream(path).read() for path in ole.listdir()}
        storage_clsids = {'/'.join(p): ole.getclsid(p) for p in ole.listdir(streams=False, storages=True)}
        root_clsid = ole.root.clsid
        parsed = project(ole)
        assert parsed.codepage == 1251
        modules = {m.streamname: m.code_raw for m in parsed.modules}
        assert set(modules) == {'ThisLibrary', 'PipeCount', 'RotateSelectedDrawingView',
                                'InsertWeldWithHoles', 'SetSizeWeldment', 'Helpers'}
        assert not any('/__SRP_' in name for name in before), 'Additional performance caches require explicit removal'
        directory = bytearray(decompress_stream(before[VBA_ROOT + 'VBA/dir']))
        marker = struct.pack('<HI', 0x31, 4)
        offsets = list(re.finditer(re.escape(marker), directory))
        assert len(offsets) == len(parsed.modules)
        for match, module in zip(offsets, parsed.modules):
            position = match.end()
            assert struct.unpack_from('<I', directory, position)[0] == module.textoffset
            struct.pack_into('<I', directory, position, 0)
    modules['InsertWeldWithHoles'] = (ROOT / 'InsertWeldWithHoles.bas').read_bytes()
    replacements = {VBA_ROOT + 'VBA/' + name: compress_source(code) for name, code in modules.items()}
    replacements[VBA_ROOT + 'VBA/dir'] = compress_source(bytes(directory))
    replacements[VBA_ROOT + 'VBA/_VBA_PROJECT'] = bytes.fromhex('cc61ffff000000')
    java_source = ROOT / 'tools/ReplaceOleStreams.java'
    classpath = str(jars / '*')
    with tempfile.TemporaryDirectory(prefix='swp-rebuild-') as staging:
        staging = Path(staging)
        streams = staging / 'streams'
        for name, content in replacements.items():
            path = streams / name
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(content)
        subprocess.run(['java', '--class-path', classpath, str(java_source),
                        str(original), str(output), str(streams)], check=True)
    # Independent OLE reader and VBA parser must recover exactly the expected project.
    with olefile.OleFileIO(str(output), raise_defects=olefile.DEFECT_INCORRECT) as ole:
        after = {'/'.join(path): ole.openstream(path).read() for path in ole.listdir()}
        assert set(before) == set(after), 'Stream membership changed'
        assert root_clsid == ole.root.clsid
        assert storage_clsids == {'/'.join(p): ole.getclsid(p) for p in ole.listdir(streams=False, storages=True)}
        for name in before:
            assert after[name] == replacements.get(name, before[name]), name
        recovered = project(ole)
        assert recovered.codepage == 1251
        assert len(recovered.modules) == len(modules)
        for module in recovered.modules:
            assert module.textoffset == 0, module.name
            assert module.code_raw == modules[module.streamname], module.name
    assert original.read_bytes() == original_bytes
    report = dict(original_sha256=digest(original_bytes), output_sha256=digest(output.read_bytes()),
                  source_sha256=digest((ROOT / 'InsertWeldWithHoles.txt').read_bytes()),
                  import_module_sha256=digest(modules['InsertWeldWithHoles']),
                  technique='source-only VBA; all module performance caches removed; project cache reset',
                  module_count=len(modules), changed_source_modules=['InsertWeldWithHoles'],
                  all_other_module_sources_byte_identical=True, other_ole_streams_byte_identical=True,
                  all_module_source_offsets_zero=True, container_reopened_with_olefile=True,
                  all_sources_reextracted_with_oletools=True, native_vba_compile=False,
                  native_solidworks_run=False, native_solidworks_load=False,
                  binary_status='experimental source-only SWP; requires native load/Compile/save',
                  technique_reference='https://github.com/mandiant/OfficePurge/tree/0e838faa8e3a887ecf3c74e810ce62acd230f53d')
    (ROOT / 'checks/swp.json').write_text(json.dumps(report, indent=2) + '\n')
    print(f'Validated source-only container: {output.name}, SHA256 {report["output_sha256"]}')


if __name__ == '__main__':
    main()
