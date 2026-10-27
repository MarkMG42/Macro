"""Build and validate the current package using Python's standard library."""
from pathlib import Path
import hashlib
import json
import re
import zipfile
import build_release
ROOT = Path(__file__).resolve().parents[1]
def main():
    build_release.main()
    txt = (ROOT / 'MERGE_DRAWINGS_V16_11.txt').read_text(encoding='utf-8-sig')
    bas = (ROOT / 'MergeDrawingsV16_11.bas').read_text(encoding='cp1251')
    assert bas == 'Attribute VB_Name = "MergeDrawingsV16_11"\n' + txt
    assert max(map(len, txt.splitlines())) < 1023
    current = maximum = 0
    for line in txt.splitlines():
        current = current + 1 if line.rstrip().endswith(' _') else 0
        maximum = max(current, maximum)
    assert maximum <= 24
    names = re.findall(r'^(?:Public|Private) (?:Sub|Function) (\w+)', txt, re.M)
    assert len(names) == len(set(n.lower() for n in names))
    forbidden = r'mshta|ActiveXObject|\bShell\s*\(|WScript|PowerShell|cmd\.exe|XMLHTTP|WinHTTP|URLDownload|<script|\.hta\b|\.(?:Run|Exec)\s*\('
    form = (ROOT / 'FORM_CODE.txt').read_text(encoding='utf-8')
    assert not re.search(forbidden, txt + form, re.I), 'Unexpected script/process facility'
    (ROOT / 'checks').mkdir(exist_ok=True)
    (ROOT / 'checks/static.json').write_text(json.dumps(dict(bas_txt_equivalent=True,
        procedures=len(names), max_continuations=maximum, excluded_script_process_patterns=True,
        native_vba_compile=False, native_windows_ui_run=False, userform_source_included=True, native_solidworks_run=False,
        antivirus_scan_performed=False), indent=2), encoding='utf-8')
    files = sorted(p for p in ROOT.rglob('*') if p.is_file() and '__pycache__' not in p.parts
                   and p.name != 'SHA256SUMS.txt')
    sums = ROOT / 'SHA256SUMS.txt'
    sums.write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest() + '  ' + p.relative_to(ROOT).as_posix() + '\n'
                            for p in files), encoding='utf-8')
    files.append(sums)
    archive = ROOT.parent / 'SOLIDWORKS_MERGE_V16_11_VBA_FORM.zip'
    with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED) as z:
        for p in files:
            z.write(p, ROOT.name + '/' + p.relative_to(ROOT).as_posix())
    with zipfile.ZipFile(archive) as z:
        assert z.testzip() is None
        for p in files:
            assert z.read(ROOT.name + '/' + p.relative_to(ROOT).as_posix()) == p.read_bytes()
    print('Validated package:', archive.name)
    print('SHA256:', hashlib.sha256(archive.read_bytes()).hexdigest())
if __name__ == '__main__':
    main()
